"""Pengunduhan dan normalisasi gambar/PDF untuk dokumen Word.

Modul ini menjadi satu pintu untuk seluruh sumber bukti dukung. URL Google
Drive tetap memakai alur konfirmasi/download yang sudah ada, sedangkan URL
HTTP(S) lain diambil sebagai resource web biasa. Tipe akhir resource selalu
divalidasi dari byte hasil unduhan, bukan dari ekstensi URL atau Content-Type
semata.
"""

from __future__ import annotations

import io
import re
import time
from urllib.parse import parse_qs, urlparse

import requests

try:
    import pymupdf as fitz

    HAS_PDF_RENDERER = True
except ImportError:  # pragma: no cover - dependency wajib pada paket aplikasi
    HAS_PDF_RENDERER = False

try:
    from PIL import Image, ImageFile, ImageOps, UnidentifiedImageError

    ImageFile.LOAD_TRUNCATED_IMAGES = True
    HAS_PIL = True
except ImportError:  # pragma: no cover - dependency wajib pada paket aplikasi
    HAS_PIL = False

HEIF_IMPORT_ERROR = None

try:
    from pillow_heif import open_heif, register_heif_opener

    register_heif_opener()
    HAS_HEIF = True
except (ImportError, OSError) as exc:  # OSError: codec native gagal dimuat
    HAS_HEIF = False
    HEIF_IMPORT_ERROR = exc


_DRIVE_DOWNLOAD_URL = "https://drive.google.com/uc"
_DRIVE_DIRECT_URL = "https://drive.usercontent.google.com/download"
MAX_WEB_IMAGE_BYTES = 25 * 1024 * 1024
DEFAULT_WEB_MAX_RETRIES = 3
DEFAULT_WEB_RETRY_DELAY = 0.5
_RETRYABLE_STATUS_CODES = frozenset({429, 500, 502, 503, 504})
_STREAM_CHUNK_SIZE = 64 * 1024
_DRIVE_FILE_ID_RE = re.compile(r"[A-Za-z0-9_-]+\Z")
_BARE_DRIVE_FILE_ID_RE = re.compile(r"[A-Za-z0-9_-]{10,}\Z")
_GOOGLE_DRIVE_HOSTS = frozenset({
    "drive.google.com",
    "www.drive.google.com",
    "docs.google.com",
    "www.docs.google.com",
    "drive.usercontent.google.com",
})


class SourceDownloadError(RuntimeError):
    """Structured failure raised while retrieving an evidence source.

    ``status_code`` is deliberately kept separate from the rendered message
    so callers never need to infer HTTP status from arbitrary URL text.
    ``kind`` is a small classification for callers that need to distinguish
    HTTP, timeout, and connection failures without parsing strings.
    """

    def __init__(
        self,
        message: str,
        *,
        source: str | None = None,
        status_code: int | None = None,
        kind: str | None = None,
    ):
        self.source = source
        self.status_code = status_code
        self.kind = kind
        super().__init__(message)


class _DownloadHTTPError(RuntimeError):
    """HTTP failure that retains the status code for retry decisions."""

    def __init__(self, status_code: int, url: str):
        self.status_code = status_code
        self.url = url
        super().__init__(f"HTTP {status_code}: {url}")


def extract_drive_file_id(link: str) -> str | None:
    """Extract a bare ID or ID from a supported public Google Drive URL.

    Generic HTTP(S) URLs are intentionally not inspected for an ``id`` query
    parameter. This keeps URLs such as ``https://example.com/download?id=123``
    on the generic HTTP(S) path.
    """
    value = str(link or "").strip()
    if _is_bare_drive_file_id(value):
        return value
    if not _is_google_drive_url(value):
        return None

    parsed = urlparse(value)
    match = re.search(r"/d/([^/?#]+)", parsed.path)
    if match and _is_drive_file_id(match.group(1)):
        return match.group(1)

    query_id = parse_qs(parsed.query).get("id", [None])[0]
    if query_id and _is_drive_file_id(query_id):
        return query_id
    return None


def _is_drive_file_id(value: str) -> bool:
    return bool(_DRIVE_FILE_ID_RE.fullmatch(str(value or "").strip()))


def _is_bare_drive_file_id(value: str) -> bool:
    return bool(_BARE_DRIVE_FILE_ID_RE.fullmatch(str(value or "").strip()))


def _is_google_drive_host(hostname: str | None) -> bool:
    """Return whether *hostname* is a supported Google-owned Drive host."""
    normalized = (hostname or "").lower().rstrip(".")
    return normalized in _GOOGLE_DRIVE_HOSTS or normalized.endswith(
        ".drive.google.com"
    )


def _is_google_drive_source(value: str) -> bool:
    """Return whether a source is a bare Drive ID or supported Drive URL."""
    value = str(value or "").strip()
    if _is_bare_drive_file_id(value):
        return True
    parsed = urlparse(value)
    return (
        parsed.scheme.lower() in {"http", "https"}
        and bool(parsed.netloc)
        and _is_google_drive_host(parsed.hostname)
    )


def _is_google_drive_url(value: str) -> bool:
    """Compatibility wrapper for supported Google Drive URLs only."""
    value = str(value or "").strip()
    return not _is_bare_drive_file_id(value) and _is_google_drive_source(value)


def _validate_http_url(value: str):
    """Parse and validate a user-supplied HTTP(S) URL."""
    parsed = urlparse(value)
    if parsed.scheme.lower() not in {"http", "https"} or not parsed.netloc:
        raise ValueError("Nilai bukan tautan HTTP(S) yang valid.")
    return parsed


def _looks_like_html(raw_bytes: bytes, content_type: str = "") -> bool:
    """Detect an HTML page returned instead of a downloadable resource."""
    prefix = raw_bytes[:512].lstrip().lower()
    return prefix.startswith(
        (b"<!doctype html", b"<html", b"<head", b"<body")
    )


def _content_type_is_html(content_type: str = "") -> bool:
    normalized_type = (content_type or "").lower()
    return "text/html" in normalized_type or "application/xhtml" in normalized_type


def _html_error(source_label: str | None = None) -> RuntimeError:
    if source_label == "Google Drive":
        return RuntimeError(
            "Google Drive mengembalikan halaman HTML, bukan file. "
            "Pastikan akses file disetel untuk siapa saja yang memiliki tautan."
        )
    return RuntimeError("Server mengembalikan halaman HTML, bukan file gambar.")


def image_bytes_to_png(
    raw_bytes: bytes,
    content_type: str = "",
    *,
    source_label: str | None = None,
):
    """Decode gambar umum/HEIC lalu kembalikan ``(BytesIO PNG, PIL.Image)``.

    Validasi dilakukan dengan membuka byte memakai Pillow. Karena itu URL tanpa
    ekstensi, Content-Type ``application/octet-stream``, maupun nama file yang
    menyesatkan tidak mengubah hasil validasi. EXIF orientation diterapkan dan
    hasil dikonversi ke mode yang aman untuk PNG/python-docx.
    """
    if not HAS_PIL:
        raise RuntimeError("Pillow tidak terinstal.")
    if not raw_bytes:
        raise RuntimeError("File gambar kosong.")
    if _looks_like_html(raw_bytes, content_type):
        raise _html_error(source_label)

    try:
        with Image.open(io.BytesIO(raw_bytes)) as source:
            source.load()
            image = ImageOps.exif_transpose(source)
            image.load()
    except (UnidentifiedImageError, OSError, ValueError) as exc:
        if _content_type_is_html(content_type):
            raise _html_error(source_label) from exc
        # Fallback langsung membuat dukungan HEIC tidak hanya bergantung pada
        # registrasi plugin Pillow dan memastikan freezer mendeteksi API native.
        if HAS_HEIF:
            try:
                heif_file = open_heif(raw_bytes, convert_hdr_to_8bit=True)
                image = heif_file.to_pillow()
                image = ImageOps.exif_transpose(image)
                image.load()
            except Exception as heif_exc:
                raise RuntimeError(
                    "Resource berhasil diunduh tetapi bukan gambar yang "
                    "didukung. File mungkin rusak atau formatnya tidak "
                    "didukung."
                ) from heif_exc
        else:
            heif_signature = raw_bytes[4:12].lower()
            detail = (
                " Decoder HEIC/HEIF tidak tersedia."
                if heif_signature in {
                    b"ftypheic",
                    b"ftypheix",
                    b"ftyphevc",
                    b"ftypmif1",
                }
                else ""
            )
            raise RuntimeError(
                "Resource berhasil diunduh tetapi bukan gambar yang "
                f"didukung.{detail}"
            ) from exc

    if image.mode not in ("RGB", "L"):
        original = image
        image = image.convert("RGB")
        original.close()

    png_file = io.BytesIO()
    image.save(png_file, format="PNG")
    png_file.seek(0)
    return png_file, image


def pdf_bytes_to_png_pages(raw_bytes: bytes):
    """Render seluruh halaman PDF menjadi daftar ``(BytesIO PNG, ukuran)``."""
    if not HAS_PDF_RENDERER:
        raise RuntimeError("PyMuPDF tidak terinstal - file PDF tidak dapat diproses.")
    if not raw_bytes:
        raise RuntimeError("File PDF kosong.")

    pages = []
    try:
        with fitz.open(stream=raw_bytes, filetype="pdf") as pdf:
            if pdf.page_count == 0:
                raise RuntimeError("File PDF tidak memiliki halaman.")
            for page in pdf:
                # 144 dpi cukup tajam untuk DOCX tanpa membuat hasil terlalu besar.
                pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
                png_file = io.BytesIO(pixmap.tobytes("png"))
                png_file.seek(0)
                pages.append((png_file, (pixmap.width, pixmap.height)))
    except RuntimeError:
        raise
    except Exception as exc:
        raise RuntimeError(f"File PDF tidak dapat dibaca: {exc}") from exc
    return pages


def _response_headers(response) -> dict:
    headers = getattr(response, "headers", None)
    return headers if headers is not None else {}


def _response_status_code(response) -> int | None:
    value = getattr(response, "status_code", None)
    return value if isinstance(value, int) else None


def _check_response(response, url: str) -> None:
    status_code = _response_status_code(response)
    if status_code is not None and status_code >= 400:
        raise _DownloadHTTPError(status_code, url)
    raise_for_status = getattr(response, "raise_for_status", None)
    if callable(raise_for_status):
        try:
            raise_for_status()
        except requests.HTTPError as exc:
            status_code = _response_status_code(response) or 0
            raise _DownloadHTTPError(status_code, url) from exc


def _size_error(max_bytes: int) -> RuntimeError:
    return RuntimeError(
        "Ukuran unduhan melebihi batas "
        f"{max_bytes // (1024 * 1024)} MB."
    )


def _content_length(headers: dict) -> int | None:
    value = headers.get("Content-Length")
    try:
        parsed = int(value)
    except (TypeError, ValueError):
        return None
    return parsed if parsed >= 0 else None


def _read_response_bytes(response, max_bytes: int) -> bytes:
    """Read a streamed response into bounded memory and return its bytes."""
    headers = _response_headers(response)
    declared_size = _content_length(headers)
    if declared_size is not None and declared_size > max_bytes:
        raise _size_error(max_bytes)

    buffer = io.BytesIO()
    iterator = getattr(response, "iter_content", None)
    if callable(iterator):
        try:
            chunks = iterator(chunk_size=_STREAM_CHUNK_SIZE)
            for chunk in chunks:
                if not chunk:
                    continue
                total = buffer.tell() + len(chunk)
                if total > max_bytes:
                    raise _size_error(max_bytes)
                buffer.write(chunk)
        except TypeError:
            # Keep compatibility with very small test doubles that expose only
            # ``content`` or a non-iterable ``iter_content`` attribute.
            raw = getattr(response, "content", b"")
            if not isinstance(raw, (bytes, bytearray, memoryview)):
                raise RuntimeError("Respons HTTP tidak dapat dibaca.")
            if len(raw) > max_bytes:
                raise _size_error(max_bytes)
            return bytes(raw)
    else:
        raw = getattr(response, "content", b"")
        if not isinstance(raw, (bytes, bytearray, memoryview)):
            raise RuntimeError("Respons HTTP tidak dapat dibaca.")
        if len(raw) > max_bytes:
            raise _size_error(max_bytes)
        return bytes(raw)

    return buffer.getvalue()


def _http_error_message(error: _DownloadHTTPError, attempts: int) -> str:
    status_code = error.status_code
    if status_code == 400:
        return f"Permintaan tidak valid (400): {error.url}"
    if status_code == 401:
        return f"Autentikasi diperlukan (401): {error.url}"
    if status_code == 403:
        return f"Akses ditolak (403): {error.url}"
    if status_code == 404:
        return f"File tidak ditemukan (404): {error.url}"
    if status_code in _RETRYABLE_STATUS_CODES:
        return (
            f"Server sementara tidak tersedia ({status_code}) setelah "
            f"{attempts} percobaan: {error.url}"
        )
    return f"Gagal mengunduh resource ({status_code}): {error.url}"


def _wait_before_retry(attempt: int, max_retries: int, retry_delay: float) -> None:
    if attempt < max_retries and retry_delay > 0:
        time.sleep(retry_delay)


def _download_drive_bytes(
    file_id: str,
    max_retries: int,
    retry_delay: float,
    timeout: float,
    max_bytes: int = MAX_WEB_IMAGE_BYTES,
):
    """Unduh byte file publik Google Drive beserta Content-Type respons."""
    session = requests.Session()
    drive_url = f"{_DRIVE_DOWNLOAD_URL}?export=download&id={file_id}"
    try:
        for attempt in range(1, max_retries + 1):
            response = None
            try:
                response = session.get(
                    _DRIVE_DOWNLOAD_URL,
                    params={"export": "download", "id": file_id},
                    stream=True,
                    timeout=timeout,
                    verify=True,
                    allow_redirects=True,
                )
                _check_response(response, drive_url)
                raw_bytes = _read_response_bytes(response, max_bytes)
                content_type = _response_headers(response).get("Content-Type", "")

                if _looks_like_html(raw_bytes, content_type):
                    cookies = getattr(response, "cookies", {}) or {}
                    token = next(
                        (
                            value
                            for key, value in cookies.items()
                            if key.startswith("download_warning")
                        ),
                        None,
                    )
                    response.close()
                    response = session.get(
                        _DRIVE_DOWNLOAD_URL if token else _DRIVE_DIRECT_URL,
                        params=(
                            {"export": "download", "id": file_id, "confirm": token}
                            if token
                            else {"id": file_id, "export": "download", "confirm": "t"}
                        ),
                        stream=True,
                        timeout=timeout,
                        verify=True,
                        allow_redirects=True,
                    )
                    _check_response(response, drive_url)
                    raw_bytes = _read_response_bytes(response, max_bytes)
                    content_type = _response_headers(response).get("Content-Type", "")

                if not raw_bytes:
                    raise RuntimeError("Google Drive mengembalikan respons kosong.")
                if _looks_like_html(raw_bytes, content_type):
                    raise _html_error("Google Drive")
                return raw_bytes, content_type
            except _DownloadHTTPError as exc:
                if exc.status_code in _RETRYABLE_STATUS_CODES and attempt < max_retries:
                    _wait_before_retry(attempt, max_retries, retry_delay)
                    continue
                raise SourceDownloadError(
                    _http_error_message(exc, attempt),
                    source=exc.url,
                    status_code=exc.status_code,
                    kind="http",
                ) from exc
            except requests.Timeout as exc:
                if attempt < max_retries:
                    _wait_before_retry(attempt, max_retries, retry_delay)
                    continue
                raise SourceDownloadError(
                    f"Waktu tunggu habis saat mengambil gambar Google Drive "
                    f"{file_id}.",
                    source=file_id,
                    kind="timeout",
                ) from exc
            except requests.ConnectionError as exc:
                if attempt < max_retries:
                    _wait_before_retry(attempt, max_retries, retry_delay)
                    continue
                raise SourceDownloadError(
                    f"Koneksi gagal saat mengambil gambar Google Drive {file_id}.",
                    source=file_id,
                    kind="connection",
                ) from exc
            except requests.RequestException as exc:
                if attempt < max_retries:
                    _wait_before_retry(attempt, max_retries, retry_delay)
                    continue
                raise SourceDownloadError(
                    f"Gagal mengunduh file Google Drive {file_id}: {exc}",
                    source=file_id,
                    kind="network",
                ) from exc
            finally:
                if response is not None:
                    response.close()
    finally:
        session.close()

    raise RuntimeError(f"Gagal mengunduh file Google Drive {file_id}.")


def download_drive_image(
    file_id: str,
    *,
    max_retries: int = DEFAULT_WEB_MAX_RETRIES,
    retry_delay: float = 2,
    timeout: float = 30,
    max_bytes: int = MAX_WEB_IMAGE_BYTES,
):
    """Unduh gambar publik Google Drive dan normalisasi hasilnya menjadi PNG."""
    if not HAS_PIL:
        raise RuntimeError("Pillow tidak terinstal.")

    raw_bytes, content_type = _download_drive_bytes(
        file_id, max_retries, retry_delay, timeout, max_bytes
    )
    return image_bytes_to_png(
        raw_bytes, content_type, source_label="Google Drive"
    )


def _download_url_bytes(
    url: str,
    timeout: float,
    max_bytes: int,
    max_retries: int = DEFAULT_WEB_MAX_RETRIES,
    retry_delay: float = DEFAULT_WEB_RETRY_DELAY,
):
    """Download a generic HTTP(S) resource with retries and a size limit."""
    if max_bytes <= 0:
        raise ValueError("Batas ukuran unduhan harus lebih besar dari nol.")
    if max_retries <= 0:
        raise ValueError("Jumlah percobaan unduhan harus lebih besar dari nol.")

    for attempt in range(1, max_retries + 1):
        response = None
        try:
            response = requests.get(
                url,
                stream=True,
                timeout=timeout,
                verify=True,
                allow_redirects=True,
            )
            _check_response(response, url)
            raw_bytes = _read_response_bytes(response, max_bytes)
            if not raw_bytes:
                raise RuntimeError(f"Server mengembalikan respons kosong: {url}")
            return raw_bytes, _response_headers(response).get("Content-Type", "")
        except _DownloadHTTPError as exc:
            if exc.status_code in _RETRYABLE_STATUS_CODES and attempt < max_retries:
                _wait_before_retry(attempt, max_retries, retry_delay)
                continue
            raise SourceDownloadError(
                _http_error_message(exc, attempt),
                source=exc.url,
                status_code=exc.status_code,
                kind="http",
            ) from exc
        except requests.Timeout as exc:
            if attempt < max_retries:
                _wait_before_retry(attempt, max_retries, retry_delay)
                continue
            raise SourceDownloadError(
                f"Waktu tunggu habis saat mengambil gambar: {url}",
                source=url,
                kind="timeout",
            ) from exc
        except requests.ConnectionError as exc:
            if attempt < max_retries:
                _wait_before_retry(attempt, max_retries, retry_delay)
                continue
            raise SourceDownloadError(
                f"Koneksi gagal saat mengambil gambar: {url}",
                source=url,
                kind="connection",
            ) from exc
        except requests.exceptions.ChunkedEncodingError as exc:
            if attempt < max_retries:
                _wait_before_retry(attempt, max_retries, retry_delay)
                continue
            raise SourceDownloadError(
                f"Gagal mengunduh resource: {url}",
                source=url,
                kind="network",
            ) from exc
        except requests.RequestException as exc:
            raise SourceDownloadError(
                f"Gagal mengunduh resource: {url}",
                source=url,
                kind="network",
            ) from exc
        finally:
            if response is not None:
                response.close()

    raise RuntimeError(f"Gagal mengunduh resource: {url}")


def _evidence_items_from_bytes(
    raw_bytes: bytes,
    content_type: str = "",
    *,
    source_label: str | None = None,
):
    """Decode a downloaded resource into image/PDF evidence items."""
    if _looks_like_html(raw_bytes, content_type):
        raise _html_error(source_label)

    # PDF magic bytes are authoritative. Content-Type is only a fallback after
    # trying Pillow, so a valid JPEG mislabeled as application/pdf still works.
    if raw_bytes.lstrip().startswith(b"%PDF-"):
        return [
            ("pdf_page", stream, size)
            for stream, size in pdf_bytes_to_png_pages(raw_bytes)
        ]

    try:
        stream, image = image_bytes_to_png(
            raw_bytes, content_type, source_label=source_label
        )
    except RuntimeError as image_error:
        if "application/pdf" not in (content_type or "").lower():
            raise
        try:
            return [
                ("pdf_page", stream, size)
                for stream, size in pdf_bytes_to_png_pages(raw_bytes)
            ]
        except RuntimeError:
            raise image_error

    try:
        size = image.size
    finally:
        image.close()
    return [("image", stream, size)]


def _resolve_source(source: str) -> tuple[str, str]:
    """Resolve a source into ``("drive", file_id)`` or ``("http", url)``."""
    value = str(source or "").strip()
    file_id = extract_drive_file_id(value)
    if file_id:
        return "drive", file_id
    if _is_google_drive_url(value):
        raise RuntimeError("Tautan Google Drive tidak memiliki file ID.")
    _validate_http_url(value)
    return "http", value


def download_image_source(
    source: str,
    *,
    timeout: float = 15,
    max_bytes: int = MAX_WEB_IMAGE_BYTES,
    max_retries: int = DEFAULT_WEB_MAX_RETRIES,
    retry_delay: float = DEFAULT_WEB_RETRY_DELAY,
):
    """Download one image source and return a normalized PNG image.

    ``source`` may be a bare Google Drive file ID, a supported Google Drive
    URL, or any other HTTP(S) URL. Drive sources use the existing
    confirmation-aware downloader; other URLs are streamed with a size limit
    and bounded retries. The final image type is determined by Pillow, not by
    the URL extension or header.
    """
    source_kind, value = _resolve_source(source)
    if source_kind == "drive":
        return download_drive_image(
            value,
            timeout=timeout,
            max_retries=max_retries,
            retry_delay=retry_delay,
            max_bytes=max_bytes,
        )

    raw_bytes, content_type = _download_url_bytes(
        value, timeout, max_bytes, max_retries, retry_delay
    )
    return image_bytes_to_png(raw_bytes, content_type)


def download_evidence_source(
    source: str,
    *,
    timeout: float = 15,
    max_bytes: int = MAX_WEB_IMAGE_BYTES,
    max_retries: int = DEFAULT_WEB_MAX_RETRIES,
    retry_delay: float = DEFAULT_WEB_RETRY_DELAY,
):
    """Download one image/PDF source as evidence layout items.

    The resolver accepts the same source forms as :func:`download_image_source`
    and keeps Google Drive confirmation handling in the shared module.
    """
    source_kind, value = _resolve_source(source)
    if source_kind == "drive":
        return download_drive_evidence(
            value,
            timeout=timeout,
            max_retries=max_retries,
            retry_delay=retry_delay,
            max_bytes=max_bytes,
        )

    raw_bytes, content_type = _download_url_bytes(
        value, timeout, max_bytes, max_retries, retry_delay
    )
    return _evidence_items_from_bytes(raw_bytes, content_type)


def download_drive_evidence(
    file_id: str,
    *,
    max_retries: int = DEFAULT_WEB_MAX_RETRIES,
    retry_delay: float = 2,
    timeout: float = 30,
    max_bytes: int = MAX_WEB_IMAGE_BYTES,
):
    """Unduh bukti Drive sebagai gambar atau seluruh halaman PDF.

    Return value berupa daftar ``(kind, stream, (width, height))``. ``kind``
    bernilai ``image`` atau ``pdf_page`` sehingga pemanggil dapat memaksa setiap
    halaman PDF ke halaman Word tersendiri.
    """
    if not HAS_PIL:
        raise RuntimeError("Pillow tidak terinstal.")
    raw_bytes, content_type = _download_drive_bytes(
        file_id, max_retries, retry_delay, timeout, max_bytes
    )
    return _evidence_items_from_bytes(
        raw_bytes, content_type, source_label="Google Drive"
    )


def download_url_image(
    url: str,
    *,
    timeout: float = 15,
    max_bytes: int = MAX_WEB_IMAGE_BYTES,
    max_retries: int = DEFAULT_WEB_MAX_RETRIES,
    retry_delay: float = DEFAULT_WEB_RETRY_DELAY,
):
    """Backward-compatible wrapper for :func:`download_image_source`."""
    return download_image_source(
        url,
        timeout=timeout,
        max_bytes=max_bytes,
        max_retries=max_retries,
        retry_delay=retry_delay,
    )


def download_url_evidence(
    url: str,
    *,
    timeout: float = 15,
    max_bytes: int = MAX_WEB_IMAGE_BYTES,
    max_retries: int = DEFAULT_WEB_MAX_RETRIES,
    retry_delay: float = DEFAULT_WEB_RETRY_DELAY,
):
    """Backward-compatible wrapper for :func:`download_evidence_source`."""
    return download_evidence_source(
        url,
        timeout=timeout,
        max_bytes=max_bytes,
        max_retries=max_retries,
        retry_delay=retry_delay,
    )
