"""Integration-level checks that generators pass complete source URLs onward."""

import io
import unittest
from unittest.mock import patch

import pandas as pd
from docx import Document
from PIL import Image

from src import bast, bapp_pml, bapp_pml_t2, bapp_ppl, bapp_ppl_t2, bukti_terima
from src.document_generator import insert_custom_url_images
from utils import evidence
from utils.images import SourceDownloadError


REMOTE_URL = (
    "https://progres-mitra.bpskaro.site/"
    "api/dokumen-mitra/62/ss-fasih"
)
BARE_DRIVE_ID = "1AbCdEfGhIjKlMnOpQrStUvWxYz123"


def _image_result(size=(120, 80)):
    image = Image.new("RGB", size, color="navy")
    stream = io.BytesIO()
    image.save(stream, format="PNG")
    stream.seek(0)
    return stream, image


def _evidence_result(_source):
    stream, image = _image_result()
    size = image.size
    image.close()
    return [("image", stream, size)]


def _document_with_placeholder(placeholder):
    document = Document()
    paragraph = document.add_paragraph(placeholder)
    return document, paragraph


class BuiltinRemoteEvidenceTests(unittest.TestCase):
    def test_bapp_termin1_ppl_and_pml_accept_extensionless_source_url(self):
        for module in (bapp_ppl, bapp_pml):
            with self.subTest(module=module.__name__):
                document, paragraph = _document_with_placeholder(
                    module.BUKTI_PLACEHOLDER
                )
                with patch.object(
                    module,
                    "_download_drive_image",
                    side_effect=lambda _source: _image_result(),
                ) as downloader:
                    count, warnings = module.insert_evidence_images(
                        document, REMOTE_URL
                    )

                self.assertEqual(count, 1)
                self.assertEqual(warnings, [])
                downloader.assert_called_once_with(REMOTE_URL)
                self.assertEqual(len(document.inline_shapes), 1)
                self.assertEqual(paragraph.text, "")

    def test_bapp_termin1_bare_drive_id_uses_legacy_drive_downloader(self):
        for module in (bapp_ppl, bapp_pml):
            with self.subTest(module=module.__name__):
                document, paragraph = _document_with_placeholder(
                    module.BUKTI_PLACEHOLDER
                )
                with patch.object(
                    module,
                    "_download_drive_image",
                    side_effect=lambda _source: _image_result(),
                ) as downloader:
                    count, warnings = module.insert_evidence_images(
                        document, BARE_DRIVE_ID
                    )

                self.assertEqual(count, 1)
                self.assertEqual(warnings, [])
                downloader.assert_called_once_with(BARE_DRIVE_ID)
                self.assertEqual(len(document.inline_shapes), 1)
                self.assertEqual(paragraph.text, "")

    def test_bapp_termin1_preserves_commas_inside_source_urls(self):
        url_with_comma = "https://example.test/api/evidence,latest"
        second_url = "https://example.test/api/evidence-2"
        for module in (bapp_ppl, bapp_pml):
            with self.subTest(module=module.__name__):
                document, paragraph = _document_with_placeholder(
                    module.BUKTI_PLACEHOLDER
                )
                with patch.object(
                    module,
                    "_download_drive_image",
                    side_effect=lambda _source: _image_result(),
                ) as downloader:
                    count, warnings = module.insert_evidence_images(
                        document, f"{url_with_comma}, {second_url}"
                    )

                self.assertEqual(count, 2)
                self.assertEqual(warnings, [])
                self.assertEqual(
                    [item.args[0] for item in downloader.call_args_list],
                    [url_with_comma, second_url],
                )
                self.assertEqual(len(document.inline_shapes), 2)
                self.assertEqual(paragraph.text, "")

    def test_bapp_termin2_ppl_and_pml_accept_extensionless_source_url(self):
        for module in (bapp_ppl_t2, bapp_pml_t2):
            with self.subTest(module=module.__name__):
                document, paragraph = _document_with_placeholder(
                    module.BUKTI_PLACEHOLDER
                )
                with patch.object(
                    module, "_download_drive_evidence", side_effect=_evidence_result
                ) as downloader:
                    count, warnings = module.insert_evidence_images(
                        document, REMOTE_URL
                    )

                self.assertEqual(count, 1)
                self.assertEqual(warnings, [])
                downloader.assert_called_once_with(REMOTE_URL)
                self.assertEqual(len(document.inline_shapes), 1)
                self.assertEqual(paragraph.text, "")

    def test_bapp_termin2_bare_drive_id_uses_legacy_drive_downloader(self):
        for module in (bapp_ppl_t2, bapp_pml_t2):
            with self.subTest(module=module.__name__):
                document, paragraph = _document_with_placeholder(
                    module.BUKTI_PLACEHOLDER
                )
                with patch.object(
                    module,
                    "_download_drive_evidence",
                    side_effect=_evidence_result,
                ) as downloader:
                    count, warnings = module.insert_evidence_images(
                        document, BARE_DRIVE_ID
                    )

                self.assertEqual(count, 1)
                self.assertEqual(warnings, [])
                downloader.assert_called_once_with(BARE_DRIVE_ID)
                self.assertEqual(len(document.inline_shapes), 1)
                self.assertEqual(paragraph.text, "")

    def test_bapp_termin2_preserves_commas_inside_source_urls(self):
        url_with_comma = "https://example.test/api/evidence,latest"
        second_url = "https://example.test/api/evidence-2"
        for module in (bapp_ppl_t2, bapp_pml_t2):
            with self.subTest(module=module.__name__):
                document, paragraph = _document_with_placeholder(
                    module.BUKTI_PLACEHOLDER
                )
                with patch.object(
                    module,
                    "_download_drive_evidence",
                    side_effect=_evidence_result,
                ) as downloader:
                    count, warnings = module.insert_evidence_images(
                        document, f"{url_with_comma}, {second_url}"
                    )

                self.assertEqual(count, 2)
                self.assertEqual(warnings, [])
                self.assertEqual(
                    [item.args[0] for item in downloader.call_args_list],
                    [url_with_comma, second_url],
                )
                self.assertEqual(len(document.inline_shapes), 2)
                self.assertEqual(paragraph.text, "")

    def test_bast_reuses_the_termin2_universal_evidence_path(self):
        document, paragraph = _document_with_placeholder(bast.BUKTI_PLACEHOLDER)
        with patch.object(
            bapp_ppl_t2, "_download_drive_evidence", side_effect=_evidence_result
        ) as downloader:
            count, warnings = bast.insert_evidence_images(
                document, REMOTE_URL, placeholder=bast.BUKTI_PLACEHOLDER
            )

        self.assertEqual(count, 1)
        self.assertEqual(warnings, [])
        downloader.assert_called_once_with(REMOTE_URL)
        self.assertEqual(len(document.inline_shapes), 1)
        self.assertEqual(paragraph.text, "")

    def test_bast_bare_drive_id_reaches_drive_evidence_downloader(self):
        document, paragraph = _document_with_placeholder(bast.BUKTI_PLACEHOLDER)
        with patch.object(
            bapp_ppl_t2, "_download_drive_evidence", side_effect=_evidence_result
        ) as downloader:
            count, warnings = bast.insert_evidence_images(
                document, BARE_DRIVE_ID, placeholder=bast.BUKTI_PLACEHOLDER
            )

        self.assertEqual(count, 1)
        self.assertEqual(warnings, [])
        downloader.assert_called_once_with(BARE_DRIVE_ID)
        self.assertEqual(len(document.inline_shapes), 1)
        self.assertEqual(paragraph.text, "")

    def test_bukti_terima_accepts_extensionless_source_url(self):
        document = Document()
        cell = document.add_table(rows=1, cols=1).cell(0, 0)
        with patch.object(
            bukti_terima,
            "_download_drive_image",
            side_effect=lambda _source: _image_result(),
        ) as downloader:
            warnings = bukti_terima._fill_person_cell(
                cell,
                {
                    "nik": "6304000000000001",
                    "nama_lengkap": "PETUGAS UJI",
                    "wilayah_tugas": "Wilayah",
                    "operator": "Operator",
                    "no_telp": "0800000000",
                    "link_bukti_terima": REMOTE_URL,
                },
                0,
            )

        self.assertEqual(warnings, [])
        downloader.assert_called_once_with(REMOTE_URL)
        self.assertEqual(len(cell._tc.xpath(".//w:drawing")), 1)

    def test_bukti_terima_bare_drive_id_uses_legacy_drive_downloader(self):
        document = Document()
        cell = document.add_table(rows=1, cols=1).cell(0, 0)
        with patch.object(
            bukti_terima,
            "_download_drive_image",
            side_effect=lambda _source: _image_result(),
        ) as downloader:
            warnings = bukti_terima._fill_person_cell(
                cell,
                {
                    "nik": "6304000000000001",
                    "nama_lengkap": "PETUGAS UJI",
                    "wilayah_tugas": "Wilayah",
                    "operator": "Operator",
                    "no_telp": "0800000000",
                    "link_bukti_terima": BARE_DRIVE_ID,
                },
                0,
            )

        self.assertEqual(warnings, [])
        downloader.assert_called_once_with(BARE_DRIVE_ID)
        self.assertEqual(len(cell._tc.xpath(".//w:drawing")), 1)

    def test_bukti_terima_pdf_warning_does_not_abort_cell_batch(self):
        document = Document()
        cell = document.add_table(rows=1, cols=1).cell(0, 0)
        with patch.object(
            bukti_terima,
            "_download_drive_image",
            side_effect=RuntimeError(
                "Resource berhasil diunduh tetapi bukan gambar yang didukung."
            ),
        ):
            warnings = bukti_terima._fill_person_cell(
                cell,
                {
                    "nik": "6304000000000001",
                    "nama_lengkap": "PETUGAS UJI",
                    "wilayah_tugas": "Wilayah",
                    "operator": "Operator",
                    "no_telp": "0800000000",
                    "link_bukti_terima": REMOTE_URL + ".pdf",
                },
                0,
            )

        self.assertEqual(len(warnings), 1)
        self.assertIn("Gagal memuat gambar", warnings[0])
        self.assertEqual(len(cell._tc.xpath(".//w:drawing")), 0)

    def test_custom_placeholder_uses_the_same_extensionless_source_contract(self):
        document, paragraph = _document_with_placeholder("Bukti: {{ss_fasih}}")
        row = pd.Series({"ss_fasih": REMOTE_URL})

        with patch(
            "utils.images.download_evidence_source", side_effect=_evidence_result
        ) as downloader:
            consumed = insert_custom_url_images(document, row)

        self.assertEqual(consumed, {"{{ss_fasih}}"})
        downloader.assert_called_once_with(REMOTE_URL)
        self.assertEqual(len(document.inline_shapes), 1)
        self.assertEqual(paragraph.text, "Bukti: ")

    def test_status_digits_in_url_do_not_change_generic_evidence_warning(self):
        document, paragraph = _document_with_placeholder("{{bukti_dukung}}")
        url = "https://example.test/api/404/image"

        def timeout(source):
            raise SourceDownloadError(
                f"Waktu tunggu habis saat mengambil gambar: {source}",
                source=source,
                kind="timeout",
            )

        count, warnings = evidence.insert_evidence(
            document,
            url,
            "{{bukti_dukung}}",
            evidence.IMAGE_LAYOUT_GRID,
            replace_text=lambda doc, replacements: None,
            image_orientation=evidence.IMAGE_ORIENTATION_PORTRAIT,
            source_downloader=timeout,
        )

        self.assertEqual(count, 0)
        self.assertEqual(len(warnings), 1)
        self.assertIn("Gagal memuat", warnings[0])
        self.assertNotIn("File tidak ditemukan", warnings[0])
        self.assertEqual(paragraph.text, "{{bukti_dukung}}")

    def test_status_digits_in_url_do_not_change_bapp_warning(self):
        url = "https://example.test/api/403/image"
        for module in (bapp_ppl, bapp_pml):
            with self.subTest(module=module.__name__):
                document, _ = _document_with_placeholder(
                    module.BUKTI_PLACEHOLDER
                )
                with patch.object(
                    module,
                    "_download_drive_image",
                    side_effect=RuntimeError(
                        f"Waktu tunggu habis saat mengambil gambar: {url}"
                    ),
                ):
                    count, warnings = module.insert_evidence_images(document, url)

                self.assertEqual(count, 0)
                self.assertEqual(len(warnings), 1)
                self.assertIn("Gagal memuat", warnings[0])
                self.assertNotIn("Akses ditolak", warnings[0])


if __name__ == "__main__":
    unittest.main()
