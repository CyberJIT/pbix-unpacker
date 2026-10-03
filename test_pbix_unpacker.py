"""
Unit tests for pbix_unpacker.py
"""

import os
import shutil
import tempfile
import zipfile
import unittest

from pbix_unpacker import (
    find_pbix_files,
    should_include_name,
    unpack_single_pbix,
)


class TestPbixUnpacker(unittest.TestCase):
    def setUp(self):
        self.test_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def _create_pbix(self, rel_path: str, file_dict: dict) -> str:
        full_path = os.path.join(self.test_dir, rel_path)
        os.makedirs(os.path.dirname(full_path), exist_ok=True)
        with zipfile.ZipFile(full_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for entry_name, content in file_dict.items():
                zf.writestr(entry_name, content)
        return full_path

    def test_filter_matching(self):
        # Inclusion / Exclusion
        self.assertTrue(should_include_name("Sales_Report.pbix", ["Sales*"], None))
        self.assertFalse(should_include_name("Sales_Report.pbix", ["Finance*"], None))
        self.assertFalse(should_include_name("Sales_Report_Old.pbix", ["Sales*"], ["Old*"]))
        self.assertTrue(should_include_name("MyFolder", None, ["Archive"]))
        self.assertFalse(should_include_name("Archive_2020", None, ["Archive"]))

    def test_unpack_and_change_detection(self):
        pbix_path = self._create_pbix(
            "ReportA.pbix",
            {
                "DataModel": b"binary_data_v1",
                "Report/Layout": b'{"sections": []}',
                "SecurityBindings": b"secret123",
            },
        )

        # 1st run: all files should be created
        created, updated, skipped, errors = unpack_single_pbix(pbix_path)
        self.assertEqual(created, 3)
        self.assertEqual(updated, 0)
        self.assertEqual(skipped, 0)
        self.assertEqual(errors, 0)

        extracted_dir = os.path.join(self.test_dir, "ReportA")
        self.assertTrue(os.path.exists(os.path.join(extracted_dir, "DataModel")))
        self.assertTrue(os.path.exists(os.path.join(extracted_dir, "Report", "Layout")))

        # 2nd run with unchanged pbix: all should be skipped
        c2, u2, s2, e2 = unpack_single_pbix(pbix_path)
        self.assertEqual(c2, 0)
        self.assertEqual(u2, 0)
        self.assertEqual(s2, 3)
        self.assertEqual(e2, 0)

        # 3rd run: modify one file in pbix, add one file
        self._create_pbix(
            "ReportA.pbix",
            {
                "DataModel": b"binary_data_v2_MODIFIED",
                "Report/Layout": b'{"sections": []}',  # Unchanged
                "SecurityBindings": b"secret123",      # Unchanged
                "Metadata/Version": b"1.2.3",          # New
            },
        )
        c3, u3, s3, e3 = unpack_single_pbix(pbix_path)
        self.assertEqual(c3, 1)  # Version created
        self.assertEqual(u3, 1)  # DataModel updated
        self.assertEqual(s3, 2)  # Layout and SecurityBindings skipped
        self.assertEqual(e3, 0)

        with open(os.path.join(extracted_dir, "DataModel"), "rb") as f:
            self.assertEqual(f.read(), b"binary_data_v2_MODIFIED")

    def test_graceful_error_handling(self):
        # Create a file that is not a valid zip archive
        corrupted_pbix = os.path.join(self.test_dir, "Corrupted.pbix")
        with open(corrupted_pbix, "wb") as f:
            f.write(b"not a valid zip file")

        c, u, s, e = unpack_single_pbix(corrupted_pbix)
        self.assertEqual(e, 1)
        self.assertEqual(c, 0)

    def test_recursive_discovery_and_pruning(self):
        # Create nested folders
        self._create_pbix("root.pbix", {"a.txt": b"1"})
        self._create_pbix("FolderA/sub_a.pbix", {"a.txt": b"1"})
        self._create_pbix("FolderB_Archive/sub_b.pbix", {"a.txt": b"1"})
        self._create_pbix("FolderC/sub_c_draft.pbix", {"a.txt": b"1"})
        self._create_pbix("FolderC/sub_c_prod.pbix", {"a.txt": b"1"})

        # Non-recursive
        non_rec = find_pbix_files(self.test_dir, recursive=False, file_include=None, file_exclude=None, folder_include=None, folder_exclude=None)
        self.assertEqual(len(non_rec), 1)
        self.assertTrue(non_rec[0].endswith("root.pbix"))

        # Recursive with folder exclude
        rec = find_pbix_files(
            self.test_dir,
            recursive=True,
            file_include=None,
            file_exclude=["*draft*"],
            folder_include=None,
            folder_exclude=["*Archive*"],
        )
        found_basenames = [os.path.basename(p) for p in rec]
        self.assertIn("root.pbix", found_basenames)
        self.assertIn("sub_a.pbix", found_basenames)
        self.assertIn("sub_c_prod.pbix", found_basenames)
        self.assertNotIn("sub_b.pbix", found_basenames)       # folder excluded
        self.assertNotIn("sub_c_draft.pbix", found_basenames) # file excluded


if __name__ == "__main__":
    unittest.main()
