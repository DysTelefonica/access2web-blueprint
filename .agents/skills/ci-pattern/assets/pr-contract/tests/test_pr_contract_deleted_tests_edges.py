#!/usr/bin/env python3
"""HR-88 (#1509): casos límite del aviso por tests borrados. stdlib exclusivamente."""

import unittest

from test_pr_contract import ALL_CHECKS
from test_pr_contract_deleted_tests import (
    BASE_TREE,
    ROUTES,
    TEST_SERVICE_WITHOUT_SEARCH,
    run_deleted_tests,
    warnings_of,
)


class DeletedTestsEdgeTests(unittest.TestCase):
    """HR-88 (#1509): referencias, movimientos y medición del aviso."""

    def test_module_no_longer_referenced_does_not_warn(self):
        p = run_deleted_tests({"tests/test_service.py": TEST_SERVICE_WITHOUT_SEARCH,
                               "app/routes.py": "def index():\n    return []\n"})
        self.assertEqual(p.returncode, 0, p.stdout + p.stderr)
        self.assertEqual(warnings_of(p), [], p.stdout)

    def test_relative_import_counts_as_a_reference(self):
        routes = ROUTES.replace("from app.service import", "from .service import")
        base = {**BASE_TREE, "app/routes.py": routes}
        p = run_deleted_tests({"tests/test_service.py": TEST_SERVICE_WITHOUT_SEARCH},
                              base_files=base)
        self.assertEqual(p.returncode, 0, p.stdout + p.stderr)
        self.assertEqual(len(warnings_of(p)), 1, p.stdout)

    def test_whole_test_file_deleted_warns_for_each_test(self):
        p = run_deleted_tests({"tests/test_service.py": None})
        self.assertEqual(p.returncode, 0, p.stdout + p.stderr)
        warnings = warnings_of(p)
        self.assertEqual(len(warnings), 2, p.stdout)
        self.assertTrue(any("::test_list_items" in w for w in warnings), warnings)
        self.assertTrue(any("::test_search_escapes_like" in w for w in warnings), warnings)

    def test_test_moved_to_another_file_does_not_warn(self):
        moved = ("from app.service import search_items\n\n\n"
                 "def test_search_escapes_like():\n    assert search_items('%') == ['\\\\%']\n")
        p = run_deleted_tests({"tests/test_service.py": TEST_SERVICE_WITHOUT_SEARCH,
                               "tests/test_search.py": moved})
        self.assertEqual(p.returncode, 0, p.stdout + p.stderr)
        self.assertEqual(warnings_of(p), [], p.stdout)

    def test_without_the_diff_file_the_rule_is_not_measured_and_says_so(self):
        p = run_deleted_tests({"tests/test_service.py": TEST_SERVICE_WITHOUT_SEARCH},
                              with_diff=False)
        self.assertEqual(p.returncode, 0, p.stdout + p.stderr)
        self.assertEqual(warnings_of(p), [], p.stdout)
        self.assertIn("HR-88", p.stderr)
        self.assertIn("--diff-file", p.stderr)

    def test_switched_off_by_policy_does_not_warn(self):
        checks = {name: True for name in ALL_CHECKS}
        checks["deleted_tests"] = False
        p = run_deleted_tests({"tests/test_service.py": TEST_SERVICE_WITHOUT_SEARCH},
                              checks=checks)
        self.assertEqual(p.returncode, 0, p.stdout + p.stderr)
        self.assertEqual(warnings_of(p), [], p.stdout)
        self.assertIn("deleted_tests", p.stderr)


if __name__ == "__main__":
    unittest.main(verbosity=2)
