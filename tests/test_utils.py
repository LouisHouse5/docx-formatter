#!/usr/bin/env python3
"""测试 utils.py 公共函数"""
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'scripts'))

import unittest
from utils import emu_to_pt, emu_to_inch, pt_to_emu, inch_to_emu, ns_tag, ns_attr, compare_para_props

class TestEmuConversion(unittest.TestCase):
    def test_emu_to_pt(self):
        self.assertEqual(emu_to_pt(12700), 1.0)
        self.assertEqual(emu_to_pt(25400), 2.0)
        self.assertIsNone(emu_to_pt(None))
        self.assertIsNone(emu_to_pt(0))

    def test_pt_to_emu(self):
        self.assertEqual(pt_to_emu(1), 12700)
        self.assertEqual(pt_to_emu(12), 152400)
        self.assertEqual(pt_to_emu(0), 0)

    def test_emu_to_inch(self):
        self.assertAlmostEqual(emu_to_inch(914400), 1.0)
        self.assertAlmostEqual(emu_to_inch(457200), 0.5)
        self.assertIsNone(emu_to_inch(None))

    def test_inch_to_emu(self):
        self.assertEqual(inch_to_emu(1), 914400)
        self.assertEqual(inch_to_emu(2), 1828800)
        self.assertEqual(inch_to_emu(0), 0)

class TestNsTag(unittest.TestCase):
    def test_ns_tag_short(self):
        result = ns_tag('w:p')
        self.assertIn('p', result)
        self.assertIn('schemas.openxmlformats', result)

    def test_ns_tag_no_prefix(self):
        result = ns_tag('val')
        self.assertIn('val', result)
        self.assertIn('schemas.openxmlformats', result)

    def test_ns_tag_w14(self):
        result = ns_tag('w14:sz')
        self.assertIn('sz', result)
        self.assertIn('microsoft.com', result)

    def test_ns_attr(self):
        result = ns_attr('val')
        self.assertIn('val', result)
        self.assertIn('schemas.openxmlformats', result)

class TestColors(unittest.TestCase):
    def test_colors_exist(self):
        from utils import Colors, log_ok, log_warn, log_err, log_info
        self.assertTrue(hasattr(Colors, 'OK'))
        self.assertTrue(hasattr(Colors, 'WARN'))
        self.assertTrue(hasattr(Colors, 'ERR'))
        self.assertTrue(hasattr(Colors, 'INFO'))

class TestCheckFileExists(unittest.TestCase):
    def test_check_file_exists_true(self):
        from utils import check_file_exists
        # Should not raise for existing file
        check_file_exists(__file__, '测试文件')

    def test_check_file_exists_false(self):
        from utils import check_file_exists
        with self.assertRaises(SystemExit) as cm:
            check_file_exists('/nonexistent/path/to/file.txt', '测试文件')
        self.assertEqual(cm.exception.code, 1)

class TestCompareParaProps(unittest.TestCase):
    def test_identical_props(self):
        info = {
            'font_name': '宋体', 'font_size': 152400, 'alignment': None,
            'bold': False, 'line_spacing': 1.5, 'line_spacing_rule': 'ONE_POINT_FIVE',
            'style': 'Normal', 'first_line_indent': 304800,
            'space_before': None, 'space_after': None,
        }
        self.assertEqual(compare_para_props(info, info), [])

    def test_font_diff(self):
        target = {'font_name': '微软雅黑', 'font_size': 152400, 'alignment': None,
                  'bold': False, 'line_spacing': 1.5, 'line_spacing_rule': 'ONE_POINT_FIVE',
                  'style': 'Normal', 'first_line_indent': 304800,
                  'space_before': None, 'space_after': None}
        tmpl = {'font_name': '宋体', 'font_size': 152400, 'alignment': None,
                'bold': False, 'line_spacing': 1.5, 'line_spacing_rule': 'ONE_POINT_FIVE',
                'style': 'Normal', 'first_line_indent': 304800,
                'space_before': None, 'space_after': None}
        issues = compare_para_props(target, tmpl)
        self.assertEqual(len(issues), 1)
        self.assertIn('font=', issues[0])
        self.assertIn('宋体', issues[0])

    def test_multiple_diffs(self):
        target = {'font_name': '微软雅黑', 'font_size': 120000, 'alignment': 1,
                  'bold': True, 'line_spacing': 1.5, 'line_spacing_rule': 'ONE_POINT_FIVE',
                  'style': 'Normal', 'first_line_indent': 304800,
                  'space_before': None, 'space_after': None}
        tmpl = {'font_name': '宋体', 'font_size': 152400, 'alignment': None,
                'bold': False, 'line_spacing': 1.5, 'line_spacing_rule': 'ONE_POINT_FIVE',
                'style': 'Normal', 'first_line_indent': 304800,
                'space_before': None, 'space_after': None}
        issues = compare_para_props(target, tmpl)
        self.assertEqual(len(issues), 4)

    def test_none_tmpl_ignored(self):
        target = {'font_name': '宋体', 'font_size': 152400, 'alignment': None,
                  'bold': False, 'line_spacing': 1.5, 'line_spacing_rule': 'ONE_POINT_FIVE',
                  'style': 'Normal', 'first_line_indent': 304800,
                  'space_before': None, 'space_after': None}
        tmpl = {'font_name': None, 'font_size': None, 'alignment': None,
                'bold': None, 'line_spacing': None, 'line_spacing_rule': 'AT_LEAST',
                'style': None, 'first_line_indent': None,
                'space_before': None, 'space_after': None}
        issues = compare_para_props(target, tmpl)
        self.assertEqual(len(issues), 1)
        self.assertIn('ls_rule', issues[0])


if __name__ == '__main__':
    unittest.main()
