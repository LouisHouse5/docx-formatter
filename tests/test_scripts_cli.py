#!/usr/bin/env python3
"""测试各脚本的 CLI 参数解析配置"""
import sys
import os
import argparse
import importlib
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'scripts'))

import unittest


def _get_parser(module_name):
    """安全导入模块并提取 parser 对象（避免 sys.argv 和文件检查干扰）"""
    saved_argv = sys.argv
    sys.argv = [module_name]
    try:
        mod = importlib.import_module(module_name)
        return mod.parser
    except SystemExit:
        # 某些模块在导入时会因文件不存在而退出，直接构造 parser
        import argparse
        return argparse.ArgumentParser()
    finally:
        sys.argv = saved_argv


class TestFixDocxTemplateCLI(unittest.TestCase):
    def test_parser_has_batch_file(self):
        parser = _get_parser('fix_docx_template')
        actions = {a.dest: a for a in parser._actions}
        self.assertIn('batch_file', actions)
        self.assertEqual(actions['batch_file'].option_strings, ['--batch-file', '-b'])

    def test_parser_has_config(self):
        parser = _get_parser('fix_docx_template')
        actions = {a.dest: a for a in parser._actions}
        self.assertIn('config', actions)
        self.assertEqual(actions['config'].option_strings, ['--config', '-c'])

    def test_parser_has_template(self):
        parser = _get_parser('fix_docx_template')
        actions = {a.dest: a for a in parser._actions}
        self.assertIn('template_opt', actions)
        self.assertIn('--template', actions['template_opt'].option_strings)
        self.assertIn('-t', actions['template_opt'].option_strings)

    def test_parse_batch_file(self):
        parser = _get_parser('fix_docx_template')
        args = parser.parse_args(['--batch-file', 'files.txt', '--template', 'tmpl.docx'])
        self.assertEqual(args.batch_file, 'files.txt')
        self.assertEqual(args.template_opt, 'tmpl.docx')

    def test_parse_config(self):
        parser = _get_parser('fix_docx_template')
        args = parser.parse_args(['--config', 'cfg.json', 'target.docx'])
        self.assertEqual(args.config, 'cfg.json')
        self.assertEqual(args.target, 'target.docx')


class TestVerifyDocxCLI(unittest.TestCase):
    def test_parse_args(self):
        parser = _get_parser('verify_docx')
        args = parser.parse_args(['target.docx', 'template.docx'])
        self.assertEqual(args.target, 'target.docx')
        self.assertEqual(args.template, 'template.docx')


class TestCopyStylesCLI(unittest.TestCase):
    def test_parse_args(self):
        parser = _get_parser('copy_styles')
        args = parser.parse_args(['tmpl.docx', 'target.docx'])
        self.assertEqual(args.template, 'tmpl.docx')
        self.assertEqual(args.target, 'target.docx')


class TestCopyHeadersFootersCLI(unittest.TestCase):
    def test_parse_args(self):
        parser = _get_parser('copy_headers_footers')
        args = parser.parse_args(['tmpl.docx', 'target.docx'])
        self.assertEqual(args.template, 'tmpl.docx')
        self.assertEqual(args.target, 'target.docx')


if __name__ == '__main__':
    unittest.main()
