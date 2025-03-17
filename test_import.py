#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
测试oect_excel_processor包的导入和基本功能
"""

import os
import sys
import importlib

# 添加当前目录到Python路径
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

def test_import():
    """测试导入包及其模块"""
    print("Python路径:")
    for path in sys.path:
        print(f"  - {path}")
    
    print("\n尝试导入包:")
    try:
        import oect_excel_processor
        print(f"✓ 成功导入oect_excel_processor包")
        print(f"  包位置: {oect_excel_processor.__file__}")
        print(f"  包版本: {oect_excel_processor.__version__}")
        print(f"  可用模块: {dir(oect_excel_processor)}")
        
        # 测试导入主要类
        try:
            from oect_excel_processor import ExcelProcessor
            print(f"✓ 成功导入ExcelProcessor类")
        except ImportError as e:
            print(f"✗ 导入ExcelProcessor类失败: {e}")
        
        try:
            from oect_excel_processor import BatchExcelProcessor
            print(f"✓ 成功导入BatchExcelProcessor类")
        except ImportError as e:
            print(f"✗ 导入BatchExcelProcessor类失败: {e}")
        
    except ImportError as e:
        print(f"✗ 导入包失败: {e}")
        
        # 尝试直接导入模块
        print("\n尝试直接导入模块:")
        try:
            sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'oect_excel_processor'))
            import excel_processor
            print(f"✓ 成功导入excel_processor模块")
        except ImportError as e:
            print(f"✗ 导入excel_processor模块失败: {e}")
        
        try:
            import batch_processor
            print(f"✓ 成功导入batch_processor模块")
        except ImportError as e:
            print(f"✗ 导入batch_processor模块失败: {e}")

def main():
    """主函数"""
    print("测试oect_excel_processor包的导入和基本功能")
    print("=" * 50)
    
    test_import()
    
    return 0

if __name__ == "__main__":
    sys.exit(main()) 