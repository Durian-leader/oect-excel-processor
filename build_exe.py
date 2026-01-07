#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
PyInstaller 打包脚本
用于将 OECT Excel Processor GUI 打包为单个可执行文件
"""

import os
import sys
import subprocess
import shutil

def build_exe():
    """构建单个可执行文件"""
    
    # 获取当前目录
    project_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 资源文件路径
    icon_path = os.path.join(project_dir, "oect_excel_processor", "resources", "icon.ico")
    
    # 检查图标文件是否存在
    if not os.path.exists(icon_path):
        print(f"警告: 图标文件不存在: {icon_path}")
        icon_arg = []
    else:
        icon_arg = [f"--icon={icon_path}"]
    
    # 构建命令
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",           # 打包为单个文件
        "--noconsole",         # 无控制台窗口
        "--clean",             # 清理临时文件
        "--name=OECT-Excel-Processor",  # 输出文件名
        *icon_arg,
        # 添加数据文件
        f"--add-data={os.path.join(project_dir, 'oect_excel_processor', 'resources')};oect_excel_processor/resources",
        # 隐藏导入
        "--hidden-import=pandas",
        "--hidden-import=numpy",
        "--hidden-import=natsort",
        "--hidden-import=xlrd",
        # 入口点
        os.path.join(project_dir, "oect_excel_processor", "gui.py"),
    ]
    
    print("开始构建...")
    print(f"命令: {' '.join(cmd)}")
    
    # 执行构建
    result = subprocess.run(cmd, cwd=project_dir)
    
    if result.returncode == 0:
        print("\n✓ 构建成功!")
        dist_path = os.path.join(project_dir, "dist", "OECT-Excel-Processor.exe")
        if os.path.exists(dist_path):
            size_mb = os.path.getsize(dist_path) / (1024 * 1024)
            print(f"  输出文件: {dist_path}")
            print(f"  文件大小: {size_mb:.1f} MB")
    else:
        print("\n✗ 构建失败!")
        sys.exit(1)


if __name__ == "__main__":
    build_exe()
