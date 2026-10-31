# -*- coding: utf-8 -*-
"""
智慧牧业数字化大屏系统 · 云端路径自适应桥接引导程序
当 Streamlit Cloud 的 Main file path 被配置为 smart_pig_demo/app.py 时，
此桥接程序自动将工作目录重定向至仓库根目录并无缝执行主看板 app.py，
彻底解决 '/mount/src/smart-livestock-demo/smart_pig_demo does not exist' 的目录层级报错。
"""

import os
import sys

# 1. 动态获取仓库根目录路径并加入环境变量首位
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT_DIR = os.path.dirname(CURRENT_DIR)

if REPO_ROOT_DIR not in sys.path:
    sys.path.insert(0, REPO_ROOT_DIR)

# 2. 切换当前工作目录至仓库根目录，保障相对路径资源（表格、模板、资产等）100% 零报错命中
os.chdir(REPO_ROOT_DIR)

# 3. 动态加载执行仓库根目录主看板程序 app.py
main_app_path = os.path.join(REPO_ROOT_DIR, "app.py")
if os.path.exists(main_app_path):
    with open(main_app_path, "r", encoding="utf-8") as f:
        code_obj = compile(f.read(), main_app_path, "exec")
        exec(code_obj, globals())
else:
    raise FileNotFoundError(f"未在根目录找到主看板程序: {main_app_path}")
