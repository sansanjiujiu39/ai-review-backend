# -*- coding: utf-8 -*-
"""
Vercel Serverless 入口
======================
@vercel/python 构建器要求入口文件顶层导出 `app`（ASGI 应用）。
此处引入项目根目录的 FastAPI app 并以 `app` 名义导出。
"""
import os
import sys

# 确保能 import 到项目根目录的 app.py 与 reviewer_data.json
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)
os.chdir(_ROOT)

from app import app as app  # noqa: E402,F401  Vercel 识别顶层 `app` 为 ASGI 应用
