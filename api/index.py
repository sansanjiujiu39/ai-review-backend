# -*- coding: utf-8 -*-
"""
Vercel Serverless 入口
======================
Vercel Python 运行时要求函数导出名为 `asgi_app` 的 ASGI 应用。
此处引入项目根目录的 FastAPI app（app.py）并导出。
"""
import os
import sys

# 确保能 import 到项目根目录的 app.py 与 reviewer_data.json
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)
os.chdir(_ROOT)

from app import app as asgi_app  # noqa: E402
