# 明鉴 · AI 评审打分后端（v1.8.0）

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

把「明鉴」创新竞赛评审专家封装为 HTTP API 服务。前端（国创赛自评网站）上传计划书文本，
后端按明鉴 V3.1 五步打分法 + 官方赛道规则 + 155+ 实证锚点库组装评审指令，调用 DeepSeek 打分，
返回结构化评分 JSON。**DeepSeek API Key 只存在后端环境变量中，不会出现在前端代码里。**

## 目录结构

```
ai-review-backend/
├── app.py                  # FastAPI 服务（打分引擎）
├── reviewer_data.json      # 评审规则数据（10 赛道规则 + 锚点库 + 七档标尺）
├── requirements.txt        # Python 依赖
├── .env.example            # 环境变量示例（复制为 .env 并填 Key）
├── render.yaml             # Render 云平台一键部署配置
├── vercel.json             # Vercel 部署配置（免费、无需信用卡）
├── api/index.py            # Vercel Serverless 入口（导出 asgi_app）
├── LICENSE                  # MIT 开源许可证
└── README.md
```

## API 接口

| 方法 | 路径 | 说明 |
|---|---|---|
| GET | `/api/health` | 健康检查，返回引擎版本与 Key 是否配置 |
| GET | `/api/tracks` | 10 个赛道列表（前端下拉框） |
| POST | `/api/review` | 打分。请求体：`{"track":"gaojiao_chuangyi","text":"计划书正文(≥50字)","material":"bp"}` |

`/api/review` 返回：
```json
{
  "ok": true,
  "track": "gaojiao_chuangyi",
  "material": "bp",
  "result": {
    "total": 62.5, "level": "省银", "level_range": "58-64",
    "summary": "一句话定调", "max_leverage": "最大改造杠杆",
    "dims": [{"name": "个人成长", "full": 30, "score": 20.5, "diagnosis": "...",
              "subdims": [{"name": "立德树人", "full": 5, "score": 3, "loss": "..."}]}],
    "highlights": [...], "weaknesses": [...],
    "suggestions": [{"priority": 1, "text": "..."}],
    "potential": "..."
  }
}
```

赛道 key（与前端 TRACK_KEYS 一致）：
`gaojiao_chuangyi / gaojiao_chuangye / honglv_gongyi / honglv_chuangyi / honglv_chuangye / zhijiao_chuangyi / zhijiao_chuangye / chanye_qimingti / chanye_chengguo / mengya`

## 本地运行

```bash
# 1. 安装依赖（建议 venv）
python -m venv venv
venv\Scripts\activate          # Windows
pip install -r requirements.txt

# 2. 配置 Key
copy .env.example .env         # 编辑 .env 填入 DEEPSEEK_API_KEY=sk-xxx

# 3. 启动（Windows 下 .env 需手动加载，或直接设置环境变量）
set DEEPSEEK_API_KEY=sk-xxx
python -m uvicorn app:app --host 0.0.0.0 --port 8000

# 4. 验证
curl http://127.0.0.1:8000/api/health
```

## 前端接入

`guochuang-score-site/index.html` 已改造为通过 `API_BASE` 调用后端，默认 `http://localhost:8000`。

- 本地联调：前端直接用默认值即可（注意后端已开启 CORS）。
- 部署后：不需要改代码，在浏览器控制台执行一次：
  ```js
  localStorage.setItem('mj_api_base', 'https://你的后端域名')
  ```
  或直接改 `index.html` 中 `const API_BASE = ...` 的默认值。

## 部署方案（三选一）

### 方案 A · Vercel（免费，无需信用卡，推荐）
1. 把本目录推送到 GitHub 仓库（已有 `vercel.json` + `api/index.py` 适配）；
2. 登录 [vercel.com](https://vercel.com)（用 GitHub 账号登录）→ **Add New → Project**；
3. Import 仓库 `ai-review-backend`，Framework Preset 选 **Other**；
4. **Environment Variables** 添加 `DEEPSEEK_API_KEY = sk-xxx`；
5. 点 **Deploy**，约 1-2 分钟完成，得到 `https://ai-review-backend.vercel.app`；
6. 前端控制台设置 `localStorage.setItem('mj_api_base','https://ai-review-backend.vercel.app')`。

> 注意：Vercel Hobby（免费）计划函数最长执行 60 秒，单次评审一般 30-50 秒，可正常使用；
> 若提示超时，重试一次即可。

### 方案 B · Render（免费，但需要绑定国际信用卡验证，不扣费）
1. 把本目录推送到 GitHub 仓库；
2. 登录 [render.com](https://render.com) → New → Blueprint / Web Service → 选择该仓库；
3. Runtime 选 **Python**，构建命令 `pip install -r requirements.txt`，启动命令
   `uvicorn app:app --host 0.0.0.0 --port $PORT`（或直接用已有的 `render.yaml`）；
4. 添加环境变量 `DEEPSEEK_API_KEY = sk-xxx`；
5. 部署完成后得到 `https://xxx.onrender.com`，前端控制台设置 `mj_api_base` 指向它。
   （国内访问 Render 可能较慢，介意可用方案 A/C）

### 方案 C · 国内云函数（腾讯云 CloudBase / 阿里云函数计算）
- 打包本目录为 zip，在云函数控制台创建 Python 3 函数；
- 入口 `app.app`（FastAPI 挂载到 `api.main` 适配云函数网关）；
- 环境变量填 `DEEPSEEK_API_KEY`；
- 云函数有免费额度，且国内访问快。

### 方案 D · 自己的服务器 / 校园服务器
```bash
pip install -r requirements.txt
nohup uvicorn app:app --host 0.0.0.0 --port 8000 &
```
反向代理（Nginx）加 HTTPS 后，前端 `mj_api_base` 指向你的域名。

## ⚠️ 安全提醒（重要）

**原网站曾把 DeepSeek API Key 硬编码在前端 `index.html` 中并部署到公开的 GitHub Pages**，
Key 已泄露。请立即：

1. 登录 [platform.deepseek.com](https://platform.deepseek.com) → API Keys → **删除/重置该 Key**；
2. 在「用量管理」中**设置消费限额/余额警报**，防止历史泄露期被盗刷；
3. 若仓库是公开的，历史 commit 中仍保留该 Key——重置后旧 Key 作废，无需清理仓库，但建议把仓库改为私有或删库重建。

## 升级说明

- 打分逻辑升级时，只需更新 `app.py` 中的 `SYSTEM_PROMPT`（明鉴方法论）与 `reviewer_data.json`（规则/锚点库），前端无需改动。
- `reviewer_data.json` 由网站原内嵌数据提取，规则源与专家包 `references/competition-rules/` 一致。

## 开源说明

本项目以 **MIT License** 开源，版权所有：© 2026 Huang Rende。
允许商用、修改与分发，使用/修改后保留版权声明即可。

- **在线演示前端**：<https://sansanjiujiu39.github.io/guochuang-score/>
- **前端仓库**：<https://github.com/sansanjiujiu39/guochuang-score>
- **打分引擎**：DeepSeek API（调用方自行配置 `DEEPSEEK_API_KEY`，本仓库不含任何真实密钥）
