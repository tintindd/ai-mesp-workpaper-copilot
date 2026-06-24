# AI-MESP Workpaper Copilot

面向 MESP 底稿的智能取数、自动重算与异常复核助手。

## 当前能力

- 识别 CO03、KSBT、3611、CKM3 支持文件
- 支持截图和 Excel/CSV 表格混合上传
- 自动判断样本、订单号、报表类型和文件类型
- 自动识别 SAP 表头并生成字段映射
- 输出异常检查、自动追问和 Evidence Traceability
- 支持上传式 Web App，不依赖固定本地 `training_files` 文件夹

## 运行上传版 Web App

```powershell
python backend\app.py
```

然后打开：

```text
http://127.0.0.1:8766/
```

在网页中选择样本文件夹或多份支持文件，点击 `Analyze` 即可。

## 部署成不用本地运行的版本

项目已支持 Docker 部署。部署到公司服务器、私有云或 PaaS 后，用户只需要打开网址并上传文件，不需要在本地运行 Python。

### Streamlit Community Cloud 部署

项目已提供 Streamlit 入口文件：

```text
streamlit_app.py
```

部署步骤：

1. 将当前代码 push 到 GitHub 仓库。
2. 打开 Streamlit Community Cloud。
3. 选择 `New app`。
4. 选择 GitHub 仓库和分支。
5. Main file path 填写：

```text
streamlit_app.py
```

6. 点击 Deploy。

部署完成后，同事只需要打开 Streamlit 生成的网址，上传支持文件或 zip 包即可分析。

建议上传 zip 包，以保留样本文件夹结构，例如：

```text
样本1/
  1.订单编号11000437-CO03-表格.xlsx
  1.订单编号11000437-CO03-截图.png
```

注意：Streamlit Community Cloud 是外部云平台。真实 SAP/审计文件可能包含敏感信息，正式使用前请确认公司允许上传到该平台。

构建镜像：

```bash
docker build -t ai-mesp-workpaper-copilot .
```

运行容器：

```bash
docker run --rm -p 8766:8766 ai-mesp-workpaper-copilot
```

访问：

```text
http://localhost:8766/
```

云平台部署时需要暴露 Web 服务端口，并设置环境变量：

```text
PORT=8766
HOST=0.0.0.0
```

注意：真实 SAP/审计支持文件可能包含敏感信息，建议优先部署在公司内网、VPN 或私有云环境，不建议直接暴露在公网。

## 本地工具脚本

如果仍需直接扫描本地训练文件夹：

```powershell
python scripts\mesp_automation_engine.py --period "2025.01.01-2025.12.31" --program SPD03015
```

默认输入：

```text
training_files/训练文件
```

默认输出：

```text
mesp_automation_result.json
```

## 敏感数据说明

`training_files/`、`mesp_automation_result.json`、`mesp_training_profile.json` 等本地数据文件默认不提交到 GitHub。

## DeepSeek API 配置

数据识别与清洗功能通过 DeepSeek API 做字段标准化、文件名清洗和后续 OCR 文本结构化。API Key 不要写入代码或提交到 GitHub。

在 Streamlit Community Cloud 中进入：

```text
App settings -> Secrets
```

添加：

```toml
DEEPSEEK_API_KEY = "你的 DeepSeek API Key"
DEEPSEEK_BASE_URL = "https://api.deepseek.com"
DEEPSEEK_MODEL = "deepseek-chat"
```

其中 `DEEPSEEK_BASE_URL` 和 `DEEPSEEK_MODEL` 可以不填，系统会默认使用 DeepSeek 官方 OpenAI-compatible API 地址和 `deepseek-chat` 模型。

## 在线 PaddleOCR 服务配置

OCR 识别建议使用在线 PaddleOCR 服务，避免在 Streamlit Cloud 中安装较重的本地 OCR 模型依赖。

在 Streamlit Community Cloud 的 `App settings -> Secrets` 中继续添加：

```toml
PADDLEOCR_API_URL = "你的在线 PaddleOCR 服务调用地址"
PADDLEOCR_API_KEY = "你的在线 PaddleOCR Token 或 API Key"
PADDLEOCR_AUTH_HEADER = "Authorization"
PADDLEOCR_AUTH_SCHEME = "Bearer"
PADDLEOCR_FILE_FIELD = "file"
```

如果平台要求的鉴权头不是 `Authorization: Bearer xxx`，可以按平台文档调整：

- `PADDLEOCR_AUTH_HEADER`：例如 `Authorization`、`X-API-Key`
- `PADDLEOCR_AUTH_SCHEME`：例如 `Bearer`；如果平台只要求原始 token，可留空
- `PADDLEOCR_FILE_FIELD`：图片上传字段名，默认 `file`
