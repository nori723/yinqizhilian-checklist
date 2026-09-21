# 银企直联材料清单（公开版）

按「先懂概念 → 再选业务 → 再看材料」三步查清单的静态网页。纯前端、单文件、无外部依赖，使用 hash 路由，可直接部署到任意静态托管。

- **在线访问**：https://nori723.github.io/yinqizhilian-checklist/
- **源码结构**：
  - `index.html` —— 网站本体（数据内嵌在 `<script id="app-data">` 中）。
  - `html工具/` —— 生成工具：`build.py` 读取《银企直联材料清单.xlsx》→ `data_standardized.json` + `mapping.json` → 套 `template.html` 产出 `index.html`。
- **如何更新**：改 `html工具/data_standardized.json` / `template.html` / `mapping.json` 后，本地重跑 `build.py` 把新 `index.html` 放回仓库根目录，提交并推送 `main` 分支，GitHub Actions 会自动重新部署。

> 说明：本仓库按全量（含行内版）发布，已确认对外公开。
