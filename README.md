# 猛士汽车售后无忧服务体系网页

## 本地预览

双击 `启动售后网点网页.command`，浏览器会打开本地页面。修改 `售后网络.xlsx` 后刷新网页即可读取最新数据。

## GitHub Pages 发布

本项目已整理为 GitHub Pages 可发布结构：

- `docs/index.html`
- `docs/app.js`
- `docs/vendor/xlsx.full.min.js`
- `docs/售后网络.xlsx`

发布步骤：

1. 将项目推送到 GitHub 仓库。
2. 进入仓库 `Settings` -> `Pages`。
3. `Build and deployment` 选择 `Deploy from a branch`。
4. Branch 选择 `main`，目录选择 `/docs`。
5. 保存后等待 GitHub Pages 部署完成。

后续如需更新数据，把新的 `售后网络.xlsx` 同步覆盖到 `docs/售后网络.xlsx` 即可。
