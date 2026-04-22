# 健康大比拼海报（GitHub Pages）

本目录已包含 GitHub Pages 入口文件 `index.html`。

## 发布步骤

1. 在 GitHub 上创建一个新仓库（例如：`health-poster`）。
2. 在本地项目目录执行：

```bash
git init
git add .
git commit -m "init: add health poster site"
git branch -M main
git remote add origin https://github.com/<你的用户名>/<你的仓库名>.git
git push -u origin main
```

3. 打开仓库页面：`Settings` -> `Pages`
4. `Build and deployment` 中选择：
   - `Source`: `Deploy from a branch`
   - `Branch`: `main` / `/ (root)`
5. 保存后等待 1-3 分钟，访问：

`https://<你的用户名>.github.io/<你的仓库名>/`

## 说明

- 页面主文件：`index.html`
- LOGO 资源：`人防logo.png`
- 原始文件仍保留：`健康大比拼海报.html`
