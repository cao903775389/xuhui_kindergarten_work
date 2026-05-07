# 徐汇区幼儿园筛选方案

这是一个静态网页项目，用于展示徐汇区公办幼儿园筛选方案与园区点位资料。

## 本地生成

```bash
node build_xuhui_kindergarten_report.mjs
```

生成文件会写入 `outputs/`，其中 `outputs/index.html` 是 GitHub Pages 的首页。

## GitHub Pages

仓库已配置 `.github/workflows/pages.yml`。推送到 `main` 分支后，GitHub Actions 会把 `outputs/` 部署到 GitHub Pages。

部署地址可在 GitHub 仓库的 `Settings -> Pages` 或 `Actions -> Deploy GitHub Pages` 中查看。
