# 上海幼儿园落地策略工具

这是一个静态策略工具，用于把徐汇、闵行、浦东的幼儿园点位、招生口径、通勤距离、租房板块和材料核验动作放在同一套数据模型里比较。

## 本地生成

```bash
npm run build
npm run test:strategy
npm run validate
```

生成文件会写入 `outputs/`，其中 `outputs/index.html` 是 GitHub Pages 首页。通用导出文件包括：

- `outputs/上海幼儿园落地策略工具.html`
- `outputs/上海幼儿园落地策略工具点位数据.csv`
- `outputs/上海幼儿园落地策略工具.xlsx`

## GitHub Pages

仓库已配置 `.github/workflows/pages.yml`。推送到 `main` 分支后，GitHub Actions 会先校验已生成的数据，再把 `outputs/` 部署到 GitHub Pages。
