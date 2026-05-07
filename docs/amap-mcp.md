# 高德地图 MCP 接入说明

当前网页已预留高德增强字段：

- 高德 POI 名称
- 高德经纬度
- 距西岸网易研发中心

## 当前限制

高德 MCP 必须使用你自己的高德 Web 服务 Key。本仓库不会提交真实 Key。

## 官方接入方式

高德官方文档提供两种方式：

1. Streamable HTTP：

```json
{
  "mcpServers": {
    "amap-maps-streamableHTTP": {
      "url": "https://mcp.amap.com/mcp?key=你的高德Key"
    }
  }
}
```

2. Node.js I/O：

```json
{
  "mcpServers": {
    "amap-maps": {
      "command": "npx",
      "args": ["-y", "@amap/amap-maps-mcp-server"],
      "env": {
        "AMAP_MAPS_API_KEY": "你的高德Key"
      }
    }
  }
}
```

本项目已提供 `.cursor/mcp.json.example`。如果在 Cursor 中使用，可以复制为 `.cursor/mcp.json` 后填入 Key。

## 后续批量增强口径

对每个园区使用以下高德能力：

- 关键词搜索：`上海市徐汇区 + 幼儿园名 + 园区名 + 地址`
- 详情搜索：用 POI ID 获取正式名称、地址、经纬度、电话、类型
- 距离测量：从园区经纬度到“西岸网易研发中心”经纬度

查询结果建议写回 `build_xuhui_kindergarten_report.mjs` 的增强字段，重新运行：

```bash
node build_xuhui_kindergarten_report.mjs
```
