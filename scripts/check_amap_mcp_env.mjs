const key = process.env.AMAP_MAPS_API_KEY;

if (!key) {
  console.error("缺少 AMAP_MAPS_API_KEY。请先在高德开放平台创建 Web 服务 Key。");
  console.error("官方 MCP Node I/O 命令：AMAP_MAPS_API_KEY=你的Key npx -y @amap/amap-maps-mcp-server");
  process.exit(1);
}

console.log(JSON.stringify({
  ok: true,
  mcpServer: "@amap/amap-maps-mcp-server",
  env: "AMAP_MAPS_API_KEY",
  keyPrefix: `${key.slice(0, 4)}***`,
}, null, 2));
