import fs from "node:fs/promises";
import path from "node:path";

const key = process.env.AMAP_MAPS_API_KEY;
if (!key) {
  console.error("缺少 AMAP_MAPS_API_KEY。");
  process.exit(1);
}

const root = process.cwd();
const csvPath = path.join(root, "outputs", "徐汇区幼儿园园区位置表.csv");
const outputPath = path.join(root, "data", "amap_enrichment.json");
const city = "上海";
const officeKeyword = "上海市徐汇区西岸网易研发中心";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const parseCsv = (text) => {
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;
  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];
    if (inQuotes) {
      if (char === '"' && next === '"') {
        cell += '"';
        i += 1;
      } else if (char === '"') {
        inQuotes = false;
      } else {
        cell += char;
      }
    } else if (char === '"') {
      inQuotes = true;
    } else if (char === ",") {
      row.push(cell);
      cell = "";
    } else if (char === "\n") {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
    } else if (char !== "\r") {
      cell += char;
    }
  }
  if (cell || row.length) {
    row.push(cell);
    rows.push(row);
  }
  return rows;
};

const toObjects = (rows) => {
  const header = rows[0];
  return rows.slice(1).filter((row) => row.length === header.length).map((row) => Object.fromEntries(header.map((name, index) => [name, row[index]])));
};

const requestAmap = async (endpoint, params) => {
  const url = new URL(endpoint);
  url.searchParams.set("key", key);
  for (const [name, value] of Object.entries(params)) {
    if (value !== undefined && value !== "") url.searchParams.set(name, value);
  }
  const response = await fetch(url);
  if (!response.ok) throw new Error(`HTTP ${response.status}: ${url.pathname}`);
  const json = await response.json();
  if (json.status !== "1") {
    throw new Error(`Amap error ${json.infocode || ""}: ${json.info || "unknown"}`);
  }
  await sleep(80);
  return json;
};

const textSearch = async (keywords) => requestAmap("https://restapi.amap.com/v3/place/text", {
  keywords,
  city,
  citylimit: "true",
  offset: "8",
  page: "1",
  extensions: "all",
});

const distance = async (origin, destination) => requestAmap("https://restapi.amap.com/v3/distance", {
  origins: origin,
  destination,
  type: "0",
});

const scorePoi = (row, poi) => {
  if (poi.adname && !`${poi.adname}`.includes("徐汇")) return -100;
  const name = `${poi.name || ""}`;
  const address = `${poi.address || ""}`;
  let score = 0;
  if (name.includes(row["幼儿园"])) score += 8;
  for (const word of row["幼儿园"].replaceAll("上海市徐汇区", "").replaceAll("民办", "").split(/[（）() /]/).filter(Boolean)) {
    if (word.length >= 2 && name.includes(word)) score += 2;
  }
  if (row["园区/分园"] !== "本部" && name.includes(row["园区/分园"].replaceAll("/", ""))) score += 2;
  if (address.includes(row["地址"].slice(0, 4))) score += 3;
  if (address.includes("徐汇")) score += 1;
  if (`${poi.type || ""}`.includes("科教文化")) score += 1;
  return score;
};

const pickPoi = (row, pois) => {
  if (!pois?.length) return undefined;
  const [best] = [...pois].map((poi) => ({ poi, score: scorePoi(row, poi) })).sort((a, b) => b.score - a.score);
  if (!best || best.score < 8) return undefined;
  return best;
};

const formatDistance = (meters) => {
  const value = Number(meters);
  if (!Number.isFinite(value)) return "";
  return `高德直线约 ${(value / 1000).toFixed(value >= 10000 ? 1 : 2)} km`;
};

const campusKey = (row) => [row["性质"], row["幼儿园"], row["园区/分园"], row["地址"]].join("|");

await fs.mkdir(path.dirname(outputPath), { recursive: true });

const rows = toObjects(parseCsv(await fs.readFile(csvPath, "utf8")));
const existing = await fs.readFile(outputPath, "utf8").then(JSON.parse).catch(() => ({}));

const officeSearch = await textSearch(officeKeyword);
const officePoi = officeSearch.pois?.[0];
if (!officePoi?.location) throw new Error("未能定位西岸网易研发中心。");

const output = { ...existing };
let updated = 0;
let failed = 0;

for (const row of rows) {
  const keyForRow = campusKey(row);
  const searchKeywords = `上海市徐汇区 ${row["幼儿园"]} ${row["园区/分园"]} ${row["地址"]}`;
  try {
    const search = await textSearch(searchKeywords);
    const picked = pickPoi(row, search.pois);
    if (!picked?.poi?.location) {
      failed += 1;
      output[keyForRow] = {
        query: searchKeywords,
        status: "not_found",
        updatedAt: new Date().toISOString(),
      };
      continue;
    }
    const poi = picked.poi;
    const dist = await distance(poi.location, officePoi.location);
    const meters = dist.results?.[0]?.distance || "";
    output[keyForRow] = {
      query: searchKeywords,
      status: "matched",
      matchScore: picked.score,
      poiId: poi.id || "",
      poiName: poi.name || "",
      poiAddress: Array.isArray(poi.address) ? poi.address.join("；") : `${poi.address || ""}`,
      poiDistrict: poi.adname || "",
      poiType: poi.type || "",
      tel: Array.isArray(poi.tel) ? poi.tel.join(" / ") : `${poi.tel || ""}`,
      location: poi.location,
      officeName: officePoi.name || "西岸网易研发中心",
      officeLocation: officePoi.location,
      officeDistanceMeters: meters,
      officeDistanceText: formatDistance(meters),
      updatedAt: new Date().toISOString(),
    };
    updated += 1;
  } catch (error) {
    failed += 1;
    output[keyForRow] = {
      ...(output[keyForRow] || {}),
      query: searchKeywords,
      status: "error",
      error: error.message,
      updatedAt: new Date().toISOString(),
    };
  }
}

await fs.writeFile(outputPath, `${JSON.stringify(output, null, 2)}\n`, "utf8");
console.log(JSON.stringify({
  rows: rows.length,
  updated,
  failed,
  office: {
    name: officePoi.name,
    address: officePoi.address,
    location: officePoi.location,
  },
  outputPath,
}, null, 2));
