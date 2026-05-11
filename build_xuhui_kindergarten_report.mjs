import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = path.join(process.cwd(), "outputs");

const pdfSource = "本地PDF：/Users/caofengyang/Downloads/徐汇区幼儿园对口地区.pdf（2026年徐汇区公办幼儿园对口居委一览表及计划班级数）";
const publicListSource = "上海本地宝《上海市徐汇区公办幼儿园名单一览》，内容来源标注为上海学前教育网，2024-03-25：https://m.sh.bendibao.com/edu/284155.html";
const privateListSource = "上海本地宝《上海市徐汇区民办幼儿园名单一览》，内容来源标注为上海学前教育网，2024-03-25：https://sh.bendibao.com/edu/2024325/284155_2.shtm";
const planSource = "上海本地宝《2026徐汇区幼儿园招生对口地段及招生计划班级数》，2026-04-15：https://m.sh.bendibao.com/edu/305291.html";
const minhangPublicListSource = "上海本地宝《上海市闵行区公办幼儿园名单一览》，内容来源标注为上海学前教育网，2024-03-25：https://sh.bendibao.com/edu/2024325/284200.shtm";
const minhangPrivateListSource = "上海本地宝《上海市闵行区民办幼儿园名单查询》，内容来源标注为上海学前教育网，2024-03-25：https://sh.bendibao.com/edu/2024325/284200_2.shtm";
const pudongPublicListSource = "上海本地宝《浦东新区幼儿园名单+电话查询》，内容来源标注为上海学前教育网，2024-03-25：https://sh.bendibao.com/edu/2024325/284151.shtm";
const pudongPrivateListSource = "上海本地宝《浦东新区民办幼儿园名单一览》，内容来源标注为上海学前教育网，2024-03-25：https://sh.bendibao.com/edu/2024325/284151_2.shtm";
const shanghaiAdmissionPolicySource = "上海市教委《上海市教育委员会关于做好2026年上海市学前教育阶段适龄幼儿入园工作的通知》，2026-04-09：https://edu.sh.gov.cn/xxgk2_zdgz_rxgkyzs_01/20260409/9a3caeed9a0e41c898149e46af9d9203.html";
const xuhuiOfficial2026Source = "上海市人民政府/徐汇区教育局《2026年徐汇区幼儿园招生工作方案》，2026-04-15：https://www.shanghai.gov.cn/xhqxqjy/20260421/c83b6efb6b4c43cb9071399f40215383.html";
const xuhuiOfficial2025Source = "上海市人民政府/徐汇区教育局《2025年徐汇区幼儿园招生工作方案》，2025-04-15：https://www.shanghai.gov.cn/xhqxqjy/20250417/85cd0238c43d49d5ae72cfbe65544a0b.html";
const minhangAdmissionPolicySource = "上观新闻/今日闵行《2026年闵行区学前教育阶段适龄幼儿入园工作各类问题解答》，2026-04-15：https://www.shobserver.cn/sgh/detail?id=1735182";
const pudongAdmissionPolicySource = "浦东新区人民政府《浦东新区教育局关于2026年本区学前教育阶段适龄幼儿入园工作的实施方案》，2026-04-15：https://www.pudong.gov.cn/zwgk/xqjy-jyjzdgz/2026/105/354450.html";

const districtKindergartenSourceDir = path.join(process.cwd(), "data", "district_kindergarten_sources");
const readJsonFile = async (filePath, fallback) => {
  try {
    return JSON.parse(await fs.readFile(filePath, "utf8"));
  } catch {
    return fallback;
  }
};
const xuhuiDistrictSource = await readJsonFile(path.join(districtKindergartenSourceDir, "xuhui.json"), { district: "徐汇", publicSchools: [], publicCampuses: [], publicCampusPhones: [], privateKindergartens: [] });
const minhangDistrictSource = await readJsonFile(path.join(districtKindergartenSourceDir, "minhang.json"), { district: "闵行", publicCampuses: [], privateKindergartens: [] });
const pudongDistrictSource = await readJsonFile(path.join(districtKindergartenSourceDir, "pudong.json"), { district: "浦东", publicCampuses: [], privateKindergartens: [] });
const districtKindergartenSources = [xuhuiDistrictSource, minhangDistrictSource, pudongDistrictSource];

const schools = xuhuiDistrictSource.publicSchools || [];

const campusRows = (xuhuiDistrictSource.publicCampuses || []).map((item) => [item.schoolId, item.campus, item.address, item.confidence, item.note]);

const officeLocation = {
  name: "西岸网易研发中心",
  address: "上海市徐汇区西岸网易研发中心",
  location: "待高德MCP查询",
};

const pendingAmapValue = "待高德MCP查询";
const amapEnrichmentPath = path.join(process.cwd(), "data", "amap_enrichment.json");
let amapEnrichmentByKey = {};
try {
  amapEnrichmentByKey = JSON.parse(await fs.readFile(amapEnrichmentPath, "utf8"));
} catch {
  amapEnrichmentByKey = {};
}

const beikeRentalSchemaPath = path.join(process.cwd(), "data", "beike_rental_filter_schema.json");
let beikeRentalSchema = {};
try {
  beikeRentalSchema = JSON.parse(await fs.readFile(beikeRentalSchemaPath, "utf8"));
} catch {
  beikeRentalSchema = {};
}

const districtLandingProfilesPath = path.join(process.cwd(), "data", "district_landing_profiles.json");
let districtLandingProfiles = { profiles: [] };
try {
  districtLandingProfiles = JSON.parse(await fs.readFile(districtLandingProfilesPath, "utf8"));
} catch {
  districtLandingProfiles = { profiles: [] };
}

const amapItemKey = ({ nature, name, campus, address }) => [nature, name, campus, address].join("|");
const getAmapEnrichment = (item) => {
  const enrichment = amapEnrichmentByKey[amapItemKey(item)] || {};
  return enrichment.status === "matched" ? enrichment : {};
};

const publicSchoolLevelById = new Map([
  [1, "一级"],
  [2, "示范园"],
  [3, "一级"],
  [4, "二级"],
  [5, "二级"],
  [6, "二级"],
  [7, "二级"],
  [8, "二级"],
  [9, "一级"],
  [10, "一级"],
  [11, "一级"],
  [12, "二级"],
  [13, "二级"],
  [14, "示范园"],
  [15, "一级"],
  [16, "一级"],
  [17, "一级"],
  [18, "二级"],
  [19, "二级"],
  [20, "一级"],
  [21, "一级"],
  [22, "一级"],
  [23, "一级"],
  [24, "示范园"],
  [25, "二级"],
  [26, "一级"],
  [27, "一级"],
  [28, "示范园"],
  [29, "一级"],
  [30, "一级"],
  [31, "一级"],
  [32, "一级"],
  [33, "一级"],
  [34, "一级"],
  [35, "一级"],
  [36, "一级"],
  [37, "示范园"],
  [38, "示范园"],
  [39, "示范园"],
  [40, "示范园"],
  [41, "示范园"],
  [42, "示范园"],
]);

const publicCampusPhoneByKey = new Map((xuhuiDistrictSource.publicCampusPhones || []).map((item) => [`${item.schoolId}|${item.campus}`, item.phone]));

const privateCampusRows = xuhuiDistrictSource.privateKindergartens || [];



const externalCampusRows = [minhangDistrictSource, pudongDistrictSource]
  .flatMap((source) => [...(source.publicCampuses || []), ...(source.privateKindergartens || [])]
    .map((item) => ({ ...item, district: item.district || source.district })));
const externalCampusCounts = externalCampusRows.reduce((counts, item) => {
  const prefix = item.district === "闵行" ? "minhang" : "pudong";
  const suffix = item.nature === "公办" ? "Public" : item.nature === "民办" ? "Private" : "Other";
  const key = `${prefix}${suffix}`;
  counts[key] = (counts[key] || 0) + 1;
  return counts;
}, {});

const beikeBoardSlugFromArea = (district, area = "") => {
  const text = `${area}`;
  const rules = district === "闵行"
    ? [
        [/春申/, "chunshen"],
        [/古美/, "gumei"],
        [/梅陇/, "meilong"],
        [/七宝/, "qibao"],
        [/华漕/, "huacao"],
        [/虹桥|龙柏/, "longbai"],
        [/莘庄/, "xinzhuangnanguangchang"],
        [/金汇/, "jinhui"],
        [/浦江/, "pujiang1"],
        [/颛桥/, "zhuanqiao"],
        [/马桥/, "maqiao"],
        [/吴泾/, "wujing"],
        [/江川|老闵行/, "laominhang"],
        [/航华/, "hanghua"],
      ]
    : [
        [/花木/, "huamu"],
        [/联洋/, "lianyang"],
        [/周浦/, "zhoupu"],
        [/唐镇/, "tangzhen"],
        [/三林/, "sanlin"],
        [/张江/, "zhangjiang"],
        [/金桥/, "jinqiao"],
        [/洋泾/, "yangjing"],
        [/世博/, "shibo"],
        [/康桥/, "kangqiao"],
        [/川沙/, "chuansha"],
        [/曹路/, "caolu"],
        [/惠南/, "huinan"],
        [/航头/, "hangtou"],
        [/新场/, "xinchang"],
        [/高桥|外高桥/, "waigaoqiao"],
        [/临港/, "lingangxincheng"],
        [/北蔡/, "beicai"],
        [/陆家嘴/, "lujiazui"],
        [/潍坊/, "weifang"],
      ];
  const districtSlug = district === "闵行" ? "minhang" : district === "浦东" ? "pudong" : "xuhui";
  return rules.find(([pattern]) => pattern.test(text))?.[1] || districtSlug;
};

const externalCampusCounters = new Map();

const byId = new Map(schools.map((school) => [school.id, school]));
const mapSearch = (name, campus, address, district = "徐汇区") => {
  const districtText = district.endsWith("区") || district.endsWith("新区") ? district : `${district}区`;
  const query = `上海市${districtText} ${name} ${campus} ${address}`;
  return `https://ditu.amap.com/search?query=${encodeURIComponent(query)}`;
};

const getAdmissionType = (committee) => {
  if (committee.includes("区域内自主招生")) return "区域自主";
  if (committee.includes("扩招")) return "扩招";
  return "固定对口";
};

const getToddlerMode = (toddler) => {
  if (`${toddler}`.includes("混龄")) return "混龄式招生";
  if (/^\d+$/.test(`${toddler}`)) return "明确托班";
  return "未列明";
};

const inferPrivateArea = (address) => {
  if (/龙吴|华泾|喜泰|龙吟|华发|漓江山水/.test(address)) return "华泾/龙吴";
  if (/梅陇|上中西|老沪闵|金塘|罗城|百色|汇城/.test(address)) return "长桥/梅陇";
  if (/田林|柳州|宜山|虹漕|虹梅|桂平|桂林西|康健|百花|古宜|杨家桥|浦北/.test(address)) return "田林/康健";
  if (/龙华|机场|龙恒|龙兰|东泉|龙水|宛南|云锦|丰谷/.test(address)) return "龙华/滨江";
  if (/永嘉|衡山|安福|五原|华亭|太原|东湖|复兴西|桃江/.test(address)) return "衡复/湖南";
  if (/吴兴|天钥|零陵|斜土|大木桥|小木桥|蒲江塘|宛平南/.test(address)) return "徐家汇/枫林";
  if (/古羊|淮海西/.test(address)) return "徐家汇/虹桥";
  if (/漕泾/.test(address)) return "漕溪/龙漕";
  return "民办/待归类";
};

const publicCampusItems = campusRows.map(([id, campus, address, confidence, note]) => {
  const school = byId.get(id);
  const admissionType = getAdmissionType(school.committee);
  const toddlerMode = getToddlerMode(school.toddler);
  const needsConfirm = confidence === "B" || admissionType !== "固定对口" || note;
  const schoolLevel = publicSchoolLevelById.get(id) || "待核验";
  const phone = publicCampusPhoneByKey.get(`${id}|${campus}`) || "待电话确认";
  const amap = getAmapEnrichment({ nature: "公办", name: school.name, campus, address });
  const searchText = [
    id,
    "公办",
    schoolLevel,
    school.name,
    school.area,
    campus,
    address,
    phone,
    amap.officeDistanceText || pendingAmapValue,
    school.toddler,
    school.small,
    school.committee,
    confidence,
    note,
    admissionType,
    toddlerMode,
  ].join(" ");

  return {
    id,
    district: "徐汇",
    nature: "公办",
    category: "公办",
    level: schoolLevel,
    name: school.name,
    area: school.area,
    campus,
    address,
    mapUrl: mapSearch(school.name, campus, address, "徐汇区"),
    phone,
    officeDistance: amap.officeDistanceText || pendingAmapValue,
    amapPoiName: amap.poiName || pendingAmapValue,
    amapPoiAddress: amap.poiAddress || pendingAmapValue,
    amapLocation: amap.location || pendingAmapValue,
    toddler: school.toddler,
    small: school.small,
    committee: school.committee,
    confidence,
    note,
    source: id === 7 ? `${planSource}；复旦附属徐汇实验幼儿园公开资料/上哪学/021school；${publicListSource}` : `${planSource}；${publicListSource}`,
    admissionType,
    toddlerMode,
    needsConfirm,
    searchText,
  };
});

const privateCampusItems = privateCampusRows.map((item, index) => {
  const campus = item.campus || "本部";
  const id = `M${index + 1}`;
  const area = inferPrivateArea(item.address);
  const amap = getAmapEnrichment({ nature: "民办", name: item.name, campus, address: item.address });
  const searchText = [
    id,
    "民办",
    item.category || "民办（类型待核验）",
    item.level,
    item.name,
    area,
    campus,
    item.address,
    item.phone,
    amap.officeDistanceText || pendingAmapValue,
    "民办招生",
    "待确认",
    "民办招生范围需电话确认",
  ].join(" ");

  return {
    id,
    district: "徐汇",
    nature: "民办",
    category: item.category || "民办（类型待核验）",
    level: item.level,
    name: item.name,
    area,
    campus,
    address: item.address,
    mapUrl: mapSearch(item.name, campus, item.address, "徐汇区"),
    phone: item.phone,
    officeDistance: amap.officeDistanceText || pendingAmapValue,
    amapPoiName: amap.poiName || pendingAmapValue,
    amapPoiAddress: amap.poiAddress || pendingAmapValue,
    amapLocation: amap.location || pendingAmapValue,
    toddler: "待电话确认",
    small: "待电话确认",
    committee: "民办招生范围、托班、小班名额、收费和材料要求需电话确认。",
    confidence: "A",
    note: "民办/私立招生条件、收费、名额每年可能变化，报名前必须电话核验。",
    source: item.source || privateListSource,
    admissionType: "民办招生",
    toddlerMode: "待确认",
    needsConfirm: true,
    searchText,
  };
});

const externalCampusItems = externalCampusRows.map((item) => {
  const districtCode = item.district === "闵行" ? "MH" : "PD";
  const nextIndex = (externalCampusCounters.get(districtCode) || 0) + 1;
  externalCampusCounters.set(districtCode, nextIndex);
  const id = `${districtCode}${nextIndex}`;
  const campus = item.campus || "本部";
  const boardSlug = item.boardSlug || beikeBoardSlugFromArea(item.district, item.area);
  const amap = getAmapEnrichment({ nature: item.nature, name: item.name, campus, address: item.address });
  const admissionType = item.nature === "公办" ? "政策待核验" : "民办招生";
  const searchText = [
    id,
    item.district,
    item.nature,
    item.category,
    item.level,
    item.name,
    item.area,
    campus,
    item.address,
    item.phone,
    item.source,
    admissionType,
  ].join(" ");

  return {
    id,
    district: item.district,
    nature: item.nature,
    category: item.category,
    level: item.level,
    name: item.name,
    area: item.area,
    campus,
    address: item.address,
    mapUrl: mapSearch(item.name, campus, item.address, item.district === "浦东" ? "浦东新区" : `${item.district}区`),
    phone: item.phone,
    officeDistance: amap.officeDistanceText || pendingAmapValue,
    amapPoiName: amap.poiName || pendingAmapValue,
    amapPoiAddress: amap.poiAddress || pendingAmapValue,
    amapLocation: amap.location || pendingAmapValue,
    toddler: "待电话确认",
    small: "待电话确认",
    committee: item.nature === "公办" ? "非徐汇区对口/招生范围需按所在区当年政策和居住地址核验。" : "民办招生范围、托班、小班名额、收费和材料要求需电话确认。",
    confidence: "B",
    note: item.note || "跨区基础数据来自公开名单；距公司、实际招生条件、收费和名额需继续用高德与电话核验。",
    source: item.source,
    admissionType,
    toddlerMode: "待确认",
    needsConfirm: true,
    boardSlug,
    searchText,
  };
});

const campusItems = [...publicCampusItems, ...privateCampusItems, ...externalCampusItems];
const standardizedDatasetDir = path.join(process.cwd(), "data", "kindergartens");
const standardizedDatasetPath = path.join(standardizedDatasetDir, "kindergarten_dataset.json");
const standardizedXuhuiPath = path.join(standardizedDatasetDir, "xuhui_kindergartens.json");
const standardizedKindergartenRows = campusItems.map((item) => ({
  datasetVersion: "2026.05.09",
  district: item.district,
  area: item.area,
  name: item.name,
  campus: item.campus,
  nature: item.nature,
  category: item.category,
  level: item.level,
  address: item.address,
  phone: item.phone,
  admission: {
    type: item.admissionType,
    toddlerPlan: item.toddler,
    smallClassPlan: item.small,
    scope: item.committee,
    toddlerMode: item.toddlerMode,
  },
  geo: {
    amapPoiName: item.amapPoiName,
    amapPoiAddress: item.amapPoiAddress,
    amapLocation: item.amapLocation,
    officeDistance: item.officeDistance,
    mapUrl: item.mapUrl,
  },
  rental: {
    beikeBoardSlug: item.boardSlug || beikeBoardSlugFromArea(item.district, item.area),
  },
  quality: {
    confidence: item.confidence,
    needsConfirm: Boolean(item.needsConfirm),
    note: item.note,
  },
  source: item.source,
}));

const standardizedDataset = {
  schemaVersion: "1.0",
  updatedAt: "2026-05-09",
  scope: ["徐汇", "闵行", "浦东"],
  officeLocation,
  rowCount: standardizedKindergartenRows.length,
  counts: {
    xuhui: standardizedKindergartenRows.filter((item) => item.district === "徐汇").length,
    minhang: standardizedKindergartenRows.filter((item) => item.district === "闵行").length,
    pudong: standardizedKindergartenRows.filter((item) => item.district === "浦东").length,
  },
  rows: standardizedKindergartenRows,
};

await fs.mkdir(standardizedDatasetDir, { recursive: true });
await fs.writeFile(standardizedDatasetPath, `${JSON.stringify(standardizedDataset, null, 2)}\n`);
await fs.writeFile(standardizedXuhuiPath, `${JSON.stringify({
  schemaVersion: "1.0",
  updatedAt: standardizedDataset.updatedAt,
  district: "徐汇",
  rowCount: standardizedDataset.counts.xuhui,
  sources: [pdfSource, planSource, publicListSource, privateListSource],
  rows: standardizedKindergartenRows.filter((item) => item.district === "徐汇"),
}, null, 2)}\n`);

const amapMatchedCount = campusItems.filter((item) => item.officeDistance !== pendingAmapValue).length;
const amapAddressGeocodeCount = campusItems.filter((item) => item.amapPoiName === "高德地址坐标（非POI精确匹配）").length;
const amapUnmatchedItems = campusItems.filter((item) => item.officeDistance === pendingAmapValue);
const amapMatchRate = `${Math.round((amapMatchedCount / campusItems.length) * 100)}%`;

const campusHeader = ["编号", "区", "性质", "办园类型", "办园等级", "幼儿园", "片区", "园区/分园", "地址", "高德POI名称", "高德POI地址", "高德经纬度", `距${officeLocation.name}`, "高德搜索链接", "联系电话", "托班计划", "小班计划", "对口居委/招生范围", "招生类型", "置信度", "备注", "主要来源"];
const campusData = campusItems.map((item) => [
  item.id,
  item.district,
  item.nature,
  item.category,
  item.level,
  item.name,
  item.area,
  item.campus,
  item.address,
  item.amapPoiName,
  item.amapPoiAddress,
  item.amapLocation,
  item.officeDistance,
  item.mapUrl,
  item.phone,
  item.toddler,
  item.small,
  item.committee,
  item.admissionType,
  item.confidence,
  item.note,
  item.source,
]);

const schoolHeader = ["编号", "幼儿园", "片区", "托班计划", "小班计划", "园区数", "对口居委/招生范围", "高德检索方式", "备注"];
const schoolData = schools.map((school) => [
  school.id,
  school.name,
  school.area,
  school.toddler,
  school.small,
  campusRows.filter(([id]) => id === school.id).length,
  school.committee,
  "优先用“上海市徐汇区 + 园区地址”，再用“幼儿园名 + 园区名”核验。",
  ["区域内自主招生", "扩招"].some((word) => school.committee.includes(word)) ? "招生不是普通固定居委一一对应，需结合当年简章和电话确认。" : "",
]);

const areaStats = [...schools.reduce((map, school) => {
  const item = map.get(school.area) || { area: school.area, schools: 0, campuses: 0, small: 0, toddlerNumeric: 0, mixed: 0 };
  item.schools += 1;
  item.campuses += campusRows.filter(([id]) => id === school.id).length;
  item.small += Number(school.small) || 0;
  if (/^\d+$/.test(school.toddler)) item.toddlerNumeric += Number(school.toddler);
  if (`${school.toddler}`.includes("混龄")) item.mixed += 1;
  map.set(school.area, item);
  return map;
}, new Map()).values()]
  .sort((a, b) => b.small - a.small || b.campuses - a.campuses)
  .map((item) => [item.area, item.schools, item.campuses, item.small, item.toddlerNumeric, item.mixed]);

const summaryRows = [
  ["指标", "值", "说明"],
  ["徐汇公办招生主体数", schools.length, "来自2026 PDF。"],
  ["公办园区/分园点位数", publicCampusItems.length, "按公开园部地址拆分；少数历史变更项标B级。"],
  ["民办/私立点位数", privateCampusItems.length, "来自徐汇区民办幼儿园名单公开资料；招生条件、收费和名额需电话确认。"],
  ["闵行/浦东基础点位数", externalCampusItems.length, `闵行${externalCampusCounts.minhangPublic || 0}个公办、${(externalCampusCounts.minhangPrivate || 0) + (externalCampusCounts.minhangOther || 0)}个民办/其他；浦东${externalCampusCounts.pudongPublic || 0}个公办、${(externalCampusCounts.pudongPrivate || 0) + (externalCampusCounts.pudongOther || 0)}个民办/其他。`],
  ["全部园区/点位数", campusData.length, "徐汇公办园区点位 + 徐汇民办/私立点位 + 闵行/浦东公开名单基础点位。"],
  ["高德增强字段", `${amapMatchedCount}/${campusItems.length}`, `已批量补POI/地址坐标、经纬度和到网易上海西岸研发中心的高德直线距离；其中${amapAddressGeocodeCount}个为地址地理编码，非POI精确匹配。`],
  ["小班计划合计", schools.reduce((sum, s) => sum + s.small, 0), "PDF计划班级数合计。"],
  ["托班明确班级数合计", schools.reduce((sum, s) => sum + (/^\d+$/.test(s.toddler) ? Number(s.toddler) : 0), 0), "不含混龄式招生。"],
  ["混龄式招生主体数", schools.filter((s) => `${s.toddler}`.includes("混龄")).length, "混龄式招生不直接等同托班班级数。"],
  ["重点核验", "汇星北园、复旦大学附属徐汇实验幼儿园", "公开资料出现地址/合并后的园部安排差异，择园前应电话确认。"],
];

const markdown = `# 徐汇区幼儿园园区位置与择园参考

## 结论摘要

- 2026 PDF 共列出 ${schools.length} 个公办招生主体，拆分为 ${publicCampusItems.length} 个公办实际园区/分园点位。
- 已补充 ${privateCampusItems.length} 个徐汇区民办/私立点位，包含级别、地址和联系电话；这些点位更适合作为当前暂无居住证情况下的兜底池。
- 已通过高德服务批量增强 POI/地址坐标、经纬度和到${officeLocation.name}的直线距离；地址地理编码不等同于园所 POI 精确匹配，未可靠匹配项保留为待核验。
- 小班计划合计 ${schools.reduce((sum, s) => sum + s.small, 0)} 个班；托班明确计划合计 ${schools.reduce((sum, s) => sum + (/^\d+$/.test(s.toddler) ? Number(s.toddler) : 0), 0)} 个班，另有 ${schools.filter((s) => `${s.toddler}`.includes("混龄")).length} 所为“混龄式招生”。
- 园区数量与小班容量最集中的区域在田林/虹梅/康健/漕河泾、长桥/凌云/梅陇、龙华/滨江与衡复/湖南几条带状居住区。
- 多园区幼儿园不能只搜园名，必须按“园区/分园 + 地址”在高德地图核验；表格中已为每个园区生成高德搜索链接。
- 需要重点电话确认：汇星幼儿园北园公开资料出现“康平路200号”和“华山路1815号”两种说法；复旦大学附属徐汇实验幼儿园为2024年合并组建，平江路17号/32号及原东安路50弄10号关系需要确认2026实际入读园区。

## 片区观察

| 片区 | 招生主体 | 园区点位 | 小班计划 | 明确托班 | 混龄主体 |
|---|---:|---:|---:|---:|---:|
${areaStats.map((row) => `| ${row[0]} | ${row[1]} | ${row[2]} | ${row[3]} | ${row[4]} | ${row[5]} |`).join("\n")}

## 使用建议

1. 先定位你家的居委，再反查对口幼儿园。公办园录取规则里，对口关系和人户一致通常比地图直线距离更关键。
2. 对多园区幼儿园，优先问清楚“2026年小班在哪个园区”，不要只用总园名判断距离。
3. 高德搜索时优先搜园区地址，例如“上海市徐汇区 五原路幼儿园 五原路400号”，再核验POI名称。
4. 托班需求要单独看，“混龄式招生”和普通托班班级数不可直接比较。
5. 接送便利性建议按步行/骑行实际路线判断，尤其关注雨天、老人接送、门口停车和跨主干道情况。

## 主要来源

- ${pdfSource}
- ${planSource}
- ${publicListSource}
- ${privateListSource}
- 复旦大学附属徐汇实验幼儿园、汇星幼儿园等少数变更项参考了园所公开资料、上哪学/021school收录的招生简章信息，并在表中标注为需核验。
`;

const escapeHtml = (value) => String(value ?? "")
  .replaceAll("&", "&amp;")
  .replaceAll("<", "&lt;")
  .replaceAll(">", "&gt;")
  .replaceAll('"', "&quot;");

const campusLookup = new Map(campusItems.map((item) => [amapItemKey(item), item]));
const campusRef = (nature, name, campus, address) => campusLookup.get([nature, name, campus, address].join("|"));
const beikeDefaultTokens = ["rt200600000001", "brp0erp10000", "l2", "l3", "ra4", "ra5", "ie1"];
const beikeBizcircleBySlug = new Map((beikeRentalSchema.filters?.bizcircle?.flatOptions || [])
  .map((option) => [option.slug, option]));
const beikeDistrictSlugByName = new Map([
  ["徐汇", "xuhui"],
  ["闵行", "minhang"],
  ["浦东", "pudong"],
]);
const beikeDistrictLabelBySlug = new Map([
  ["xuhui", "徐汇全区"],
  ["minhang", "闵行全区"],
  ["pudong", "浦东全区"],
]);
const beikeAreaLabel = (slug) => beikeDistrictLabelBySlug.get(slug) || beikeBizcircleBySlug.get(slug)?.label || slug;
const beikeRentalUrl = (slug = "xuhui") => {
  const safeSlug = slug || "xuhui";
  return `https://sh.zu.ke.com/zufang/${safeSlug}/${beikeDefaultTokens.join("")}/?showMore=1`;
};
const beikeRentalConditionText = "整租 / 0-10000元 / 三居或四居+ / 100㎡以上 / 有电梯";
const districtProfiles = districtLandingProfiles.profiles || [];
const districtCounts = campusItems.reduce((map, item) => {
  const current = map.get(item.district) || { total: 0, public: 0, private: 0 };
  current.total += 1;
  if (item.nature === "公办") current.public += 1;
  if (item.nature === "民办") current.private += 1;
  map.set(item.district, current);
  return map;
}, new Map());
const itemDistance = (item) => item?.officeDistance && item.officeDistance !== pendingAmapValue ? item.officeDistance : "高德未匹配，待电话/地图确认";
const poiLine = (item) => item?.amapPoiName && item.amapPoiName !== pendingAmapValue
  ? `${item.amapPoiName}${item.amapPoiAddress && item.amapPoiAddress !== pendingAmapValue ? `；${item.amapPoiAddress}` : ""}`
  : "高德POI未可靠匹配";
const phoneDialValue = (value) => {
  const digits = String(value ?? "").match(/\d+/g)?.join("") || "";
  return digits.length >= 7 ? digits : "";
};
const renderPhoneLinks = (value) => {
  const phone = String(value ?? "待电话确认");
  const parts = phone
    .split(/\s*(?:\/|、|，|,|；|;)\s*/u)
    .filter(Boolean);
  if (!parts.length || phone.includes("待电话确认")) return escapeHtml(phone);
  return parts.map((part) => {
    const dial = phoneDialValue(part);
    return dial ? `<a class="phone-link" href="tel:${escapeHtml(dial)}">${escapeHtml(part)}</a>` : escapeHtml(part);
  }).join('<span class="phone-separator"> / </span>');
};

const decisionRecommendations = [
  {
    rank: "公办 1",
    tag: "首选争取",
    tagClass: "green",
    type: "公办 / 一级",
    item: campusRef("公办", "园南幼儿园", "本部", "龙川北路园南一村27号"),
    rentArea: "园南一村 / 汇成五村 / 上中路",
    rentSlug: "changqiao",
    backupRentSlug: "zhiwuyuan",
    strategy: "公办争取线",
    risk: "暂无居住证时需严格匹配居委、租赁材料和当年招生顺位。",
    judgement: "租房预算、居住环境和公办争取的平衡最好，是当前主线。",
  },
  {
    rank: "公办 2",
    tag: "备选争取",
    tagClass: "blue",
    type: "公办 / 一级",
    item: campusRef("公办", "徐汇实验幼儿园", "本部", "龙瑞路135号"),
    rentArea: "华沁家园 / 华滨家园 / 龙瑞路",
    rentSlug: "huajing",
    backupRentSlug: "xuhui",
    strategy: "公办争取线",
    risk: "华泾/罗秀具体居委边界要反查，不能只按地图距离判断。",
    judgement: "距离公司近于长桥，面积更容易做大，但需严格反查居委。",
  },
  {
    rank: "公办 3",
    tag: "近公司",
    tagClass: "amber",
    type: "公办 / 二级",
    item: campusRef("公办", "阳光幼儿园", "小班部", "龙水南路龙南三村7号"),
    rentArea: "龙南三四村 / 樟树苑 / 龙水南路",
    rentSlug: "longhua",
    backupRentSlug: "xuhuibinjiang",
    strategy: "近公司机会项",
    risk: "龙华/滨江满足 100㎡以上、1 万内、电梯的房源更少。",
    judgement: "高德距离最近，适合捡漏；不要作为唯一租房主线。",
  },
  {
    rank: "公办 4",
    tag: "容量好",
    tagClass: "blue",
    type: "公办 / 示范园",
    item: campusRef("公办", "上海幼儿园", "上中园", "上中路402号"),
    rentArea: "上中路 / 长桥 / 植物园",
    rentSlug: "zhiwuyuan",
    backupRentSlug: "changqiao",
    strategy: "公办争取线",
    risk: "示范园关注度高，扩招和自主口径需以当年简章为准。",
    judgement: "示范园且片区与租房需求匹配，但公办竞争和材料门槛仍高。",
  },
  {
    rank: "公办 5",
    tag: "通勤好",
    tagClass: "amber",
    type: "公办 / 一级",
    item: campusRef("公办", "龙华幼儿园", "龙恒园", "龙华西路31弄15号"),
    rentArea: "龙华西路 / 龙恒路 / 云锦路",
    rentSlug: "longhua",
    backupRentSlug: "xuhuibinjiang",
    strategy: "近公司机会项",
    risk: "预算、房源面积、公办顺位三项压力同时偏高。",
    judgement: "距离公司好，但租金和公办顺位压力都更大，作为机会项。",
  },
  {
    rank: "民办 1",
    tag: "主兜底",
    tagClass: "green",
    type: "普惠民办 / 一级",
    item: campusRef("民办", "汇城苑幼稚园", "本部", "百色路汇城五村75号"),
    rentArea: "汇成五村 / 园南 / 百色路",
    rentSlug: "changqiao",
    backupRentSlug: "zhiwuyuan",
    strategy: "民办兜底线",
    risk: "需电话确认剩余名额、收费、是否接受当前材料状态。",
    judgement: "最适合和园南/长桥公办争取线并行，优先电话确认名额。",
  },
  {
    rank: "民办 2",
    tag: "近公司",
    tagClass: "green",
    type: "民办 / 一级",
    item: campusRef("民办", "胡姬港湾幼儿园", "本部", "丰谷路205弄35号"),
    rentArea: "龙华 / 云锦路 / 丰谷路",
    rentSlug: "longhua",
    backupRentSlug: "xuhuibinjiang",
    strategy: "近公司兜底线",
    risk: "名单地址与高德 POI 地址有差异，必须电话确认实际园区。",
    judgement: "高德距离近；名单地址与高德POI地址有差异，必须电话确认实际园区。",
  },
  {
    rank: "民办 3",
    tag: "长桥兜底",
    tagClass: "blue",
    type: "民办 / 二级",
    item: campusRef("民办", "牛牛幼稚园", "本部", "罗城路700弄95号"),
    rentArea: "罗城路 / 长桥 / 上中西路",
    rentSlug: "changqiao",
    backupRentSlug: "zhiwuyuan",
    strategy: "长桥兜底线",
    risk: "二级民办，需实地看园并确认班额、师资和收费。",
    judgement: "和长桥租房主线兼容，适合作为第二民办兜底。",
  },
  {
    rank: "民办 4",
    tag: "华泾兜底",
    tagClass: "blue",
    type: "普惠民办 / 二级",
    item: campusRef("民办", "凯琴数码幼儿园", "本部", "华泾路995号"),
    rentArea: "华泾 / 华沁 / 华滨 / 罗秀",
    rentSlug: "huajing",
    backupRentSlug: "xuhui",
    strategy: "华泾兜底线",
    risk: "需确认普惠口径、招生名额、收费和可入园时间。",
    judgement: "适合华泾北/罗秀租房线，距离公司中等，需确认名额和收费。",
  },
  {
    rank: "民办 5",
    tag: "康健备选",
    tagClass: "amber",
    type: "民办 / 一级",
    item: campusRef("民办", "杜鹃园幼稚园", "本部", "桂林西街9弄57号"),
    rentArea: "桂林西街 / 康健 / 田林南",
    rentSlug: "kangjian",
    backupRentSlug: "tianlin",
    strategy: "康健备选线",
    risk: "生活成熟但通勤略远，建议作为备选而不是第一主线。",
    judgement: "一级民办，适合作为康健/田林方向的兜底备选。",
  },
];

const rentalBoards = [
  { slug: "changqiao", tag: "首选", title: "长桥", fit: "园南幼儿园、汇城苑、牛牛", note: "预算、面积和材料办理可行性最平衡。" },
  { slug: "zhiwuyuan", tag: "首选备选", title: "植物园", fit: "上海幼儿园上中园、园南幼儿园", note: "靠近上中路/长桥南侧，适合找大户型。" },
  { slug: "huajing", tag: "备选首选", title: "华泾", fit: "徐汇实验幼儿园、凯琴数码幼儿园", note: "100㎡以上房源更友好，通勤仍可接受。" },
  { slug: "longhua", tag: "近公司", title: "龙华", fit: "阳光幼儿园、龙华幼儿园、胡姬港湾", note: "通勤最好，但预算和大户型筛选压力高。" },
  { slug: "kangjian", tag: "生活成熟", title: "康健", fit: "杜鹃园幼稚园", note: "生活便利，适合作为康健/田林方向备选。" },
  { slug: "xuhui", tag: "放宽", title: "徐汇全区", fit: "预算不足时扩大池子", note: "当二级商圈房源过少时，用全区链接补充看房池。" },
];

const architectureReview = {
  updatedAt: "2026-05-09",
  title: "上海家庭第一阶段落地执行方案架构 Review",
  summary: "当前项目已从单一区域择园页面升级为区级落地决策数据产品：以每区统一源数据为输入，串联标准化幼儿园数据集、政策来源、高德 POI/距离、贝壳租房参数、区域路线评分和网页/Excel 输出。",
  modules: [
    {
      name: "数据采集层",
      status: "已扩展",
      input: "每区统一源数据文件、官方/公开名单、徐汇 2026 招生计划、闵行/浦东公开名单、贝壳参数、高德 API 查询结果",
      output: "标准化 JSON、CSV、Excel、HTML",
      risk: "部分公开名单年度口径不同，尤其闵行民办需报名季复核。",
    },
    {
      name: "标准数据层",
      status: "本次优化",
      input: "data/district_kindergarten_sources/xuhui.json、minhang.json、pudong.json",
      output: "data/kindergartens/kindergarten_dataset.json；data/kindergartens/xuhui_kindergartens.json",
      risk: "后续新增区域必须先补标准字段，不应直接写 UI 文案。",
    },
    {
      name: "POI 与通勤层",
      status: "已接入",
      input: "园所名称、园区、地址、办公点坐标",
      output: "高德 POI/地址坐标、经纬度、到西岸网易研发中心直线距离",
      risk: "直线距离不是实际通勤；地址地理编码不等于 POI 精确匹配。",
    },
    {
      name: "租房联动层",
      status: "已结构化",
      input: "贝壳一级区/二级商圈 slug 与固定筛选 token",
      output: "整租、0-10000、三居/四居+、100㎡以上、有电梯的结构化链接",
      risk: "贝壳库存和价格实时变化，页面只保存查询条件，不固化房源。",
    },
    {
      name: "推荐策略层",
      status: "可解释",
      input: "家庭约束、区域画像、公办材料风险、民办兜底、租房可执行性、通勤距离",
      output: "徐汇/闵行/浦东区级择园策略、每区公办争取线、每区民办兜底线、行动清单",
      risk: "不承诺精确录取概率，只做政策顺位和执行可行性分层。",
    },
    {
      name: "展示与发布层",
      status: "已发布",
      input: "标准数据、策略文案、来源和质量状态",
      output: "GitHub Pages HTML、Excel、CSV、JSON 数据文件",
      risk: "静态页面无法自动刷新实时房源，需要手动重新生成和部署。",
    },
  ],
  optimizations: [
    "把徐汇数据从生成脚本迁移到 data/district_kindergarten_sources/xuhui.json，与闵行/浦东保持同一源数据结构。",
    "新增数据集 schema，明确后续新增行政区和模块的字段边界。",
    "新增网页端架构模块和三区择园策略模块，用信息流式卡片展示核心模块、输入输出和风险。",
    "新增架构评审文档和架构 JSON，方便后续迭代时先改数据层，再改策略层，最后改 UI。",
    "保留静态站发布路径，同时发布标准 JSON，外部工具可直接消费数据集。",
  ],
  nextSteps: [
    "把看房记录、电话核验记录、报名材料状态拆成独立 JSON 模块。",
    "把推荐策略权重外置为配置文件，避免硬编码在页面生成脚本里。",
    "把高德增强脚本纳入固定 pipeline，记录 POI 匹配分数和地址编码比例。",
    "后续扩展新区域时，先补区级源数据、贝壳板块、政策来源，再进入标准数据和推荐策略。",
  ],
};

const architectureReviewMarkdown = `# ${architectureReview.title}

更新时间：${architectureReview.updatedAt}

## 总体判断

${architectureReview.summary}

当前架构可以继续支撑“90 天落地执行方案”，但需要坚持一个原则：新增能力先进入标准数据层，再进入策略层，最后进入 UI。不要再把新数据直接写成页面文案或散落在推荐卡片里。

## 当前核心逻辑

${architectureReview.modules.map((item) => `### ${item.name}

- 状态：${item.status}
- 输入：${item.input}
- 输出：${item.output}
- 风险：${item.risk}
`).join("\n")}

## 本次架构优化

${architectureReview.optimizations.map((item) => `- ${item}`).join("\n")}

## 后续扩展建议

${architectureReview.nextSteps.map((item) => `- ${item}`).join("\n")}

## 当前数据规模

- 标准幼儿园点位：${standardizedDataset.rowCount}
- 徐汇：${standardizedDataset.counts.xuhui}
- 闵行：${standardizedDataset.counts.minhang}
- 浦东：${standardizedDataset.counts.pudong}
- 高德覆盖：${amapMatchedCount}/${campusItems.length}
- 地址地理编码：${amapAddressGeocodeCount}
`;

const architectureReviewPath = path.join(process.cwd(), "data", "project_architecture_review.json");
const architectureReviewDocPath = path.join(process.cwd(), "docs", "ARCHITECTURE_REVIEW.md");
await fs.mkdir(path.dirname(architectureReviewDocPath), { recursive: true });
await fs.writeFile(architectureReviewPath, `${JSON.stringify({
  ...architectureReview,
  metrics: {
    kindergartenRows: standardizedDataset.rowCount,
    districtCounts: standardizedDataset.counts,
    amapCoverage: `${amapMatchedCount}/${campusItems.length}`,
    amapAddressGeocodeCount,
    beikeDefaultTokens: beikeDefaultTokens.join(""),
  },
}, null, 2)}\n`);
await fs.writeFile(architectureReviewDocPath, architectureReviewMarkdown);

const renderRentalLinks = (row) => `
        <div class="rent-links">
          <a href="${escapeHtml(beikeRentalUrl(row.rentSlug))}" target="_blank" rel="noopener">推荐商圈：${escapeHtml(beikeAreaLabel(row.rentSlug))}</a>
          ${row.backupRentSlug && row.backupRentSlug !== "xuhui" ? `<a href="${escapeHtml(beikeRentalUrl(row.backupRentSlug))}" target="_blank" rel="noopener">备选商圈：${escapeHtml(beikeAreaLabel(row.backupRentSlug))}</a>` : ""}
          <a href="${escapeHtml(beikeRentalUrl("xuhui"))}" target="_blank" rel="noopener">徐汇全区放宽查找</a>
        </div>
`;

const renderSchoolDecisionCards = (nature) => decisionRecommendations
  .filter((row) => row.item?.nature === nature)
  .map((row) => `
      <article class="school-card">
        <header>
          <div>
            <span class="tag ${row.tagClass}">${escapeHtml(row.rank)} · ${escapeHtml(row.tag)}</span>
            <h3>${escapeHtml(row.item?.name || "")}${row.item?.campus ? ` · ${escapeHtml(row.item.campus)}` : ""}</h3>
          </div>
          <strong>${escapeHtml(itemDistance(row.item))}</strong>
        </header>
        <div class="card-meta">
          <div><span>性质/等级</span><b>${escapeHtml(row.type)}</b></div>
          <div><span>招生类型</span><b>${escapeHtml(row.item?.admissionType || "待确认")}</b></div>
          <div><span>联系电话</span><b>${renderPhoneLinks(row.item?.phone || "待电话确认")}</b></div>
          <div><span>租房板块</span><b>${escapeHtml(row.rentArea)}</b></div>
        </div>
        <p class="address-line">${escapeHtml(row.item?.address || "")}</p>
        <p>${escapeHtml(row.judgement)}</p>
        <div class="risk-box"><strong>${escapeHtml(row.strategy)}</strong><span>${escapeHtml(row.risk)}</span></div>
        ${renderRentalLinks(row)}
      </article>
`).join("");

const renderRentBoardCards = () => rentalBoards.map((board) => `
        <article class="rent-board-card">
          <header><h3>${escapeHtml(board.title)}</h3><span class="tag ${board.slug === "xuhui" ? "amber" : "blue"}">${escapeHtml(board.tag)}</span></header>
          <p><strong>适配园所：</strong>${escapeHtml(board.fit)}</p>
          <p>${escapeHtml(board.note)}</p>
          <a href="${escapeHtml(beikeRentalUrl(board.slug))}" target="_blank" rel="noopener">打开${escapeHtml(beikeAreaLabel(board.slug))}房源</a>
        </article>
`).join("");

const rentalAreaOptions = [
  ["xuhui", "徐汇全区"], ["changqiao", "徐汇 · 长桥"], ["huajing", "徐汇 · 华泾"], ["longhua", "徐汇 · 龙华"], ["kangjian", "徐汇 · 康健"], ["zhiwuyuan", "徐汇 · 植物园"],
  ["minhang", "闵行全区"], ["chunshen", "闵行 · 春申"], ["gumei", "闵行 · 古美"], ["meilong", "闵行 · 梅陇"], ["zhuanqiao", "闵行 · 颛桥"],
  ["pudong", "浦东全区"], ["zhangjiang", "浦东 · 张江"], ["huamu", "浦东 · 花木"], ["sanlin", "浦东 · 三林"], ["kangqiao", "浦东 · 康桥"],
];

const rentalSnapshotCards = [
  {
    title: "徐汇南部看房入口",
    area: "长桥 / 植物园 / 华泾",
    price: "0-10000 元",
    layout: "三居 / 四居+，100㎡以上，有电梯",
    reason: "优先匹配当前预算、面积、老人接送和幼儿园兜底需求。",
    slug: "changqiao",
  },
  {
    title: "闵行生活平衡入口",
    area: "春申 / 古美 / 梅陇",
    price: "0-10000 元",
    layout: "三居 / 四居+，100㎡以上，有电梯",
    reason: "当徐汇大户型压力过高时，用闵行改善居住体验和亲子社区。",
    slug: "chunshen",
  },
  {
    title: "浦东成长空间入口",
    area: "张江 / 花木 / 三林",
    price: "0-10000 元",
    layout: "三居 / 四居+，100㎡以上，有电梯",
    reason: "只在通勤验证通过后进入主线，用于比较新社区和长期空间。",
    slug: "zhangjiang",
  },
];

const renderRentalSnapshotCards = () => rentalSnapshotCards.map((item) => `
        <article class="rental-result-card">
          <a class="rental-photo" href="${escapeHtml(beikeRentalUrl(item.slug))}" target="_blank" rel="noopener" aria-label="打开${escapeHtml(item.title)}贝壳房源">
            <span>贝壳实时房源图</span>
            <small>点击查看图片与库存</small>
          </a>
          <div class="rental-result-body">
            <header><h3>${escapeHtml(item.title)}</h3><span class="tag blue">${escapeHtml(item.area)}</span></header>
            <p>${escapeHtml(item.layout)}</p>
            <div class="rental-meta"><span>${escapeHtml(item.price)}</span><span>整租</span><span>有电梯</span></div>
            <p class="module-reason">${escapeHtml(item.reason)}</p>
            <a href="${escapeHtml(beikeRentalUrl(item.slug))}" target="_blank" rel="noopener">打开贝壳实时结果</a>
          </div>
        </article>
`).join("");

const districtCapacityStats = [...districtCounts.entries()].map(([district, counts]) => ({
  label: district,
  value: counts.total,
  sub: `${counts.public} 公办 / ${counts.private} 民办`,
}));
const areaCapacityStats = areaStats.slice(0, 10).map((row) => ({
  label: row[0],
  value: row[2],
  sub: `${row[1]} 主体 / ${row[3]} 小班`,
}));
const renderBarChart = (items) => {
  const max = Math.max(...items.map((item) => item.value), 1);
  return items.map((item) => `
        <div class="bar-row">
          <span>${escapeHtml(item.label)}</span>
          <div class="bar-track"><i style="width:${Math.max(8, Math.round((item.value / max) * 100))}%"></i></div>
          <b>${escapeHtml(item.value)}</b>
          <small>${escapeHtml(item.sub)}</small>
        </div>
`).join("");
};

const renderArchitectureReviewCards = () => architectureReview.modules.map((item) => `
        <article class="module-item">
          <div class="module-time">${escapeHtml(item.status)}<br>${escapeHtml(item.name)}</div>
          <div class="module-body">
            <h3>${escapeHtml(item.name)}</h3>
            <p><strong>输入：</strong>${escapeHtml(item.input)}</p>
            <p><strong>输出：</strong>${escapeHtml(item.output)}</p>
            <p class="module-reason"><strong>风险：</strong>${escapeHtml(item.risk)}</p>
          </div>
          <div class="module-status"><span class="tag ${item.status === "本次优化" ? "green" : item.status === "已接入" || item.status === "已结构化" ? "blue" : "amber"}">${escapeHtml(item.status)}</span></div>
        </article>
`).join("");

const renderDistrictRouteCards = () => districtProfiles.map((profile) => {
  const counts = districtCounts.get(profile.district) || { total: 0, public: 0, private: 0 };
  return `
        <article class="route-card">
          <header><h3>${escapeHtml(profile.district)} · ${escapeHtml(profile.route)}</h3><span class="tag ${profile.district === "徐汇" ? "green" : profile.district === "闵行" ? "blue" : "amber"}">${escapeHtml(profile.priority)}</span></header>
          <p class="reason">${escapeHtml(profile.summary)}</p>
          <div class="route-metrics">
            <span><b>${counts.total}</b>基础点位</span>
            <span><b>${counts.public}</b>公办</span>
            <span><b>${counts.private}</b>民办</span>
          </div>
          <p><strong>适合：</strong>${escapeHtml(profile.fit)}</p>
          <p><strong>主要风险：</strong>${escapeHtml(profile.risks.join("；"))}</p>
          <p><strong>放弃条件：</strong>${escapeHtml(profile.dropCondition)}</p>
          <div class="rent-links">
            ${profile.recommendedBoards.map((board) => `<a href="${escapeHtml(beikeRentalUrl(board.slug))}" target="_blank" rel="noopener">${escapeHtml(board.label)}房源</a>`).join("")}
          </div>
        </article>
`;
}).join("");

const districtKindergartenStrategyItems = [
  {
    district: "徐汇",
    tag: "通勤保守线",
    boards: ["changqiao", "huajing", "longhua"],
    publicStrategy: "公办按居委和材料顺位争取，优先验证南部板块、区域自主和扩招口径。",
    privateStrategy: "民办/普惠民办作为同步兜底，重点联系长桥、龙华、华泾、康健方向。",
    action: "看房前确认居委和房东材料配合；看房后用居委反查公办园，再电话确认目标园。",
  },
  {
    district: "闵行",
    tag: "生活平衡线",
    boards: ["chunshen", "gumei", "meilong"],
    publicStrategy: "公办先按租房地址所在街镇和当年梯队排序核验，不把近公司距离当作录取依据。",
    privateStrategy: "民办和中外合作点位数量更适合做兜底池，但年度名单和收费必须电话确认。",
    action: "优先比较春申、古美、梅陇、颛桥，筛 100㎡以上电梯房源，再按小区所属居委反查园所。",
  },
  {
    district: "浦东",
    tag: "成长空间线",
    boards: ["zhangjiang", "huamu", "sanlin"],
    publicStrategy: "浦东板块跨度大，先分清按地段招生和全区招生，再判断来沪人员材料是否可进入统筹。",
    privateStrategy: "民办兜底可用，但必须叠加通勤实测，避免园所可行、上班不可行。",
    action: "只在张江、花木、三林、周浦、康桥等板块通勤验证通过后进入主线。",
  },
];

const renderDistrictKindergartenStrategyCards = () => districtKindergartenStrategyItems.map((item) => {
  const counts = districtCounts.get(item.district) || { total: 0, public: 0, private: 0 };
  return `
        <article class="route-card">
          <header><h3>${escapeHtml(item.district)}怎么选</h3><span class="tag ${item.district === "徐汇" ? "green" : item.district === "闵行" ? "blue" : "amber"}">${escapeHtml(item.tag)}</span></header>
          <div class="route-metrics">
            <span><b>${counts.total}</b>点位</span>
            <span><b>${counts.public}</b>公办</span>
            <span><b>${counts.private}</b>民办</span>
          </div>
          <p><strong>公办策略：</strong>${escapeHtml(item.publicStrategy)}</p>
          <p><strong>民办策略：</strong>${escapeHtml(item.privateStrategy)}</p>
          <p><strong>下一步：</strong>${escapeHtml(item.action)}</p>
          <div class="rent-links">
            ${item.boards.map((slug) => `<a href="${escapeHtml(beikeRentalUrl(slug))}" target="_blank" rel="noopener">${escapeHtml(beikeAreaLabel(slug))}房源</a>`).join("")}
          </div>
        </article>
`;
}).join("");

const landingScoreItems = [
  ["通勤", "高", "30-45分钟内优先，直线距离只做预筛。"],
  ["幼儿园可落位", "高", "公办争取与民办兜底必须同步。"],
  ["租房可执行", "极高", "能签、能搬、能放家具比理想小区更重要。"],
  ["材料可办理", "极高", "房东配合居住登记/备案是硬条件。"],
  ["居住品质", "中高", "避免过度压抑，但不追求一步到位。"],
  ["社区亲子环境", "中高", "有孩子生态和老人接送便利优先。"],
  ["预算压力", "高", "1万内、100㎡以上、电梯是硬筛选。"],
  ["后续迁移弹性", "中", "当前房子是第一阶段，不是终身方案。"],
];

const renderLandingScoreCards = () => landingScoreItems.map(([name, weight, note]) => `
        <article class="logic-card"><strong>${escapeHtml(name)} · ${escapeHtml(weight)}</strong><span>${escapeHtml(note)}</span></article>
`).join("");

const admissionHistoryItems = [
  {
    title: "徐汇 2025-2026 规则连续性",
    level: "可用于概率分层",
    text: "2025 和 2026 方案均体现相对就近、户籍优先、人户一致优先；外省市户籍在保障本区户籍基础上，按居住证、社保、居住情况等排序统筹。",
    source: xuhuiOfficial2026Source,
  },
  {
    title: "徐汇 2026 对口与计划",
    level: "可用于容量判断",
    text: "42 个公办招生主体、81 个公办园区点位、小班计划和对口居委用于判断板块容量，但不能直接换算为个人录取概率。",
    source: planSource,
  },
  {
    title: "闵行候选路线",
    level: "需电话核验",
    text: "闵行公开口径强调按报名条件排序、验证通过后进入录取阶段，溢出由区教育行政部门统筹分流；本页已接入公开名单基础点位，但年度招生口径仍需报名季核验。",
    source: minhangAdmissionPolicySource,
  },
  {
    title: "浦东候选路线",
    level: "需单独评估",
    text: "浦东有按地段招生和全区招生口径，来沪人员随迁子女可按条件统筹；板块跨度大，必须同时验证通勤和招生类别。",
    source: pudongAdmissionPolicySource,
  },
];

const renderAdmissionHistoryCards = () => admissionHistoryItems.map((item) => `
        <article class="logic-card">
          <strong>${escapeHtml(item.title)} · ${escapeHtml(item.level)}</strong>
          <span>${escapeHtml(item.text)}</span>
          <a href="${escapeHtml(item.source.match(/https?:\/\/\S+/)?.[0] || "#")}" target="_blank" rel="noopener">查看来源</a>
        </article>
`).join("");

const renderExternalCandidateCards = (district) => externalCampusItems
  .filter((item) => item.district === district)
  .slice(0, 6)
  .map((item) => `
        <article class="candidate-card">
          <header><h3>${escapeHtml(item.name)}</h3><span class="tag ${item.nature === "公办" ? "green" : "amber"}">${escapeHtml(item.nature)} / ${escapeHtml(item.level)}</span></header>
          <p>${escapeHtml(item.area)} · ${escapeHtml(item.address)}</p>
          <p>${renderPhoneLinks(item.phone)}</p>
          <div class="rent-links"><a href="${escapeHtml(beikeRentalUrl(item.boardSlug || beikeDistrictSlugByName.get(item.district)))}" target="_blank" rel="noopener">查看${escapeHtml(beikeAreaLabel(item.boardSlug || beikeDistrictSlugByName.get(item.district)))}房源</a><a href="${escapeHtml(beikeRentalUrl(beikeDistrictSlugByName.get(item.district) || "xuhui"))}" target="_blank" rel="noopener">${escapeHtml(item.district)}全区放宽查找</a><a href="${escapeHtml(item.mapUrl)}" target="_blank" rel="noopener">打开高德</a></div>
        </article>
`).join("");

const unmatchedPreview = amapUnmatchedItems.slice(0, 12)
  .map((item) => `${item.nature}-${item.name}${item.campus === "本部" ? "" : ` ${item.campus}`}`)
  .join("、");

const html = `<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>上海家庭第一阶段落地执行方案</title>
  <link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'%3E%3Crect width='64' height='64' rx='12' fill='%23235c9f'/%3E%3Cpath d='M16 42h32M20 42V24l12-8 12 8v18M28 42V30h8v12' stroke='white' stroke-width='4' fill='none' stroke-linecap='round' stroke-linejoin='round'/%3E%3C/svg%3E">
  <style>
    :root {
      color-scheme: light;
      --bg: #f4f6f8;
      --paper: #ffffff;
      --ink: #111827;
      --soft: #667085;
      --line: #e4e7ec;
      --blue: #2563eb;
      --green: #047857;
      --amber: #b45309;
      --red: #b42318;
      --rail: #f2f4f7;
      --shadow: 0 10px 30px rgba(17, 24, 39, 0.06);
    }
    * { box-sizing: border-box; }
    html { scroll-behavior: smooth; }
    body {
      margin: 0;
      overflow-x: hidden;
      background: var(--bg);
      color: var(--ink);
      font: 14px/1.58 -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
    }
    a { color: var(--blue); text-decoration: none; }
    a:hover { text-decoration: underline; }
    .shell { max-width: 1480px; margin: 0 auto; padding: 0 24px; }
    .topbar {
      position: sticky;
      top: 0;
      z-index: 20;
      border-bottom: 1px solid rgba(228, 231, 236, 0.95);
      background: rgba(255, 255, 255, 0.94);
      backdrop-filter: blur(12px);
    }
    .topbar .shell {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 18px;
      min-height: 64px;
    }
    .brand { font-weight: 900; font-size: 16px; letter-spacing: 0; }
    .nav { display: flex; gap: 8px; color: var(--soft); font-size: 13px; white-space: nowrap; }
    .nav a { color: inherit; padding: 7px 10px; border-radius: 999px; }
    .nav a:hover { background: var(--rail); color: var(--ink); text-decoration: none; }
    .hero {
      background: #fff;
      border-bottom: 1px solid var(--line);
    }
    .hero-grid {
      display: grid;
      grid-template-columns: minmax(0, 1fr) 420px;
      gap: 44px;
      align-items: center;
      min-height: 390px;
      padding: 42px 24px 38px;
    }
    h1 {
      margin: 0;
      font-size: clamp(34px, 5vw, 58px);
      line-height: 1.06;
      letter-spacing: 0;
    }
    .lead {
      max-width: 760px;
      margin: 22px 0 0;
      color: var(--soft);
      font-size: 17px;
      line-height: 1.7;
    }
    .hero-actions { display: flex; flex-wrap: wrap; gap: 12px; margin-top: 30px; }
    .button {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-height: 42px;
      padding: 0 16px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--paper);
      color: var(--ink);
      font-weight: 700;
      font-size: 14px;
      box-shadow: 0 4px 14px rgba(23, 32, 51, 0.05);
    }
    .button.primary { background: var(--blue); border-color: var(--blue); color: #fff; }
    .scoreboard {
      background: var(--paper);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
      padding: 20px;
    }
    .score-row {
      display: grid;
      grid-template-columns: 1fr auto;
      gap: 12px;
      padding: 14px 0;
      border-bottom: 1px solid var(--line);
    }
    .score-row:last-child { border-bottom: 0; }
    .score-row span { color: var(--soft); }
    .score-row strong { font-size: 28px; line-height: 1; }
    main { padding: 28px 0 56px; }
    section { scroll-margin-top: 82px; }
    .app-layout {
      display: grid;
      grid-template-columns: 280px minmax(0, 1fr);
      gap: 18px;
      align-items: start;
    }
    .module-sidebar {
      position: sticky;
      top: 82px;
      display: grid;
      gap: 10px;
      max-height: calc(100vh - 100px);
      overflow: auto;
      padding: 4px 0;
    }
    .side-module {
      display: grid;
      gap: 8px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--paper);
      padding: 14px;
      color: var(--ink);
    }
    .side-module:hover { text-decoration: none; border-color: #bdd3ea; box-shadow: var(--shadow); }
    .side-module strong { font-size: 15px; }
    .side-module span { color: var(--soft); font-size: 12px; line-height: 1.45; }
    .side-module small { color: var(--blue); font-weight: 800; }
    .module-content { min-width: 0; }
    .section-title {
      display: flex;
      align-items: end;
      justify-content: space-between;
      gap: 18px;
      margin: 34px 0 16px;
    }
    h2 { margin: 0; font-size: 26px; line-height: 1.2; letter-spacing: 0; }
    .section-title p { max-width: 660px; margin: 0; color: var(--soft); }
    .principles {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 14px;
    }
    .principle {
      min-height: 190px;
      background: var(--paper);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 18px;
    }
    .principle strong { display: block; font-size: 18px; margin-bottom: 8px; }
    .principle p { margin: 0; color: var(--soft); }
    .principle ul { margin: 14px 0 0; padding-left: 18px; color: var(--ink); }
    .workflow {
      display: grid;
      grid-template-columns: 320px minmax(0, 1fr);
      gap: 16px;
    }
    .steps {
      background: var(--paper);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 18px;
    }
    .step {
      display: grid;
      grid-template-columns: 34px 1fr;
      gap: 12px;
      padding: 12px 0;
      border-bottom: 1px solid var(--line);
    }
    .step:last-child { border-bottom: 0; }
    .step-number {
      width: 34px;
      height: 34px;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      border-radius: 50%;
      background: var(--rail);
      font-weight: 800;
      color: var(--blue);
    }
    .step strong { display: block; }
    .step span { color: var(--soft); font-size: 13px; }
    .matrix {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 14px;
    }
    .matrix-card {
      background: var(--paper);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 18px;
    }
    .matrix-card h3 { margin: 0 0 10px; font-size: 17px; }
    .matrix-card p { margin: 0; color: var(--soft); }
    .personal-grid {
      display: grid;
      grid-template-columns: minmax(0, 1fr) 360px;
      gap: 16px;
      align-items: stretch;
    }
    .decision-card {
      background: var(--paper);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
      padding: 20px;
    }
    .decision-card h3 { margin: 0 0 12px; font-size: 20px; }
    .decision-card p { margin: 0 0 14px; color: var(--soft); }
    .decision-card ul { margin: 0; padding-left: 18px; }
    .decision-card li { margin: 7px 0; }
    .verdict {
      display: grid;
      gap: 12px;
      background: #fff8ed;
      border: 1px solid #fed7aa;
      border-radius: 8px;
      padding: 18px;
    }
    .verdict strong { font-size: 18px; color: #7c3e06; }
    .verdict span { color: #7c3e06; }
    .rank-list {
      display: grid;
      gap: 10px;
    }
    .rank-item {
      display: grid;
      grid-template-columns: 38px 1fr auto;
      gap: 12px;
      align-items: center;
      padding: 14px;
      background: var(--paper);
      border: 1px solid var(--line);
      border-radius: 8px;
    }
    .rank-number {
      width: 38px;
      height: 38px;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      border-radius: 50%;
      background: var(--blue);
      color: #fff;
      font-weight: 900;
    }
    .rank-item strong { display: block; font-size: 16px; }
    .rank-item span { color: var(--soft); font-size: 13px; }
    .logic-table { margin-top: 14px; max-height: none; }
    .logic-table table { min-width: 980px; }
    .source-list {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 12px;
    }
    .notice + .source-list { margin-top: 14px; }
    .source-card {
      background: var(--paper);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 14px;
    }
    .source-card strong { display: block; margin-bottom: 6px; }
    .source-card span { color: var(--soft); font-size: 13px; }
    .summary-strip {
      display: grid;
      grid-template-columns: minmax(0, 1.2fr) repeat(3, minmax(0, 1fr));
      gap: 12px;
      margin-top: 18px;
    }
    .summary-card {
      background: var(--paper);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 16px;
      min-height: 118px;
    }
    .summary-card strong { display: block; font-size: 17px; margin-bottom: 8px; }
    .summary-card span { color: var(--soft); }
    .summary-card.urgent {
      background: #fff8ed;
      border-color: #fed7aa;
      color: #7c3e06;
    }
    .summary-card.urgent span { color: #7c3e06; }
    .profile-grid {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 12px;
    }
    .profile-card {
      background: var(--paper);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 16px;
    }
    .profile-card strong { display: block; margin-bottom: 8px; font-size: 16px; }
    .profile-card span { color: var(--soft); }
    .recommendation-table {
      overflow: auto;
      margin-top: 14px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--paper);
    }
    .recommendation-table table { min-width: 1180px; }
    .recommendation-table td:nth-child(1),
    .recommendation-table td:nth-child(2),
    .recommendation-table td:nth-child(3) { max-width: 260px; }
    .school-decision-layout {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 16px;
      margin-top: 16px;
      align-items: start;
    }
    .school-column {
      display: grid;
      gap: 12px;
    }
    .school-column > h3 {
      margin: 0;
      font-size: 18px;
      line-height: 1.3;
    }
    .school-card {
      display: grid;
      gap: 12px;
      background: var(--paper);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 16px;
      box-shadow: 0 10px 28px rgba(23, 32, 51, 0.05);
    }
    .school-card header {
      display: flex;
      align-items: flex-start;
      justify-content: space-between;
      gap: 12px;
      padding-bottom: 12px;
      border-bottom: 1px solid var(--line);
    }
    .school-card h3 {
      margin: 9px 0 0;
      font-size: 20px;
      line-height: 1.28;
    }
    .school-card header > strong {
      color: var(--blue);
      white-space: nowrap;
      font-size: 15px;
    }
    .school-card p {
      margin: 0;
      color: var(--soft);
    }
    .school-card .address-line {
      color: var(--ink);
      font-weight: 700;
    }
    .card-meta {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 10px;
    }
    .card-meta div {
      min-height: 62px;
      padding: 10px;
      border-radius: 8px;
      background: #f9fbfd;
      border: 1px solid var(--line);
    }
    .card-meta span {
      display: block;
      margin-bottom: 4px;
      color: var(--soft);
      font-size: 12px;
    }
    .card-meta b {
      display: block;
      line-height: 1.35;
      overflow-wrap: anywhere;
    }
    .risk-box {
      display: grid;
      gap: 4px;
      padding: 12px;
      border-radius: 8px;
      background: #fff8ed;
      border: 1px solid #fed7aa;
      color: #7c3e06;
    }
    .risk-box span {
      color: #7c3e06;
      font-size: 13px;
    }
    .rent-links {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
    }
    .rent-links a,
    .rent-board-card a {
      display: inline-flex;
      align-items: center;
      min-height: 34px;
      padding: 0 10px;
      border: 1px solid #bdd3ea;
      border-radius: 8px;
      background: #eef6ff;
      color: var(--blue);
      font-weight: 800;
      font-size: 12px;
    }
    .phone-link {
      color: var(--blue);
      font-weight: 800;
      white-space: nowrap;
    }
    .phone-separator {
      color: var(--soft);
      margin: 0 2px;
    }
    .distance-list {
      display: grid;
      gap: 5px;
      color: var(--soft);
      font-size: 12px;
    }
    .distance-list b { color: var(--ink); }
    .todo-board {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 12px;
    }
    .todo-column {
      background: var(--paper);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 16px;
    }
    .todo-column h3 { margin: 0 0 10px; font-size: 17px; }
    .todo-item {
      display: grid;
      grid-template-columns: 20px 1fr;
      gap: 10px;
      padding: 9px 0;
      border-top: 1px solid var(--line);
    }
    .todo-item input {
      width: 18px;
      min-height: 18px;
      margin-top: 2px;
      accent-color: var(--blue);
    }
    .todo-item span { color: var(--ink); }
    .todo-item small { display: block; color: var(--soft); margin-top: 2px; }
    .todo-item input:checked + span {
      color: var(--soft);
      text-decoration: line-through;
    }
    .judgement-grid {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 12px;
    }
    .judgement-card {
      background: var(--paper);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 16px;
    }
    .judgement-card h3 { margin: 0 0 8px; font-size: 16px; }
    .judgement-card p { margin: 0; color: var(--soft); }
    .priority-grid {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 12px;
    }
    .priority-card {
      display: grid;
      gap: 12px;
      background: var(--paper);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 16px;
      min-height: 210px;
    }
    .priority-card header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 10px;
    }
    .priority-card h3 { margin: 0; font-size: 17px; }
    .priority-card p { margin: 0; color: var(--soft); }
    .priority-card .reason { color: var(--ink); font-weight: 700; }
    .route-grid {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 12px;
    }
    .route-card,
    .candidate-card {
      display: grid;
      gap: 12px;
      background: var(--paper);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 16px;
    }
    .route-card header,
    .candidate-card header {
      display: flex;
      align-items: flex-start;
      justify-content: space-between;
      gap: 10px;
    }
    .route-card h3,
    .candidate-card h3 { margin: 0; font-size: 18px; line-height: 1.25; }
    .route-card p,
    .candidate-card p { margin: 0; color: var(--soft); }
    .route-card .reason { color: var(--ink); font-weight: 700; }
    .route-metrics {
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 8px;
      padding: 10px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #f9fbfd;
      color: var(--soft);
      font-size: 12px;
    }
    .route-metrics b { display: block; color: var(--ink); font-size: 20px; }
    .candidate-grid {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 12px;
      margin-top: 12px;
    }
    .rent-panel {
      background: var(--paper);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 16px;
      margin-bottom: 12px;
    }
    .rent-panel strong { display: block; margin-bottom: 6px; font-size: 17px; }
    .rent-panel p { margin: 0; color: var(--soft); }
    .rent-board-grid {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 12px;
    }
    .rent-board-card {
      display: grid;
      gap: 10px;
      background: var(--paper);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 16px;
      min-height: 190px;
    }
    .rent-board-card header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 10px;
    }
    .rent-board-card h3 { margin: 0; font-size: 17px; }
    .rent-board-card p { margin: 0; color: var(--soft); }
    .rent-builder {
      display: grid;
      gap: 14px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--paper);
      padding: 16px;
      margin-bottom: 12px;
    }
    .rent-builder-grid {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 10px;
    }
    .rent-url-preview {
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      padding-top: 12px;
      border-top: 1px solid var(--line);
    }
    .rent-url-preview code {
      max-width: 100%;
      overflow-wrap: anywhere;
      color: var(--soft);
      font-size: 12px;
    }
    .rental-result-grid {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 12px;
      margin: 12px 0;
    }
    .rental-result-card {
      overflow: hidden;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--paper);
    }
    .rental-photo {
      display: grid;
      align-content: end;
      min-height: 148px;
      padding: 14px;
      color: #fff;
      background:
        linear-gradient(180deg, rgba(17,24,39,0.05), rgba(17,24,39,0.78)),
        linear-gradient(135deg, #8fb6d8, #d9e5ce 52%, #c7b49b);
    }
    .rental-photo:hover { text-decoration: none; }
    .rental-photo span { font-weight: 900; }
    .rental-photo small { color: rgba(255,255,255,0.86); }
    .rental-result-body {
      display: grid;
      gap: 10px;
      padding: 14px;
    }
    .rental-result-body header {
      display: flex;
      justify-content: space-between;
      gap: 10px;
      align-items: flex-start;
    }
    .rental-result-body h3 { margin: 0; font-size: 17px; }
    .rental-result-body p { margin: 0; color: var(--soft); }
    .rental-meta { display: flex; flex-wrap: wrap; gap: 8px; }
    .rental-meta span {
      border-radius: 999px;
      background: var(--rail);
      color: var(--soft);
      padding: 4px 8px;
      font-size: 12px;
      font-weight: 800;
    }
    .action-steps {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 12px;
      counter-reset: actions;
    }
    .action-card {
      position: relative;
      background: var(--paper);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 18px;
      min-height: 170px;
    }
    .action-card::before {
      counter-increment: actions;
      content: counter(actions);
      display: inline-flex;
      align-items: center;
      justify-content: center;
      width: 30px;
      height: 30px;
      margin-bottom: 12px;
      border-radius: 50%;
      background: var(--blue);
      color: #fff;
      font-weight: 900;
    }
    .action-card strong { display: block; font-size: 17px; margin-bottom: 8px; }
    .action-card p { margin: 0; color: var(--soft); }
    .logic-compact {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 12px;
      margin-top: 16px;
    }
    .logic-card {
      background: #f9fbfd;
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 14px;
    }
    .logic-card strong { display: block; margin-bottom: 6px; }
    .logic-card span { color: var(--soft); font-size: 13px; }
    .details-section {
      padding: 4px 0 0;
    }
    .filter-panel {
      background: var(--paper);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
      padding: 16px;
    }
    .toolbar {
      display: grid;
      grid-template-columns: minmax(240px, 1.3fr) repeat(5, minmax(130px, 1fr));
      gap: 10px;
      align-items: end;
    }
    label { display: grid; gap: 6px; color: var(--soft); font-size: 12px; font-weight: 700; }
    input, select {
      width: 100%;
      min-height: 40px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #fff;
      color: var(--ink);
      padding: 8px 10px;
      font: inherit;
    }
    .result-bar {
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      margin-top: 14px;
      padding-top: 14px;
      border-top: 1px solid var(--line);
    }
    .result-count strong { font-size: 22px; }
    .chips { display: flex; flex-wrap: wrap; gap: 8px; }
    .chip {
      display: inline-flex;
      align-items: center;
      min-height: 28px;
      padding: 0 10px;
      border-radius: 999px;
      background: var(--rail);
      color: var(--soft);
      font-size: 12px;
      font-weight: 700;
    }
    .table-box {
      overflow: auto;
      margin-top: 16px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--paper);
      max-height: 76vh;
    }
    table { border-collapse: collapse; width: 100%; min-width: 1960px; }
    th, td { border-bottom: 1px solid var(--line); padding: 11px 12px; text-align: left; vertical-align: top; }
    th {
      position: sticky;
      top: 0;
      z-index: 3;
      background: #eef4f8;
      white-space: nowrap;
      font-size: 12px;
      color: #42536a;
    }
    td { font-size: 13px; }
    td:nth-child(14), td:nth-child(16) { max-width: 330px; }
    .school-name { font-weight: 800; font-size: 14px; }
    .sub { color: var(--soft); font-size: 12px; margin-top: 3px; }
    .tag {
      display: inline-flex;
      align-items: center;
      min-height: 24px;
      padding: 0 8px;
      border-radius: 999px;
      font-size: 12px;
      font-weight: 800;
      background: var(--rail);
      color: #42536a;
      white-space: nowrap;
    }
    .tag.green { background: #e8f5ef; color: var(--green); }
    .tag.blue { background: #eaf2fb; color: var(--blue); }
    .tag.amber { background: #fff5d7; color: var(--amber); }
    .tag.red { background: #fff0ee; color: var(--red); }
    .areas {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 12px;
    }
    .area-card {
      background: var(--paper);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 14px;
    }
    .area-card strong { display: block; margin-bottom: 10px; }
    .area-metrics { display: grid; grid-template-columns: repeat(3, 1fr); gap: 8px; color: var(--soft); font-size: 12px; }
    .area-metrics b { display: block; color: var(--ink); font-size: 18px; }
    .chart-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 12px;
      margin-bottom: 12px;
    }
    .chart-card {
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--paper);
      padding: 16px;
    }
    .chart-card h3 { margin: 0 0 12px; font-size: 17px; }
    .bar-row {
      display: grid;
      grid-template-columns: 96px minmax(0, 1fr) 42px;
      gap: 10px;
      align-items: center;
      padding: 8px 0;
      border-top: 1px solid var(--line);
    }
    .bar-row:first-of-type { border-top: 0; }
    .bar-row span { font-weight: 800; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .bar-row small { grid-column: 2 / 4; color: var(--soft); font-size: 12px; }
    .bar-track {
      height: 9px;
      border-radius: 999px;
      background: var(--rail);
      overflow: hidden;
    }
    .bar-track i {
      display: block;
      height: 100%;
      border-radius: inherit;
      background: var(--blue);
    }
    .notice {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 14px;
    }
    .notice-card {
      border: 1px solid #fed7aa;
      border-radius: 8px;
      background: #fff8ed;
      color: #7c3e06;
      padding: 16px;
    }
    .notice-card h3 { margin: 0 0 8px; font-size: 16px; }
    .notice-card p { margin: 0; }
    .module-stream {
      display: grid;
      gap: 12px;
    }
    .module-item {
      display: grid;
      grid-template-columns: 132px minmax(0, 1fr) auto;
      gap: 18px;
      align-items: start;
      background: var(--paper);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 16px;
    }
    .module-time {
      color: var(--soft);
      font-size: 13px;
      line-height: 1.5;
    }
    .module-body h3 {
      margin: 0 0 8px;
      font-size: 18px;
      line-height: 1.3;
    }
    .module-body p {
      margin: 0;
      color: var(--soft);
    }
    .module-body .module-reason {
      margin-top: 10px;
      padding-top: 10px;
      border-top: 1px dashed var(--line);
      color: var(--ink);
      font-size: 13px;
    }
    .module-status {
      display: grid;
      justify-items: end;
      gap: 8px;
      min-width: 128px;
    }
    .data-pills {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      margin-top: 12px;
    }
    .data-pill {
      display: inline-flex;
      align-items: center;
      min-height: 26px;
      padding: 0 9px;
      border-radius: 999px;
      background: var(--rail);
      color: var(--soft);
      font-size: 12px;
      font-weight: 800;
    }
    footer {
      border-top: 1px solid var(--line);
      color: var(--soft);
      padding: 24px 0 36px;
    }
    @media (max-width: 1100px) {
      .app-layout { grid-template-columns: 1fr; }
      .module-sidebar { position: static; grid-template-columns: repeat(2, minmax(0, 1fr)); max-height: none; }
      .hero-grid, .workflow, .notice, .personal-grid { grid-template-columns: 1fr; }
      .rent-builder-grid, .rental-result-grid, .chart-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      .module-item { grid-template-columns: 110px minmax(0, 1fr); }
      .module-status { grid-column: 2; justify-items: start; }
      .principles, .areas, .judgement-grid, .priority-grid, .action-steps, .logic-compact, .summary-strip, .profile-grid, .todo-board, .rent-board-grid, .route-grid, .candidate-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      .school-decision-layout { grid-template-columns: 1fr; }
      .toolbar { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      .source-list { grid-template-columns: 1fr; }
    }
    @media (max-width: 720px) {
      .shell { padding: 0 14px; }
      .topbar .shell { align-items: flex-start; flex-direction: column; padding-top: 12px; padding-bottom: 12px; }
      .nav { overflow: auto; width: 100%; padding-bottom: 2px; }
      .hero-grid { padding: 38px 14px 34px; min-height: auto; }
      .module-sidebar, .rent-builder-grid, .rental-result-grid, .chart-grid { grid-template-columns: 1fr; }
      .module-item { grid-template-columns: 1fr; }
      .module-status { grid-column: auto; justify-items: start; }
      .scoreboard { padding: 14px; }
      .principles, .areas, .matrix, .toolbar, .judgement-grid, .priority-grid, .action-steps, .logic-compact, .summary-strip, .profile-grid, .todo-board, .rent-board-grid, .route-grid, .candidate-grid, .card-meta { grid-template-columns: 1fr; }
      .school-card header { display: grid; }
      .school-card header > strong { white-space: normal; }
      .rank-item { grid-template-columns: 34px 1fr; }
      .rank-item .tag { grid-column: 2; justify-self: start; }
      .section-title { display: block; }
      .section-title p { margin-top: 8px; }
    }
  </style>
</head>
<body>
  <header class="topbar">
    <div class="shell">
      <div class="brand">上海家庭第一阶段落地执行方案</div>
      <nav class="nav" aria-label="页面导航">
        <a href="#mission">当前任务</a>
        <a href="#modules">数据模块</a>
        <a href="#architecture">架构</a>
        <a href="#routes">区域路线</a>
        <a href="#decision">幼儿园执行</a>
        <a href="#admission-history">招生概率</a>
        <a href="#rent">租房联动</a>
        <a href="#todos">入园待办</a>
        <a href="#query">详细查询</a>
        <a href="#sources">来源</a>
      </nav>
    </div>
  </header>

  <section class="hero">
    <div class="shell hero-grid">
      <div>
        <h1>90 天内完成上海家庭落地。</h1>
        <p class="lead">当前不是一次性决定 15 年教育路线，而是先完成幼儿园落位、租房签约、材料稳定、家具搬迁、通勤验证和家庭恢复运转。徐汇幼儿园是执行模块之一，现已把闵行、浦东公开名单纳入同一套落地路线比较。</p>
        <div class="hero-actions">
          <a class="button primary" href="#routes">看区域路线</a>
          <a class="button" href="#decision">看幼儿园执行</a>
          <a class="button" href="徐汇区幼儿园园区位置与择园参考.xlsx">打开 Excel 表</a>
        </div>
      </div>
      <aside class="scoreboard" aria-label="数据摘要">
        <div class="score-row"><span>覆盖区域</span><strong>3</strong></div>
        <div class="score-row"><span>全部园区/点位</span><strong>${campusItems.length}</strong></div>
        <div class="score-row"><span>徐汇点位</span><strong>${publicCampusItems.length + privateCampusItems.length}</strong></div>
        <div class="score-row"><span>闵行/浦东基础点位</span><strong>${externalCampusItems.length}</strong></div>
        <div class="score-row"><span>高德坐标/POI</span><strong>${amapMatchedCount}</strong></div>
      </aside>
    </div>
  </section>

  <main class="shell app-layout">
    <aside class="module-sidebar" aria-label="模块导航">
      <a class="side-module" href="#decision"><strong>幼儿园选择策略</strong><span>家庭情况、短期/长期目标、三区择园策略、入园待办。</span><small>查看方案推荐</small></a>
      <a class="side-module" href="#query"><strong>幼儿园数据查询</strong><span>查询基础数据源，按区、性质、等级、招生口径筛选。</span><small>查 930 条点位</small></a>
      <a class="side-module" href="#rent"><strong>租房查询</strong><span>选择贝壳筛选参数，生成实时查询入口和看房方向。</span><small>生成房源链接</small></a>
      <a class="side-module" href="#sources"><strong>数据来源参考</strong><span>查看数据集、采集来源、POI、贝壳参数和置信度说明。</span><small>追溯来源</small></a>
    </aside>
    <div class="module-content">
    <section id="mission">
      <div class="section-title">
        <h2>当前任务</h2>
        <p>第一阶段目标是让家庭进入稳定运行，而不是追求一步到位的终局方案。</p>
      </div>
      <div class="profile-grid">
        <article class="profile-card"><strong>90 天目标</strong><span>幼儿园落位、租房签约、材料稳定、家具搬迁、通勤验证、家庭恢复正常运转。</span></article>
        <article class="profile-card"><strong>家庭约束</strong><span>暂无居住证；办公点在西岸网易研发中心；至少 3 房、100㎡以上、预算 10000 以内。</span></article>
        <article class="profile-card"><strong>执行原则</strong><span>先解决“有”和“稳”，再谈长期最优；房东配合材料是租房硬门槛。</span></article>
        <article class="profile-card"><strong>非当前目标</strong><span>不在本阶段一次性决定买房、小学初中路径或长期定居区域。</span></article>
      </div>
    </section>

    <section id="modules">
      <div class="section-title">
        <h2>数据模块</h2>
        <p>参考信息流式结构，把每个决策模块拆成可追溯、可扩展、可继续追加的数据单元。</p>
      </div>
      <div class="module-stream">
        <article class="module-item">
          <div class="module-time">数据集<br>2026-05-09</div>
          <div class="module-body">
            <h3>幼儿园标准数据集</h3>
            <p>徐汇、闵行、浦东统一为同一套字段：区、片区、园所、分园、性质、等级、地址、电话、招生口径、高德坐标、租房板块、置信度和来源。</p>
            <div class="data-pills">
              <span class="data-pill">徐汇 ${standardizedDataset.counts.xuhui}</span>
              <span class="data-pill">闵行 ${standardizedDataset.counts.minhang}</span>
              <span class="data-pill">浦东 ${standardizedDataset.counts.pudong}</span>
              <span class="data-pill">总计 ${standardizedDataset.rowCount}</span>
            </div>
            <p class="module-reason">扩展规则：新增行政区时优先新增 <code>data/district_kindergarten_sources/{district}.json</code>，再补高德增强和贝壳板块映射；页面模块只消费标准字段。</p>
          </div>
          <div class="module-status"><span class="tag green">已标准化</span><a href="data/kindergartens/kindergarten_dataset.json" target="_blank" rel="noopener">查看 JSON</a></div>
        </article>
        <article class="module-item">
          <div class="module-time">执行层<br>90 天</div>
          <div class="module-body">
            <h3>落地路线模块</h3>
            <p>把区域选择、租房入口、材料风险、幼儿园争取线和民办兜底线拆开维护，避免页面变成单一长列表。</p>
            <p class="module-reason">下一步可继续追加“看房记录”“电话核验记录”“通勤实测记录”“报名材料状态”等模块。</p>
          </div>
          <div class="module-status"><span class="tag blue">可扩展</span><a href="#routes">进入路线</a></div>
        </article>
        <article class="module-item">
          <div class="module-time">核验层<br>持续更新</div>
          <div class="module-body">
            <h3>来源与置信度模块</h3>
            <p>每条数据保留来源、置信度和核验提醒；高德 POI 与地址编码分开标注，避免把“地址坐标”误当成“园所 POI 精确匹配”。</p>
            <p class="module-reason">这部分决定方案能不能复盘：后续新增数据时必须先说明来源和年度口径。</p>
          </div>
          <div class="module-status"><span class="tag amber">需持续复核</span><a href="#sources">看来源</a></div>
        </article>
      </div>
    </section>

    <section id="architecture">
      <div class="section-title">
        <h2>架构与数据流</h2>
        <p>当前核心链路是：采集公开数据，标准化为 JSON，补齐 POI 与租房参数，再生成推荐策略、网页、Excel 和可追溯来源。</p>
      </div>
      <div class="module-stream">
${renderArchitectureReviewCards()}
        <article class="module-item">
          <div class="module-time">Review<br>结论</div>
          <div class="module-body">
            <h3>当前最重要的架构原则</h3>
            <p>新增行政区、电话核验、看房记录、报名材料状态等能力，都应先进入独立数据文件，再由生成脚本消费；UI 只负责呈现和交互，不再承载原始数据维护。</p>
            <div class="data-pills">
              <span class="data-pill">标准数据 ${standardizedDataset.rowCount}</span>
              <span class="data-pill">高德覆盖 ${amapMatchedCount}/${campusItems.length}</span>
              <span class="data-pill">贝壳 token ${beikeDefaultTokens.join("")}</span>
            </div>
          </div>
          <div class="module-status"><span class="tag green">已落地</span><a href="docs/ARCHITECTURE_REVIEW.md" target="_blank" rel="noopener">查看 Review</a></div>
        </article>
      </div>
    </section>

    <section id="routes">
      <div class="section-title">
        <h2>区域路线</h2>
        <p>数据级扩展覆盖徐汇、闵行、浦东。首页先判断哪条路线最可执行，再进入幼儿园和租房细节。</p>
      </div>
      <div class="route-grid">
${renderDistrictRouteCards()}
      </div>
      <div class="section-title">
        <h2>落地评分模型</h2>
        <p>权重服务于第一阶段执行，不服务于终身学区排序。</p>
      </div>
      <div class="logic-compact">
${renderLandingScoreCards()}
      </div>
    </section>

    <section id="decision">
      <div class="section-title">
        <h2>幼儿园执行</h2>
        <p>策略主线调整为“每个区怎么选”：每区都拆成公办争取线、民办兜底线、租房板块和下一步核验动作。</p>
      </div>
      <div class="summary-strip">
        <article class="summary-card urgent">
          <strong>最大风险</strong>
          <span>暂无居住证会明显影响公办报名材料完整性和录取顺位，公办小班概率当前偏低。</span>
        </article>
        <article class="summary-card">
          <strong>首选路线</strong>
          <span>徐汇南部 + 闵行边界，兼顾通勤、材料和居住体验。</span>
        </article>
        <article class="summary-card">
          <strong>备选</strong>
          <span>浦东作为成长空间线，用于比较新社区和长期居住感。</span>
        </article>
        <article class="summary-card">
          <strong>第一行动</strong>
          <span>看房前先问房东是否配合居住登记、租赁合同/备案和居住证材料。</span>
        </article>
      </div>
      <div class="notice">
        <article class="notice-card">
          <h3>最终建议</h3>
          <p>先按“徐汇南部保守执行 + 闵行生活平衡”双线看房；若徐汇房源压抑或材料不配合，及时切到闵行春申/古美/梅陇。浦东只在通勤验证通过后进入主线。</p>
        </article>
        <article class="notice-card">
          <h3>高德数据口径</h3>
          <p>本页 ${amapMatchedCount}/${campusItems.length} 个点位已写入高德 POI 或地址坐标，覆盖率 ${amapMatchRate}；其中 ${amapAddressGeocodeCount} 个为地址地理编码而非 POI 精确匹配。距离均为到网易上海西岸研发中心的直线距离，实际通勤仍需复核。</p>
        </article>
      </div>
      <div class="section-title">
        <h2>三区择园策略</h2>
        <p>先判断区域是否可执行，再在区内选择公办争取线和民办兜底线。</p>
      </div>
      <div class="route-grid">
${renderDistrictKindergartenStrategyCards()}
      </div>
      <div class="section-title">
        <h2>徐汇园所样例</h2>
        <p>下面保留徐汇 5 所公办和 5 所民办样例，用作“区内如何落到园所”的示范；闵行、浦东完整点位在详细查询里按区筛选。</p>
      </div>
      <div class="school-decision-layout">
        <div class="school-column">
          <h3>徐汇公办争取样例</h3>
${renderSchoolDecisionCards("公办")}
        </div>
        <div class="school-column">
          <h3>徐汇民办兜底样例</h3>
${renderSchoolDecisionCards("民办")}
        </div>
      </div>
      <div class="section-title">
        <h2>闵行/浦东基础数据</h2>
        <p>完整点位在底部详细查询表中按区、性质和关键词筛选；这里每区只展示 6 个样例，便于先进入对应租房板块和高德核验。</p>
      </div>
      <h3>闵行样例</h3>
      <div class="candidate-grid">
${renderExternalCandidateCards("闵行")}
      </div>
      <h3>浦东样例</h3>
      <div class="candidate-grid">
${renderExternalCandidateCards("浦东")}
      </div>
    </section>

    <section id="judgement">
      <div class="section-title">
        <h2>公办/民办策略</h2>
        <p>当前最稳妥的结构是公办作为争取线，民办作为兜底线，两条线同步推进。</p>
      </div>
      <div class="judgement-grid">
        <article class="judgement-card"><h3>公办概率</h3><p>暂无居住证时，非沪籍公办录取顺位和材料完整性都不占优；公办只能作为争取线。</p></article>
        <article class="judgement-card"><h3>材料链</h3><p>租房地址、居住登记、居住证、租赁合同/备案要尽量一致；房东配合度是硬条件。</p></article>
        <article class="judgement-card"><h3>普惠民办</h3><p>应提升为并行主线，优先联系汇城苑、胡姬港湾、牛牛、凯琴、杜鹃园这 5 所民办兜底园。</p></article>
        <article class="judgement-card"><h3>租房预算</h3><p>1 万内、3 房、100 平以上、优先电梯和好环境，西岸核心区难度高，南部板块更现实。</p></article>
      </div>
      <div class="logic-compact">
        <article class="logic-card"><strong>资格层</strong><span>先看居住证、居住登记、居委和招生范围是否匹配。</span></article>
        <article class="logic-card"><strong>民办层</strong><span>先问是否接受非沪籍/居住登记凭证、剩余名额、收费和开放日。</span></article>
        <article class="logic-card"><strong>便利层</strong><span>看实际园区地址和接送路线，不只看幼儿园总名。</span></article>
        <article class="logic-card"><strong>偏好层</strong><span>最后比较托班、小班容量、老人接送和遛娃空间。</span></article>
      </div>
    </section>

    <section id="admission-history">
      <div class="section-title">
        <h2>招生概率与历年信息</h2>
        <p>公开渠道通常不发布逐园录取概率；本页只把可追溯信息转成概率分层，不伪造精确录取率。</p>
      </div>
      <div class="notice">
        <article class="notice-card">
          <h3>当前判断</h3>
          <p>暂无居住证时，公办园不能作为唯一方案；优先级应是“公办争取 + 民办兜底 + 材料先行”。</p>
        </article>
        <article class="notice-card">
          <h3>能用的信息</h3>
          <p>能用的是招生顺位、报名条件、对口居委、小班计划、民办名额电话核验和历年政策连续性。</p>
        </article>
        <article class="notice-card">
          <h3>不能硬算的信息</h3>
          <p>逐园报名人数、实际录取人数、统筹去向通常不公开，不能据此给出“百分比式概率”。</p>
        </article>
      </div>
      <div class="logic-compact">
${renderAdmissionHistoryCards()}
      </div>
    </section>

    <section id="rent">
      <div class="section-title">
        <h2>租房联动</h2>
        <p>租房入口基于贝壳结构化参数生成，不使用关键词搜索。默认条件固定为：${escapeHtml(beikeRentalConditionText)}。</p>
      </div>
      <div class="rent-builder">
        <div class="rent-builder-grid">
          <label>区域/商圈
            <select id="rentArea">
              ${rentalAreaOptions.map(([slug, label]) => `<option value="${escapeHtml(slug)}" ${slug === "changqiao" ? "selected" : ""}>${escapeHtml(label)}</option>`).join("")}
            </select>
          </label>
          <label>租赁方式
            <select id="rentType">
              <option value="rt200600000001" selected>整租</option>
              <option value="">不限</option>
              <option value="rt200600000002">合租</option>
            </select>
          </label>
          <label>租金
            <select id="rentPrice">
              <option value="brp0erp10000" selected>0-10000 元</option>
              <option value="rp7">7000-20000 元</option>
              <option value="">不限</option>
            </select>
          </label>
          <label>户型
            <select id="rentRooms">
              <option value="l2l3" selected>三居 / 四居+</option>
              <option value="l2">三居</option>
              <option value="l3">四居+</option>
              <option value="">不限</option>
            </select>
          </label>
          <label>面积
            <select id="rentSize">
              <option value="ra4ra5" selected>100-120㎡ / 120㎡以上</option>
              <option value="ra4">100-120㎡</option>
              <option value="ra5">120㎡以上</option>
              <option value="">不限</option>
            </select>
          </label>
          <label>朝向
            <select id="rentFace">
              <option value="">不限</option>
              <option value="f100500000003">朝南</option>
              <option value="f100500000009">南北</option>
            </select>
          </label>
          <label>租期
            <select id="rentTerm">
              <option value="">不限</option>
              <option value="rmp2">年租</option>
              <option value="rmp1">月租</option>
            </select>
          </label>
          <label>电梯
            <select id="rentElevator">
              <option value="ie1" selected>有电梯</option>
              <option value="">不限</option>
              <option value="ie0">无电梯</option>
            </select>
          </label>
        </div>
        <div class="rent-url-preview">
          <code id="rentUrlText">${escapeHtml(beikeRentalUrl("changqiao"))}</code>
          <a class="button primary" id="rentUrlLink" href="${escapeHtml(beikeRentalUrl("changqiao"))}" target="_blank" rel="noopener">打开贝壳查询</a>
        </div>
      </div>
      <div class="rent-panel">
        <strong>统一筛选 token</strong>
        <p><code>${beikeDefaultTokens.join("")}</code>，含整租、0-10000 元、三居/四居+、100-120㎡/120㎡以上、有电梯。二级商圈只替换 URL 中的区域 slug。</p>
        <p>房源图片和库存以贝壳实时页面为准；贝壳公开页面存在登录/风控限制，本静态页不缓存具体房源图片，避免展示过期或不可追溯图片。</p>
      </div>
      <div class="rental-result-grid">
${renderRentalSnapshotCards()}
      </div>
      <div class="rent-board-grid">
${renderRentBoardCards()}
      </div>
    </section>

    <section id="todos">
      <div class="section-title">
        <h2>入园待办清单</h2>
        <p>按不同入园时间点拆成可勾选任务，勾选状态会保存在当前浏览器本地，方便持续同步进度。</p>
      </div>
      <div class="todo-board">
        <article class="todo-column">
          <h3>立即处理</h3>
          <label class="todo-item"><input type="checkbox" data-todo="confirm-year"><span>确认目标入园年份和年龄段<small>2026 插班/小班、2027 小班、或更晚入园。</small></span></label>
          <label class="todo-item"><input type="checkbox" data-todo="rental-material"><span>看房前确认房东是否配合材料<small>居住登记、租赁合同/备案、地址一致。</small></span></label>
          <label class="todo-item"><input type="checkbox" data-todo="beike-shortlist"><span>建立贝壳房源候选表<small>记录小区、价格、面积、是否电梯、是否可办材料。</small></span></label>
          <label class="todo-item"><input type="checkbox" data-todo="private-backup"><span>同步联系普惠民办/民办园<small>优先电话问汇城苑、胡姬港湾、牛牛、凯琴、杜鹃园是否有名额。</small></span></label>
          <label class="todo-item"><input type="checkbox" data-todo="private-fee"><span>确认民办收费和退费规则<small>记录保育教育费、餐费、校车/延时服务等。</small></span></label>
        </article>
        <article class="todo-column">
          <h3>租房后 1-6 个月</h3>
          <label class="todo-item"><input type="checkbox" data-todo="residence-register"><span>办理居住登记/居住证相关手续<small>确认办理周期和所需材料。</small></span></label>
          <label class="todo-item"><input type="checkbox" data-todo="committee-check"><span>确认小区所属居委<small>用居委名称反查公办园招生范围。</small></span></label>
          <label class="todo-item"><input type="checkbox" data-todo="commute-test"><span>实测到公司和幼儿园路线<small>早高峰、雨天、老人接送都要测一次。</small></span></label>
          <label class="todo-item"><input type="checkbox" data-todo="school-call"><span>电话确认目标幼儿园材料口径<small>尤其是区域自主、扩招、多园区和 B 级地址。</small></span></label>
          <label class="todo-item"><input type="checkbox" data-todo="private-visit"><span>预约民办/普惠民办看园<small>实地看接送路线、园区环境、班额和师资。</small></span></label>
        </article>
        <article class="todo-column">
          <h3>招生季 3-5 月</h3>
          <label class="todo-item"><input type="checkbox" data-todo="policy-read"><span>核对目标区当年招生政策<small>报名条件、验证时间、录取批次每年可能调整。</small></span></label>
          <label class="todo-item"><input type="checkbox" data-todo="online-register"><span>完成信息登记和报名<small>按官方平台要求提交材料。</small></span></label>
          <label class="todo-item"><input type="checkbox" data-todo="private-deposit"><span>确认民办保位节点<small>问清保位费、退费期限和是否影响公办等待。</small></span></label>
          <label class="todo-item"><input type="checkbox" data-todo="verification"><span>准备现场/线上验证材料<small>居住证、租赁材料、户口本、出生证等。</small></span></label>
          <label class="todo-item"><input type="checkbox" data-todo="after-admission"><span>录取后办理体检和入园材料<small>同步安排老人接送路线和备用方案。</small></span></label>
        </article>
      </div>
    </section>

    <section id="query" class="details-section">
      <div class="section-title">
        <h2>详细查询</h2>
        <p>需要细查时再用这里：输入居委、小区、幼儿园或地址关键词，再叠加片区、招生类型、托班和置信度条件。</p>
      </div>
      <div class="filter-panel">
        <div class="toolbar">
          <label>关键词
            <input id="search" placeholder="例：康平、龙南、五原路、托班">
          </label>
          <label>片区
            <select id="area">
              <option value="">全部片区</option>
              ${[...new Set(campusItems.map((item) => item.area))].sort().map((area) => `<option>${escapeHtml(area)}</option>`).join("")}
            </select>
          </label>
          <label>区
            <select id="district">
              <option value="">全部区</option>
              ${[...new Set(campusItems.map((item) => item.district))].sort().map((district) => `<option>${escapeHtml(district)}</option>`).join("")}
            </select>
          </label>
          <label>招生类型
            <select id="admission">
              <option value="">全部类型</option>
              <option value="固定对口">固定对口</option>
              <option value="区域自主">区域自主</option>
              <option value="扩招">扩招</option>
              <option value="民办招生">民办招生</option>
              <option value="政策待核验">政策待核验</option>
            </select>
          </label>
          <label>托班模式
            <select id="toddler">
              <option value="">全部托班</option>
              <option value="明确托班">明确托班</option>
              <option value="混龄式招生">混龄式招生</option>
              <option value="待确认">待确认</option>
            </select>
          </label>
          <label>置信度
            <select id="confidence">
              <option value="">全部置信度</option>
              <option value="A">A</option>
              <option value="B">B</option>
            </select>
          </label>
          <label>电话确认
            <select id="confirm">
              <option value="">全部</option>
              <option value="yes">需要确认</option>
              <option value="no">暂不突出</option>
            </select>
          </label>
        </div>
        <div class="result-bar">
          <div class="result-count"><strong id="resultCount">${campusItems.length}</strong> 个园区匹配</div>
          <div class="chips">
            <span class="chip">先搜居委</span>
            <span class="chip">再看园区地址</span>
            <span class="chip">可筛民办兜底</span>
            <span class="chip">B 级务必电话确认</span>
          </div>
        </div>
      </div>
      <div class="table-box">
        <table>
          <thead>
            <tr>
              <th>幼儿园/园区</th>
              <th>区</th>
              <th>性质</th>
              <th>办园类型</th>
              <th>办园等级</th>
              <th>片区</th>
              <th>招生类型</th>
              <th>地址</th>
              <th>高德POI</th>
              <th>距公司</th>
              <th>联系电话</th>
              <th>托班</th>
              <th>小班</th>
              <th>置信度</th>
              <th>对口居委/招生范围</th>
              <th>地图</th>
              <th>备注</th>
            </tr>
          </thead>
          <tbody id="rows">
            ${campusItems.map((item) => `
              <tr
                data-area="${escapeHtml(item.area)}"
                data-district="${escapeHtml(item.district)}"
                data-confidence="${escapeHtml(item.confidence)}"
                data-admission="${escapeHtml(item.admissionType)}"
                data-toddler="${escapeHtml(item.toddlerMode)}"
                data-confirm="${item.needsConfirm ? "yes" : "no"}"
                data-text="${escapeHtml(item.searchText)}">
                <td><div class="school-name">${escapeHtml(item.name)}</div><div class="sub">${escapeHtml(item.campus)}</div></td>
                <td>${escapeHtml(item.district)}</td>
                <td><span class="tag ${item.nature === "民办" ? "amber" : "green"}">${escapeHtml(item.nature)}</span></td>
                <td>${escapeHtml(item.category)}</td>
                <td>${escapeHtml(item.level)}</td>
                <td>${escapeHtml(item.area)}</td>
                <td><span class="tag ${item.admissionType === "固定对口" ? "green" : item.admissionType === "扩招" || item.admissionType === "民办招生" ? "amber" : "blue"}">${escapeHtml(item.admissionType)}</span></td>
                <td>${escapeHtml(item.address)}</td>
                <td><div>${escapeHtml(item.amapPoiName)}</div><div class="sub">${escapeHtml(item.amapPoiAddress)}</div></td>
                <td>${escapeHtml(item.officeDistance)}</td>
                <td>${renderPhoneLinks(item.phone)}</td>
                <td><span class="tag ${item.toddlerMode === "混龄式招生" ? "amber" : "blue"}">${escapeHtml(item.toddler)}</span></td>
                <td>${Number.isFinite(Number(item.small)) ? escapeHtml(item.small) + " 班" : escapeHtml(item.small)}</td>
                <td><span class="tag ${item.confidence === "B" ? "red" : "green"}">${escapeHtml(item.confidence)}</span></td>
                <td>${escapeHtml(item.committee)}</td>
                <td><a href="${escapeHtml(item.mapUrl)}" target="_blank" rel="noopener">打开高德</a></td>
                <td>${escapeHtml(item.note) || (item.needsConfirm ? "建议结合当年简章或电话确认。" : '<span class="sub">-</span>')}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    </section>

    <section id="areas">
      <div class="section-title">
        <h2>片区容量可视化</h2>
        <p>用柱形图先看三区点位规模和徐汇片区容量；个人录取仍以居委、材料和当年政策为准。</p>
      </div>
      <div class="chart-grid">
        <article class="chart-card">
          <h3>三区点位规模</h3>
${renderBarChart(districtCapacityStats)}
        </article>
        <article class="chart-card">
          <h3>徐汇重点片区点位</h3>
${renderBarChart(areaCapacityStats)}
        </article>
      </div>
      <div class="areas">
        ${areaStats.slice(0, 12).map((row) => `
          <article class="area-card">
            <strong>${escapeHtml(row[0])}</strong>
            <div class="area-metrics">
              <span><b>${row[1]}</b>主体</span>
              <span><b>${row[2]}</b>点位</span>
              <span><b>${row[3]}</b>小班</span>
            </div>
          </article>
        `).join("")}
      </div>
    </section>

    <section id="sources">
      <div class="section-title">
        <h2>风险与来源</h2>
        <p>这些信息用于解释结论依据；政策和租房页面会变化，最终以当年官方公告、实际看房和电话确认为准。</p>
      </div>
      <div class="notice">
        <article class="notice-card">
          <h3>地址与园区分流</h3>
          <p>汇星幼儿园北园、复旦大学附属徐汇实验幼儿园存在公开资料差异或合并后的园部安排变化，建议直接电话确认 2026 年小班实际入读园区。</p>
        </article>
        <article class="notice-card">
          <h3>自主招生与扩招</h3>
          <p>“区域内自主招生”和“另向某街道扩招”不等于普通固定居委一一对应，应结合当年简章、报名条件和录取顺位理解。</p>
        </article>
        <article class="notice-card">
          <h3>高德匹配质量</h3>
          <p>高德已写入 ${amapMatchedCount} 个 POI/地址坐标，其中 ${amapAddressGeocodeCount} 个为地址地理编码；${amapUnmatchedItems.length} 个点位未写入坐标，主要是老地址、同名/近名园所或地图返回跨区结果。待核验示例：${escapeHtml(unmatchedPreview)}。</p>
        </article>
      </div>
      <div class="source-list">
        <article class="source-card">
          <strong><a href="https://edu.sh.gov.cn/xxgk2_zdgz_rxgkyzs_01/20260409/9a3caeed9a0e41c898149e46af9d9203.html" target="_blank" rel="noopener">上海市 2026 学前教育入园通知</a></strong>
          <span>来源：${escapeHtml(shanghaiAdmissionPolicySource)}</span>
          <span>用途：作为三区招生流程、规范程序、一网通办和不得测试等共性政策底层依据。</span>
        </article>
        <article class="source-card">
          <strong><a href="https://www.shanghai.gov.cn/xhqxqjy/20260421/c83b6efb6b4c43cb9071399f40215383.html" target="_blank" rel="noopener">2026 徐汇区幼儿园招生工作方案</a></strong>
          <span>来源：${escapeHtml(xuhuiOfficial2026Source)}</span>
          <span>用途：确认相对就近、户籍优先、人户一致优先、外省市户籍排序统筹和录取批次。</span>
        </article>
        <article class="source-card">
          <strong><a href="https://www.shanghai.gov.cn/xhqxqjy/20250417/85cd0238c43d49d5ae72cfbe65544a0b.html" target="_blank" rel="noopener">2025 徐汇区幼儿园招生工作方案</a></strong>
          <span>来源：${escapeHtml(xuhuiOfficial2025Source)}</span>
          <span>用途：用于观察徐汇近两年政策连续性；只能支持顺位判断，不能推出逐园精确录取概率。</span>
        </article>
        <article class="source-card">
          <strong>公办招生范围与计划班级数</strong>
          <span>来源：${escapeHtml(pdfSource)}</span>
          <span>用途：生成 42 个公办招生主体、对口居委、小班计划、托班/混龄信息，是公办筛选的主数据源。</span>
        </article>
        <article class="source-card">
          <strong><a href="https://m.sh.bendibao.com/edu/305291.html" target="_blank" rel="noopener">2026 徐汇区幼儿园招生对口地段及计划</a></strong>
          <span>来源：${escapeHtml(planSource)}</span>
          <span>用途：与本地 PDF 互相校验公办招生计划、对口地段和扩招/自主招生口径。</span>
        </article>
        <article class="source-card">
          <strong><a href="https://www.shobserver.cn/sgh/detail?id=1735182" target="_blank" rel="noopener">2026 闵行入园问答</a></strong>
          <span>来源：${escapeHtml(minhangAdmissionPolicySource)}</span>
          <span>用途：确认闵行报名、验证、梯队录取、溢出统筹和区级咨询渠道。</span>
        </article>
        <article class="source-card">
          <strong><a href="https://www.pudong.gov.cn/zwgk/xqjy-jyjzdgz/2026/105/354450.html" target="_blank" rel="noopener">2026 浦东学前教育入园实施方案</a></strong>
          <span>来源：${escapeHtml(pudongAdmissionPolicySource)}</span>
          <span>用途：确认浦东按地段/全区招生、人户一致优先、来沪人员随迁子女条件和统筹口径。</span>
        </article>
        <article class="source-card">
          <strong><a href="https://m.sh.bendibao.com/edu/305384.html" target="_blank" rel="noopener">2026 徐汇幼儿园报名材料</a></strong>
          <span>用途：核对非本市户籍幼儿、父母居住证、租赁合同、居住登记和地址一致性要求。</span>
          <span>本页结论中的“暂无居住证风险”来自该材料口径与招生顺位判断。</span>
        </article>
        <article class="source-card">
          <strong><a href="https://m.sh.bendibao.com/zffw/285087.html" target="_blank" rel="noopener">上海居住证办理条件</a></strong>
          <span>用途：确认居住登记、合法稳定住所、租赁合同/备案等材料链要求。</span>
          <span>租房看房清单中的“房东是否配合材料”来自该口径。</span>
        </article>
        <article class="source-card">
          <strong><a href="https://m.sh.bendibao.com/edu/284155.html" target="_blank" rel="noopener">徐汇区公办幼儿园名单</a></strong>
          <span>来源：${escapeHtml(publicListSource)}</span>
          <span>用途：补充公办园区/分园地址、办园等级和电话，与 2026 对口表拆分为 81 个公办园区点位。</span>
        </article>
        <article class="source-card">
          <strong><a href="https://sh.bendibao.com/edu/2024325/284155_2.shtm" target="_blank" rel="noopener">徐汇区民办幼儿园名单</a></strong>
          <span>来源：${escapeHtml(privateListSource)}</span>
          <span>用途：补充 46 个民办/私立点位的级别、地址和联系电话，作为公办概率偏低时的兜底池。</span>
        </article>
        <article class="source-card">
          <strong>闵行幼儿园基础数据</strong>
          <span>来源：${escapeHtml(minhangPublicListSource)}；${escapeHtml(minhangPrivateListSource)}</span>
          <span>用途：补充闵行 ${externalCampusCounts.minhangPublic || 0} 个公办点位、${(externalCampusCounts.minhangPrivate || 0) + (externalCampusCounts.minhangOther || 0)} 个民办/中外合作点位，用于徐汇南部以外的生活平衡路线比较；闵行民办名单存在年度口径差异，报名季必须电话确认。</span>
        </article>
        <article class="source-card">
          <strong>浦东幼儿园基础数据</strong>
          <span>来源：${escapeHtml(pudongPublicListSource)}；${escapeHtml(pudongPrivateListSource)}</span>
          <span>用途：补充浦东 ${externalCampusCounts.pudongPublic || 0} 个公办点位、${(externalCampusCounts.pudongPrivate || 0) + (externalCampusCounts.pudongOther || 0)} 个民办/中外合作点位，用于成长空间路线比较；板块跨度大，必须叠加通勤核验。</span>
        </article>
        <article class="source-card">
          <strong>区级源数据</strong>
          <span>来源：<code>data/district_kindergarten_sources/xuhui.json</code>、<code>minhang.json</code>、<code>pudong.json</code>。</span>
          <span>用途：统一保存每个区的原始采集结果。徐汇不再作为脚本内特殊数据源维护，闵行/浦东也按同一结构进入加工链路。</span>
        </article>
        <article class="source-card">
          <strong>标准化幼儿园数据集</strong>
          <span>来源：<code>data/kindergartens/kindergarten_dataset.json</code> 与 <code>data/kindergartens/xuhui_kindergartens.json</code>，字段定义见 <code>data/kindergarten_dataset.schema.json</code>。</span>
          <span>用途：将徐汇、闵行、浦东统一成可扩展数据层，后续新增行政区或模块时不再改页面主结构。</span>
        </article>
        <article class="source-card">
          <strong>高德地图 POI 与距离数据</strong>
          <span>来源：高德地图 MCP/API 查询结果，落地在 <code>data/amap_enrichment.json</code>。</span>
          <span>用途：补充 POI/地址坐标、经纬度，以及到${escapeHtml(officeLocation.name)}的直线距离；当前覆盖 ${amapMatchedCount}/${campusItems.length}，覆盖率 ${amapMatchRate}，地址地理编码 ${amapAddressGeocodeCount} 个。</span>
        </article>
        <article class="source-card">
          <strong><a href="https://sh.zu.ke.com/" target="_blank" rel="noopener">贝壳上海租房结构化筛选</a></strong>
          <span>来源：贝壳租房页面参数解析，结构化数据保存在 <code>data/beike_rental_filter_schema.json</code>。</span>
          <span>用途：生成整租、0-10000 元、三居/四居+、100㎡以上、有电梯的房源链接，固定 token 为 <code>${beikeDefaultTokens.join("")}</code>。</span>
        </article>
        <article class="source-card">
          <strong><a href="https://sh.zu.ke.com/wzdt/" target="_blank" rel="noopener">贝壳徐汇租房板块</a></strong>
          <span>用途：解析徐汇一级区域和长桥、华泾、龙华、康健、植物园、徐汇滨江、田林等二级商圈 slug。</span>
          <span>说明：贝壳房源价格和库存实时变化，本页只保留查询入口，不固化具体房源价格。</span>
        </article>
        <article class="source-card">
          <strong><a href="https://www.chooffice.com/1357.html" target="_blank" rel="noopener">西岸网易研发中心办公点参考</a></strong>
          <span>用途：确认办公点在徐汇滨江/龙耀路/云锦路一带，作为通勤和园所距离估算的目标点。</span>
          <span>说明：实际通勤仍需按步行、骑行、驾车路线二次验证。</span>
        </article>
        <article class="source-card">
          <strong>第三方园所公开资料</strong>
          <span>来源：上哪学、021school、园所公开页面等，用于交叉核对部分民办园电话、收费、托幼一体和地址差异。</span>
          <span>示例：<a href="https://www.xxh-edu.com/brand/57.html" target="_blank" rel="noopener">汇城苑幼稚园</a>、<a href="https://www.021school.cn/schools/12780" target="_blank" rel="noopener">凯琴数码幼儿园</a>、<a href="https://www.021school.cn/schools/12790" target="_blank" rel="noopener">小神童幼儿园</a>。第三方信息仅作辅助，最终以园所电话确认为准。</span>
        </article>
      </div>
    </section>
    </div>
  </main>

  <footer>
    <div class="shell">来源、数据口径和置信度说明已合并到“风险与来源”模块；页面结论仅用于择园和租房初筛，报名前以当年官方公告、园所电话和实地看房为准。</div>
  </footer>

  <script>
    const search = document.querySelector("#search");
    const area = document.querySelector("#area");
    const district = document.querySelector("#district");
    const admission = document.querySelector("#admission");
    const toddler = document.querySelector("#toddler");
    const confidence = document.querySelector("#confidence");
    const confirmSelect = document.querySelector("#confirm");
    const resultCount = document.querySelector("#resultCount");
    const rows = Array.from(document.querySelectorAll("#rows tr"));
    const todoInputs = Array.from(document.querySelectorAll("[data-todo]"));
    const rentControls = ["rentArea", "rentType", "rentPrice", "rentRooms", "rentSize", "rentFace", "rentTerm", "rentElevator"]
      .map((id) => document.querySelector("#" + id))
      .filter(Boolean);
    const rentUrlText = document.querySelector("#rentUrlText");
    const rentUrlLink = document.querySelector("#rentUrlLink");

    function applyFilters() {
      const q = search.value.trim().toLowerCase();
      let visible = 0;
      for (const row of rows) {
        const okText = !q || row.dataset.text.toLowerCase().includes(q);
        const okArea = !area.value || row.dataset.area === area.value;
        const okDistrict = !district.value || row.dataset.district === district.value;
        const okAdmission = !admission.value || row.dataset.admission === admission.value;
        const okToddler = !toddler.value || row.dataset.toddler === toddler.value;
        const okConfidence = !confidence.value || row.dataset.confidence === confidence.value;
        const okConfirm = !confirmSelect.value || row.dataset.confirm === confirmSelect.value;
        const show = okText && okArea && okDistrict && okAdmission && okToddler && okConfidence && okConfirm;
        row.hidden = !show;
        if (show) visible += 1;
      }
      resultCount.textContent = visible;
    }

    for (const control of [search, area, district, admission, toddler, confidence, confirmSelect]) {
      control.addEventListener("input", applyFilters);
      control.addEventListener("change", applyFilters);
    }

    for (const input of todoInputs) {
      const key = "xuhui-kindergarten-todo-" + input.dataset.todo;
      input.checked = localStorage.getItem(key) === "1";
      input.addEventListener("change", () => {
        localStorage.setItem(key, input.checked ? "1" : "0");
      });
    }

    function updateRentUrl() {
      if (!rentUrlText || !rentUrlLink) return;
      const values = Object.fromEntries(rentControls.map((control) => [control.id, control.value]));
      const tokens = [
        values.rentType,
        values.rentPrice,
        values.rentRooms,
        values.rentSize,
        values.rentFace,
        values.rentTerm,
        values.rentElevator,
      ].filter(Boolean).join("");
      const url = "https://sh.zu.ke.com/zufang/" + (values.rentArea || "xuhui") + "/" + tokens + "/?showMore=1";
      rentUrlText.textContent = url;
      rentUrlLink.href = url;
    }

    for (const control of rentControls) {
      control.addEventListener("change", updateRentUrl);
    }
    updateRentUrl();
  </script>
</body>
</html>`;

const csvEscape = (value) => `"${String(value ?? "").replaceAll('"', '""')}"`;
const toCsv = (rows) => rows.map((row) => row.map(csvEscape).join(",")).join("\n");

await fs.mkdir(outputDir, { recursive: true });
await fs.mkdir(path.join(outputDir, "data", "kindergartens"), { recursive: true });
await fs.mkdir(path.join(outputDir, "data", "district_kindergarten_sources"), { recursive: true });
await fs.mkdir(path.join(outputDir, "docs"), { recursive: true });
await fs.writeFile(path.join(outputDir, "徐汇区幼儿园园区位置表.csv"), toCsv([campusHeader, ...campusData]), "utf8");
await fs.writeFile(path.join(outputDir, "徐汇区幼儿园择园参考.md"), markdown, "utf8");
await fs.writeFile(path.join(outputDir, "徐汇区幼儿园园区位置与择园参考.html"), html, "utf8");
await fs.writeFile(path.join(outputDir, "index.html"), html, "utf8");
await fs.writeFile(path.join(outputDir, "data", "project_architecture_review.json"), `${JSON.stringify({
  ...architectureReview,
  metrics: {
    kindergartenRows: standardizedDataset.rowCount,
    districtCounts: standardizedDataset.counts,
    amapCoverage: `${amapMatchedCount}/${campusItems.length}`,
    amapAddressGeocodeCount,
    beikeDefaultTokens: beikeDefaultTokens.join(""),
  },
}, null, 2)}\n`, "utf8");
for (const source of districtKindergartenSources) {
  const slug = source.district === "徐汇" ? "xuhui" : source.district === "闵行" ? "minhang" : "pudong";
  await fs.writeFile(path.join(outputDir, "data", "district_kindergarten_sources", `${slug}.json`), `${JSON.stringify(source, null, 2)}\n`, "utf8");
}
await fs.copyFile(path.join(process.cwd(), "data", "kindergarten_dataset.schema.json"), path.join(outputDir, "data", "kindergarten_dataset.schema.json"));
await fs.writeFile(path.join(outputDir, "docs", "ARCHITECTURE_REVIEW.md"), architectureReviewMarkdown, "utf8");
await fs.writeFile(path.join(outputDir, "data", "kindergartens", "kindergarten_dataset.json"), `${JSON.stringify(standardizedDataset, null, 2)}\n`, "utf8");
await fs.writeFile(path.join(outputDir, "data", "kindergartens", "xuhui_kindergartens.json"), `${JSON.stringify({
  schemaVersion: "1.0",
  updatedAt: standardizedDataset.updatedAt,
  district: "徐汇",
  rowCount: standardizedDataset.counts.xuhui,
  sources: [pdfSource, planSource, publicListSource, privateListSource],
  rows: standardizedKindergartenRows.filter((item) => item.district === "徐汇"),
}, null, 2)}\n`, "utf8");

const workbook = Workbook.create();
const summary = workbook.worksheets.add("摘要");
const schoolSheet = workbook.worksheets.add("招生主体");
const campusSheet = workbook.worksheets.add("园区点位");
const areaSheet = workbook.worksheets.add("片区统计");
const sourceSheet = workbook.worksheets.add("来源与核验规则");

summary.getRange(`A1:C${summaryRows.length}`).values = summaryRows;
schoolSheet.getRange(`A1:I${schoolData.length + 1}`).values = [schoolHeader, ...schoolData];
campusSheet.getRange(`A1:V${campusData.length + 1}`).values = [campusHeader, ...campusData];
areaSheet.getRange(`A1:F${areaStats.length + 1}`).values = [["片区", "招生主体数", "园区点位数", "小班计划数", "明确托班班级数", "混龄招生主体数"], ...areaStats];
const sourceRows = [
  ["项目", "说明"],
  ["PDF主表", pdfSource],
  ["2026线上主表", planSource],
  ["上海市2026学前教育入园通知", shanghaiAdmissionPolicySource],
  ["徐汇2026招生工作方案", `${xuhuiOfficial2026Source}；用于判断户籍优先、外省市户籍排序、录取批次、报名验证和统筹规则。`],
  ["徐汇2025招生工作方案", `${xuhuiOfficial2025Source}；用于观察近两年政策连续性，不用于推导逐园精确录取概率。`],
  ["闵行2026入园问答", `${minhangAdmissionPolicySource}；用于确认报名、验证、梯队录取、溢出统筹和咨询渠道。`],
  ["浦东2026入园实施方案", `${pudongAdmissionPolicySource}；用于确认按地段/全区招生、人户一致优先、来沪人员条件和统筹口径。`],
  ["2026报名材料", "https://m.sh.bendibao.com/edu/305384.html；用于核对非沪籍幼儿、父母居住证、租赁合同、居住登记和地址一致性要求。"],
  ["上海居住证办理条件", "https://m.sh.bendibao.com/zffw/285087.html；用于确认居住登记、合法稳定住所、租赁合同/备案等材料链要求。"],
  ["公办园部地址主来源", publicListSource],
  ["民办/私立地址与电话来源", privateListSource],
  ["闵行公办地址与电话来源", minhangPublicListSource],
  ["闵行民办地址与电话来源", minhangPrivateListSource],
  ["浦东公办地址与电话来源", pudongPublicListSource],
  ["浦东民办地址与电话来源", pudongPrivateListSource],
  ["区级源数据", `data/district_kindergarten_sources/；徐汇、闵行、浦东均按同一结构保存采集结果。当前共${campusItems.length}条，徐汇${publicCampusItems.length + privateCampusItems.length}条，闵行${(externalCampusCounts.minhangPublic || 0) + (externalCampusCounts.minhangPrivate || 0) + (externalCampusCounts.minhangOther || 0)}条，浦东${(externalCampusCounts.pudongPublic || 0) + (externalCampusCounts.pudongPrivate || 0) + (externalCampusCounts.pudongOther || 0)}条。`],
  ["闵行/浦东历史合并数据", "data/cross_district_kindergartens.json；保留为迁移前的跨区合并源，当前生成链路优先使用 data/district_kindergarten_sources/minhang.json 和 pudong.json。"],
  ["标准化幼儿园数据集", `data/kindergartens/kindergarten_dataset.json；共${standardizedDataset.rowCount}条，徐汇${standardizedDataset.counts.xuhui}条、闵行${standardizedDataset.counts.minhang}条、浦东${standardizedDataset.counts.pudong}条；字段定义见data/kindergarten_dataset.schema.json。`],
  ["徐汇标准化数据", "data/kindergartens/xuhui_kindergartens.json；将徐汇公办计划、园区地址、等级、电话、民办兜底和来源统一为与闵行/浦东一致的字段结构。"],
  ["架构Review", "docs/ARCHITECTURE_REVIEW.md 与 data/project_architecture_review.json；说明数据采集、标准数据层、POI、租房、推荐策略和发布层的核心逻辑与优化边界。"],
  ["区域落地画像", "data/district_landing_profiles.json；用于维护徐汇、闵行、浦东三条第一阶段落地路线。"],
  ["高德MCP/API接入口径", `目标办公点：${officeLocation.name}；高德匹配为“网易上海西岸研发中心”。用关键词搜索、地址地理编码和距离测量补齐高德POI/地址坐标、经纬度和距离。`],
  ["高德增强落地数据", `data/amap_enrichment.json；用于补充POI名称/地址坐标、POI地址、经纬度、距公司直线距离。当前覆盖${amapMatchedCount}/${campusItems.length}，其中地址地理编码${amapAddressGeocodeCount}个。`],
  ["高德检索口径", "表格中的高德链接使用“上海市徐汇区 + 幼儿园名 + 园区名 + 地址”生成，用于逐点打开核验。"],
  ["贝壳租房参数", `data/beike_rental_filter_schema.json；用于生成徐汇及二级商圈结构化租房链接，固定token：${beikeDefaultTokens.join("")}。`],
  ["贝壳租房口径", "https://sh.zu.ke.com/；房源价格和库存实时变化，本页只保存查询条件和入口，不固化具体房源。"],
  ["办公点参考", "https://www.chooffice.com/1357.html；用于确认西岸网易研发中心靠近龙耀路、徐汇滨江一带。"],
  ["第三方园所资料", "上哪学、021school、园所公开页面等；用于交叉核对部分民办园电话、收费、托幼一体和地址差异，最终以园所电话确认为准。"],
  ["A级", "园部地址在公办园地址清单中明确列出，且无明显冲突。"],
  ["B级", "公开资料有变更、合并或地址冲突，位置大体可定位，但实际入读园区需电话确认。"],
  ["C级", "仅有第三方或地图信息，缺少官方交叉验证。本表目前未使用C级作为主确认。"],
  ["电话确认重点", "汇星幼儿园北园；复旦大学附属徐汇实验幼儿园；区域内自主招生、扩招园、民办/私立园。"],
];
sourceSheet.getRange(`A1:B${sourceRows.length}`).values = sourceRows;

for (const sheet of [summary, schoolSheet, campusSheet, areaSheet, sourceSheet]) {
  sheet.getRange("A1:Z1").format = { fontWeight: "bold", fill: "#E8F2FF", wrapText: true };
  sheet.freezePanes.freezeRows(1);
}

summary.getRange("A:A").format.columnWidthPx = 150;
summary.getRange("B:B").format.columnWidthPx = 150;
summary.getRange("C:C").format.columnWidthPx = 560;
summary.getRange(`A1:C${summaryRows.length}`).format.wrapText = true;

schoolSheet.getRange("A:A").format.columnWidthPx = 54;
schoolSheet.getRange("B:B").format.columnWidthPx = 190;
schoolSheet.getRange("C:C").format.columnWidthPx = 120;
schoolSheet.getRange("D:F").format.columnWidthPx = 88;
schoolSheet.getRange("G:G").format.columnWidthPx = 520;
schoolSheet.getRange("H:I").format.columnWidthPx = 280;
schoolSheet.getRange(`A1:I${schoolData.length + 1}`).format.wrapText = true;

campusSheet.getRange("A:A").format.columnWidthPx = 54;
campusSheet.getRange("B:B").format.columnWidthPx = 70;
campusSheet.getRange("C:E").format.columnWidthPx = 90;
campusSheet.getRange("F:F").format.columnWidthPx = 190;
campusSheet.getRange("G:G").format.columnWidthPx = 120;
campusSheet.getRange("H:I").format.columnWidthPx = 150;
campusSheet.getRange("J:K").format.columnWidthPx = 160;
campusSheet.getRange("L:M").format.columnWidthPx = 120;
campusSheet.getRange("N:N").format.columnWidthPx = 360;
campusSheet.getRange("O:O").format.columnWidthPx = 120;
campusSheet.getRange("P:Q").format.columnWidthPx = 88;
campusSheet.getRange("R:R").format.columnWidthPx = 520;
campusSheet.getRange("S:T").format.columnWidthPx = 90;
campusSheet.getRange("U:V").format.columnWidthPx = 360;
campusSheet.getRange(`A1:V${campusData.length + 1}`).format.wrapText = true;
campusSheet.freezePanes.freezeColumns(4);

areaSheet.getRange("A:A").format.columnWidthPx = 160;
areaSheet.getRange("B:F").format.columnWidthPx = 110;

sourceSheet.getRange("A:A").format.columnWidthPx = 150;
sourceSheet.getRange("B:B").format.columnWidthPx = 760;
sourceSheet.getRange(`A1:B${sourceRows.length}`).format.wrapText = true;

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(path.join(outputDir, "徐汇区幼儿园园区位置与择园参考.xlsx"));

console.log(JSON.stringify({
  outputDir,
  campuses: campusData.length,
  schools: schools.length,
  publicCampuses: publicCampusItems.length,
  privateCampuses: privateCampusItems.length,
  xlsx: path.join(outputDir, "徐汇区幼儿园园区位置与择园参考.xlsx"),
  csv: path.join(outputDir, "徐汇区幼儿园园区位置表.csv"),
  md: path.join(outputDir, "徐汇区幼儿园择园参考.md"),
  html: path.join(outputDir, "徐汇区幼儿园园区位置与择园参考.html"),
  index: path.join(outputDir, "index.html"),
}, null, 2));
