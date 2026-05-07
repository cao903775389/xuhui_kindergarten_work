import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = path.join(process.cwd(), "outputs");

const pdfSource = "本地PDF：/Users/caofengyang/Downloads/徐汇区幼儿园对口地区.pdf（2026年徐汇区公办幼儿园对口居委一览表及计划班级数）";
const publicListSource = "上海本地宝《上海市徐汇区公办幼儿园名单一览》，内容来源标注为上海学前教育网，2024-03-25：https://m.sh.bendibao.com/edu/284155.html";
const privateListSource = "上海本地宝《上海市徐汇区民办幼儿园名单一览》，内容来源标注为上海学前教育网，2024-03-25：https://sh.bendibao.com/edu/2024325/284155_2.shtm";
const planSource = "上海本地宝《2026徐汇区幼儿园招生对口地段及招生计划班级数》，2026-04-15：https://m.sh.bendibao.com/edu/305291.html";

const schools = [
  { id: 1, name: "襄阳南路第一幼儿园", area: "衡复/湖南", toddler: "3", small: 4, committee: "慎成里、肇嘉浜、陕西、建新、息村、桃源村、嘉善、永康、张家弄、复中" },
  { id: 2, name: "五原路幼儿园", area: "衡复/湖南", toddler: "3", small: 4, committee: "上海新村、东湖、金波、延庆、陕新、新乐、复襄、淮海、安福、春华、淮中、建岳、太原、永太、永嘉新村、复永、兴武、华康、武康" },
  { id: 3, name: "汇星幼儿园", area: "徐家汇/天平", toddler: "1", small: 3, committee: "天平、高安、康平、吴兴、宛平、余庆、广元、安亭、徐汇新村、张家浜肇谨、零陵、科汇、启明、沈马、王家堂、南赵巷、殷家角、汇翠、名园、德昌" },
  { id: 4, name: "乐山幼儿园", area: "徐家汇/交大", toddler: "1", small: 3, committee: "乐山一村、乐山二三村、乐山四五村、乐山六七村、乐山八九村、文定、汇站、交大、豪庭" },
  { id: 5, name: "汇家幼儿园", area: "徐家汇/虹桥路", toddler: "1", small: 3, committee: "番禺、南丹、柿子湾、潘家宅、陈家宅、虹二、虹交、西塘、东塘" },
  { id: 6, name: "枫林幼儿园", area: "枫林/斜土", toddler: "1", small: 2, committee: "日新、日二、日五、日六第一、日六第二、大三、大四、康巨、茶陵、景泰、恒益、西木南、西木北、江南" },
  { id: 7, name: "复旦大学附属徐汇实验幼儿园", area: "枫林/斜土", toddler: "1", small: 4, committee: "平江、肇一、上影、日七、大五、医清、四季园、肇清、医学院" },
  { id: 8, name: "东安一村幼儿园", area: "枫林/斜土", toddler: "1", small: 2, committee: "东安一村北、东安一村南、安康、振兴、张东、爱华、汇园、东安二村、沈家里、东安四村、谨斜、东安苑、枫林新村、南康" },
  { id: 9, name: "龙山幼儿园", area: "枫林/宛南", toddler: "1", small: 4, committee: "天一二、天三、天四、宛南一二、宛三、宛四、宛五、宛六、庄家宅、黄家宅、华容、龙山一、龙山二、徐汇苑" },
  { id: 10, name: "龙华幼儿园", area: "龙华/西岸", toddler: "2", small: 5, committee: "上缝、强生、东蔡、狮城、佳友、华富、宏润花园、漕溪四村单号、丰谷、云锦、机场、尚海湾、盛大、民苑、丰谷三、俞一、俞二、俞三、龙华新村、周家湾、汇龙、红馨、百汇园" },
  { id: 11, name: "汇霖幼儿园", area: "田林/虹桥", toddler: "1", small: 4, committee: "古宜、长春、吴中、桂林苑、虹南、虹星、怡桂苑、钦北、华悦家园、鑫耀" },
  { id: 12, name: "阳光幼儿园", area: "龙华/滨江", toddler: "1", small: 3, committee: "龙南三四、龙南五、龙南六、龙南七、滨江、樟树苑、东泉、张家园" },
  { id: 13, name: "漕溪新村幼儿园", area: "漕溪/龙漕", toddler: "1", small: 3, committee: "金谷园、南一村一、南一村二、漕溪四村双号、凯翔、龙漕、嘉萱苑、漕东" },
  { id: 14, name: "望德幼儿园", area: "康健/南站", toddler: "4", small: 8, committee: "康健路、习勤、薛家宅、科苑、九弄、中海、梓树园、月河、宾阳、馨汇南苑、公园道" },
  { id: 15, name: "瑞德幼儿园", area: "长桥/汇成", toddler: "2", small: 3, committee: "汇成一村、汇成二村、汇成三村、汇成四村、华东二、楼园、金牛、挹翠苑、罗城、光华、汇澜苑、金塘" },
  { id: 16, name: "益思幼儿园", area: "田林/漕河泾", toddler: "1", small: 2, committee: "田林一二村、田林五六七村、田林八九十村、田林十三村、田林十一村、千鹤三、爱建园、万科华尔兹、新苑一二、新苑三四、新苑五六七、田林十四村第一、田林十四村第二" },
  { id: 17, name: "樱花园幼儿园", area: "康健/桂林", toddler: "2", small: 3, committee: "月季百藤、樱花园、茶花桂花、丁香迎春、玉兰园、寿益坊（海上华庭）、紫鹃园、紫薇园（牡丹园、玫瑰园）、金桂苑、紫荆党校" },
  { id: 18, name: "长海幼儿园", area: "康健/桂林", toddler: "混龄式招生", small: 2, committee: "长顺海、长虹坊、长青坊、长丰坊、长兴坊、师大新村、桂康" },
  { id: 19, name: "康沁幼儿园", area: "康健/桂林", toddler: "混龄式招生", small: 2, committee: "寿昌山、寿祥坊、寿益坊（寿益坊、桂林西街101弄）、康乐、桂二、冠生园、康宁馨、康强海（康强坊）" },
  { id: 20, name: "长桥第一幼儿园", area: "长桥", toddler: "混龄式招生", small: 2, committee: "长桥一村、长桥五村（书香逸居、长桥五村）、闵朱、舒城" },
  { id: 21, name: "长桥第二幼儿园", area: "长桥/凌云", toddler: "2", small: 4, committee: "长桥新村一、长桥新二村、长桥七村、兴荣苑、和平、龙州、梅陇十一第一、梅陇十一第二、陇南、闵秀" },
  { id: 22, name: "长桥第三幼儿园", area: "长桥", toddler: "2", small: 5, committee: "长桥三村一、长桥三村二（长桥三村东区）、长桥八村、平福、长桥四村" },
  { id: 23, name: "园南幼儿园", area: "长桥/园南", toddler: "1", small: 3, committee: "园南一村、园南二村、园南三村、汇成五村、华东一" },
  { id: 24, name: "梅陇幼儿园", area: "凌云/梅陇", toddler: "2", small: 5, committee: "梅三、梅六、梅苑一、梅苑二、凌云、梅四、理工一、理工二、华理苑、书香苑、家乐苑、阳光（阳光新景）" },
  { id: 25, name: "梅陇第二幼儿园", area: "凌云/梅陇", toddler: "混龄式招生", small: 2, committee: "梅陇五村、梅陇七村、梅陇八村、梅陇十村、梅陇九村、阳光（阳光绿园）、长陇苑" },
  { id: 26, name: "华建幼儿园", area: "华泾", toddler: "1", small: 3, committee: "华建、建华村、北杨村" },
  { id: 27, name: "位育幼儿园", area: "华泾", toddler: "2", small: 2, committee: "华泾四村、华泾五村、漓江山水、联合居委" },
  { id: 28, name: "果果幼儿园", area: "华泾", toddler: "1", small: 4, committee: "华欣家园、华发、馨宁、大桥、明丰新纪苑、名苑、华臻" },
  { id: 29, name: "星辰幼儿园", area: "华泾/罗秀", toddler: "1", small: 2, committee: "徐汇新城、罗秀三村、华滨家园、华沁家园、罗秀、罗秀二村" },
  { id: 30, name: "徐汇实验幼儿园", area: "华泾/罗秀", toddler: "1", small: 2, committee: "中海瀛台、百龙、港口" },
  { id: 31, name: "印象幼儿园", area: "华泾", toddler: "1", small: 2, committee: "印象、华阳、沙家浜" },
  { id: 32, name: "科技逸夫幼儿园", area: "南站/石龙", toddler: "1", small: 1, committee: "正南、东荡、南站居委（临）" },
  { id: 33, name: "盛华幼儿园", area: "华泾", toddler: "1", small: 3, committee: "盛华、华泾绿苑、光华绿苑" },
  { id: 34, name: "艺树幼儿园", area: "徐家汇/田林", toddler: "2", small: 3, committee: "锦馨苑、千鹤第六、小安桥、华鼎广场、千鹤五、尚汇豪庭、千鹤一、千鹤二；另向徐家汇街道、田林街道扩招" },
  { id: 35, name: "桂平幼儿园", area: "漕河泾/虹梅", toddler: "2", small: 1, committee: "联莘、欣园" },
  { id: 36, name: "田林第六幼儿园", area: "田林/虹梅", toddler: "3", small: 5, committee: "田林十二村、田林三四村、古一、古二、古三、古四、东兰、航天新苑、永兆、漕河泾开发园区（临）；另向田林街道、虹梅路街道扩招" },
  { id: 37, name: "紫薇实验幼儿园", area: "康健/漕河泾", toddler: "5", small: 6, committee: "紫薇园（桂平路123弄）、康强海（海上名邸）；区域内自主招生" },
  { id: 38, name: "上海幼儿园", area: "长桥/凌云", toddler: "2", small: 5, committee: "体育花苑、长桥五村（上中路100弄）、长桥三村二（尚海悦庭）；另向凌云路街道、长桥街道、华泾镇扩招" },
  { id: 39, name: "乌鲁木齐南路幼儿园", area: "衡复/湖南", toddler: "3", small: 4, committee: "区域内自主招生" },
  { id: 40, name: "科技幼儿园", area: "徐家汇/田林", toddler: "3", small: 9, committee: "区域内自主招生" },
  { id: 41, name: "宛南实验幼儿园", area: "龙华/滨江", toddler: "6", small: 6, committee: "区域内自主招生" },
  { id: 42, name: "机关建国幼儿园", area: "衡复/湖南/滨江", toddler: "2", small: 6, committee: "区域内自主招生" },
];

const campusRows = [
  [1, "复中园", "复兴中路1260弄2号", "A", ""],
  [1, "南园", "襄阳南路317号", "A", ""],
  [1, "陕南园", "陕西南路540号", "A", ""],
  [1, "北园", "襄阳南路207号", "A", ""],
  [2, "永嘉园", "永嘉路420号", "A", ""],
  [2, "武康园", "五原路400号", "A", ""],
  [2, "兴国园", "武康路280弄24号", "A", ""],
  [3, "南园", "宛平南路19弄3号", "A", ""],
  [3, "北园", "康平路200号", "B", "2024公办名单写康平路200号；2024招生简章/部分旧资料写华山路1815号，建议电话确认当前小班实际园区。"],
  [4, "本部", "乐山路18号", "A", ""],
  [5, "北园", "番禺路800弄24号", "A", ""],
  [5, "南园", "番禺路1188号", "A", ""],
  [6, "本部", "小木桥路440弄30号", "A", ""],
  [7, "托小班部", "平江路17号", "B", "2024年7月由原平江路幼儿园与原复旦大学医学院幼儿园合并组建；2025/公开资料显示新小班启用平江路32号，需确认2026实际安排。"],
  [7, "中大班部/新小班候选", "平江路32号", "B", "公开资料显示办学地址为平江路32号、17号；原复旦医学院幼儿园东安路50弄10号为历史/合并相关地址。"],
  [8, "零陵园", "零陵路250弄33号", "A", ""],
  [8, "东安园", "东安一村39号", "A", "公办名单列出东安一村39号及零陵路250弄33号。"],
  [9, "一分园", "宛南四村16号", "A", ""],
  [9, "二分园", "中山南二路999弄5号", "A", ""],
  [9, "本部", "天钥新村90号", "A", ""],
  [10, "丰谷园", "丰谷路205弄34号", "A", ""],
  [10, "龙恒园", "龙华西路31弄15号", "A", ""],
  [10, "龙华园", "龙华西路285弄14号", "A", ""],
  [11, "吴中园", "吴中东路500弄67号", "A", ""],
  [11, "钦州园", "钦州北路898号", "A", ""],
  [12, "中大班部", "天钥桥南路1249弄11号", "A", ""],
  [12, "小班部", "龙水南路龙南三村7号", "A", ""],
  [13, "中大班部", "龙漕路139号", "A", ""],
  [13, "小班部", "漕东路193号", "A", ""],
  [14, "冠生园", "冠生园路28号", "A", ""],
  [14, "南宁园", "南宁路636号", "A", ""],
  [15, "楼园园", "老沪闵路706弄37号", "A", ""],
  [15, "金塘园", "老沪闵路333弄70号", "A", ""],
  [16, "东园", "田林九村6号", "A", ""],
  [16, "西园", "宜山路701弄53号", "A", ""],
  [17, "南园", "百花街398号", "A", ""],
  [17, "北园", "虹漕南路杨家桥88号", "A", ""],
  [18, "本部", "桂林西街15弄2号", "A", ""],
  [19, "本部", "桂林西街151弄20号甲", "A", ""],
  [20, "本部", "长桥一村56号", "A", ""],
  [21, "凌云园", "梅陇十一村97号", "A", ""],
  [21, "长桥园", "长桥二村34号", "A", ""],
  [22, "长桥园", "长桥三村124号", "A", ""],
  [22, "平福园", "上中路483弄32号", "A", ""],
  [23, "本部", "龙川北路园南一村27号", "A", ""],
  [24, "嘉川园", "梅陇四村56号甲", "A", ""],
  [24, "梅陇园", "梅陇六村65号", "A", ""],
  [25, "本部", "梅陇五村54号", "A", ""],
  [26, "本部", "老沪闵路1300号", "A", ""],
  [27, "本部", "建华路102号", "A", ""],
  [28, "华欣园", "龙吴路2422号", "A", ""],
  [28, "华发园", "华发路100弄22号", "A", ""],
  [29, "本部", "罗秀路11号", "A", ""],
  [30, "本部", "龙瑞路135号", "A", ""],
  [31, "本部", "望月路882号", "A", ""],
  [32, "本部", "石龙路818弄8号", "A", ""],
  [33, "本部", "望月路401号", "A", ""],
  [34, "本部", "古井路160号", "A", ""],
  [35, "本部", "桂平路260弄14号", "A", ""],
  [36, "本部", "田林十二村40号", "A", ""],
  [36, "贝贝分园", "田林四村18号", "A", ""],
  [36, "东兰分园", "古美路1107弄65号", "A", ""],
  [37, "桂平园", "桂平路123弄23号", "A", ""],
  [37, "浦北园", "浦北路173号", "A", ""],
  [37, "全州园", "宜州路26号", "A", ""],
  [38, "凌云园", "上中西路378号", "A", ""],
  [38, "冠军园", "老沪闵路729弄41号乙", "A", ""],
  [38, "上中园", "上中路402号", "A", ""],
  [39, "本部", "乌鲁木齐南路14号", "A", ""],
  [39, "境内部", "淮海路1788号", "A", ""],
  [39, "境外部", "淮海中路1480号", "A", ""],
  [40, "宜山园1部", "宜山路655弄1号", "A", ""],
  [40, "嘉陵园", "嘉陵路28号", "A", ""],
  [40, "文定园", "文定路476号", "A", ""],
  [40, "宜山园10部", "宜山路655弄10号", "A", ""],
  [41, "瑞宁部", "瑞宁路816号", "A", ""],
  [41, "滨江部", "瑞宁路851号", "A", ""],
  [41, "大木桥部", "大木桥路323号", "A", ""],
  [42, "建国园", "建国西路570号", "A", ""],
  [42, "安亭园", "安亭路112号", "A", ""],
  [42, "滨江园", "云锦路183弄30号", "A", ""],
];

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

const publicCampusPhoneByKey = new Map([
  ["1|复中园", "021-64335881"],
  ["1|南园", "021-64335881"],
  ["1|陕南园", "021-64335881"],
  ["1|北园", "64335881"],
  ["2|永嘉园", "64759122"],
  ["2|武康园", "54035366"],
  ["2|兴国园", "64360627"],
  ["3|南园", "021-64689142"],
  ["3|北园", "021-64689142"],
  ["4|本部", "64072867"],
  ["5|北园", "62837361"],
  ["5|南园", "64860098"],
  ["6|本部", "64035845"],
  ["7|托小班部", "64034223"],
  ["7|中大班部/新小班候选", "64034223"],
  ["8|零陵园", "64041758"],
  ["8|东安园", "64175625"],
  ["9|一分园", "64383244"],
  ["9|二分园", "64869504"],
  ["9|本部", "64383295"],
  ["10|丰谷园", "64280780"],
  ["10|龙恒园", "64576761"],
  ["10|龙华园", "64571589"],
  ["11|吴中园", "62196848"],
  ["11|钦州园", "64668657"],
  ["12|中大班部", "54095768"],
  ["12|小班部", "021-34604160"],
  ["13|中大班部", "64823976"],
  ["13|小班部", "34616320"],
  ["14|冠生园", "64751426"],
  ["14|南宁园", "64751426"],
  ["15|楼园园", "021-64234155"],
  ["15|金塘园", "54963078"],
  ["16|东园", "64361935"],
  ["16|西园", "64361935"],
  ["17|南园", "021-54181491"],
  ["17|北园", "54362050"],
  ["18|本部", "54208285"],
  ["19|本部", "54363105"],
  ["20|本部", "64100149"],
  ["21|凌云园", "34318987"],
  ["21|长桥园", "64102557"],
  ["22|长桥园", "64102397"],
  ["22|平福园", "64102397"],
  ["23|本部", "64763190"],
  ["24|嘉川园", "64100015"],
  ["24|梅陇园", "64100015"],
  ["25|本部", "64775092"],
  ["26|本部", "2164548743"],
  ["27|本部", "64960505"],
  ["28|华欣园", "54829893"],
  ["28|华发园", "54829893"],
  ["29|本部", "54010448"],
  ["30|本部", "54045192"],
  ["31|本部", "64960096"],
  ["32|本部", "64400295"],
  ["33|本部", "021-54852338"],
  ["34|本部", "64388770"],
  ["35|本部", "64031921"],
  ["36|本部", "64366380"],
  ["36|贝贝分园", "64367248"],
  ["36|东兰分园", "64707260"],
  ["37|桂平园", "54217515"],
  ["37|浦北园", "64830788"],
  ["37|全州园", "54500996"],
  ["38|凌云园", "64106663 / 64100564 / 64250961"],
  ["38|冠军园", "64250961"],
  ["38|上中园", "64106663 / 64100564 / 64250961"],
  ["39|本部", "64319939"],
  ["39|境内部", "64319939"],
  ["39|境外部", "64330160"],
  ["40|宜山园1部", "64854450 / 64362975"],
  ["40|嘉陵园", "54363003"],
  ["40|文定园", "64221181"],
  ["40|宜山园10部", "64854450 / 64362975"],
  ["41|瑞宁部", "64363606"],
  ["41|滨江部", "64161261"],
  ["41|大木桥部", "64168714"],
  ["42|建国园", "64372841"],
  ["42|安亭园", "64333332"],
  ["42|滨江园", "54252106"],
]);

const privateCampusRows = [
  { name: "杜鹃园幼稚园", level: "一级", category: "民办（类型待核验）", address: "桂林西街9弄57号", phone: "021-54202188", source: publicListSource },
  { name: "胡姬港湾幼儿园", level: "一级", category: "民办（类型待核验）", address: "丰谷路205弄35号", phone: "64560614", source: publicListSource },
  { name: "汇城苑幼稚园", level: "一级", category: "普惠民办", address: "百色路汇城五村75号", phone: "64235119", source: publicListSource },
  { name: "维亚幼儿园", level: "二级", address: "华亭路71弄1号", phone: "54036901" },
  { name: "四季方馨幼儿园", level: "二级", address: "五原路112号", phone: "021-54036603" },
  { name: "滨江幼儿园", level: "二级", address: "云锦路80弄10号", phone: "64380700" },
  { name: "汇宝幼儿园", level: "二级", address: "宛平南路592号", phone: "64690445" },
  { name: "蒙世学堂幼儿园", level: "二级", address: "斜土路2421号", phone: "021-64686261*8103" },
  { name: "樱花园幼稚园", level: "二级", address: "百花街380号", phone: "54185106" },
  { name: "新宜幼稚园", level: "二级", address: "古宜路170弄7号", phone: "64684874" },
  { name: "蓓蕾幼稚园", campus: "本部", level: "二级", address: "蒲江塘路50号玉兰花苑6幢", phone: "64647491" },
  { name: "蓓蕾幼稚园", campus: "分园", level: "二级", address: "天钥桥路380弄11号", phone: "64647491" },
  { name: "淇莲幼儿园", level: "二级", address: "浦北路50号", phone: "54669909" },
  { name: "田林街道中心幼儿园", level: "二级", address: "田林六村10号", phone: "64705223" },
  { name: "胡姬港湾新汇幼儿园", level: "二级", address: "浦北路21弄42号", phone: "54669810" },
  { name: "小神童幼儿园", level: "二级", category: "普惠民办", address: "龙吴路988弄25号", phone: "021-54360296" },
  { name: "安琪曈幼稚园", level: "二级", address: "衡山路9弄2号", phone: "64665309" },
  { name: "鲁浦幼儿园", campus: "本部", level: "二级", address: "宛南一村23号", phone: "64384206" },
  { name: "鲁浦幼儿园", campus: "分园", level: "二级", address: "宛南五村1号", phone: "64384206" },
  { name: "康文云锦幼儿园", level: "二级", address: "龙兰路398号", phone: "64282682" },
  { name: "爱文幼儿园", level: "二级", address: "古羊路160号1幢", phone: "021-62090135" },
  { name: "童稻幼儿园", level: "二级", address: "淮海西路365弄2号楼、3号楼", phone: "52668270" },
  { name: "嘉宝幼儿园", level: "二级", address: "吴兴路75号", phone: "64373773" },
  { name: "金贝贝幼儿园", level: "二级", address: "龙吟路300号", phone: "54820000" },
  { name: "世纪昂立幼儿园", level: "二级", address: "龙山新村115号（近零陵路）", phone: "54890979 / 64382277" },
  { name: "爱菊幼儿园", level: "二级", address: "复兴西路70号", phone: "64043162" },
  { name: "培蕾幼稚园", campus: "本部", level: "二级", address: "梅陇三村49号", phone: "64768480" },
  { name: "培蕾幼稚园", campus: "分园", level: "二级", address: "梅陇六村41号", phone: "64767077" },
  { name: "枫叶交响幼儿园", level: "二级", address: "太原路87号", phone: "64730053" },
  { name: "泰宁田林幼儿园", level: "二级", address: "田林十一村36号", phone: "64755118-804" },
  { name: "田林东方幼儿园", level: "二级", address: "宜山路田林十四村27号", phone: "021-64085912" },
  { name: "澳宝幼儿园", level: "二级", address: "永嘉路356弄31号", phone: "64720200" },
  { name: "吉的堡小蜻蜓幼儿园", level: "二级", address: "虹梅路1035弄30号", phone: "64368108" },
  { name: "东泉大地幼儿园", level: "二级", address: "东泉路65弄9号", phone: "54084130" },
  { name: "吉的堡新徐汇幼儿园", level: "二级", address: "小木桥路101弄20号", phone: "021-54245100" },
  { name: "爱悠小红花幼儿园", level: "二级", address: "梅陇路130号1幢、2幢", phone: "021-54333289" },
  { name: "中山幼儿园", level: "二级", address: "桃江路42号", phone: "021-33565515" },
  { name: "创意幼儿园", level: "二级", address: "柳州路田林十村6号", phone: "64828813 / 64820881" },
  { name: "牛牛幼稚园", level: "二级", address: "罗城路700弄95号", phone: "54115988" },
  { name: "领幼幼儿园", level: "二级", address: "天钥桥路1057弄3号", phone: "64458520" },
  { name: "漕河泾新汇幼儿园", level: "二级", address: "漕泾一村30号", phone: "64364348" },
  { name: "凯琴数码幼儿园", level: "二级", category: "普惠民办", address: "华泾路995号", phone: "54822626" },
  { name: "陇龙幼稚园", level: "二级", address: "梅陇十一村96号", phone: "64549910" },
  { name: "世蒙幼儿园", level: "二级", address: "东湖路21号", phone: "021-54038979" },
  { name: "蒂伊幼稚园", level: "二级", address: "永嘉路383号", phone: "64749388" },
  { name: "吉的堡新汇幼儿园", level: "二级", address: "康健小区虹漕南路百花街18号、6号", phone: "54183331" },
];

const byId = new Map(schools.map((school) => [school.id, school]));
const mapSearch = (name, campus, address) => {
  const query = `上海市徐汇区 ${name} ${campus} ${address}`;
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
    nature: "公办",
    category: "公办",
    level: schoolLevel,
    name: school.name,
    area: school.area,
    campus,
    address,
    mapUrl: mapSearch(school.name, campus, address),
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
    nature: "民办",
    category: item.category || "民办（类型待核验）",
    level: item.level,
    name: item.name,
    area,
    campus,
    address: item.address,
    mapUrl: mapSearch(item.name, campus, item.address),
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

const campusItems = [...publicCampusItems, ...privateCampusItems];

const campusHeader = ["编号", "性质", "办园类型", "办园等级", "幼儿园", "片区", "园区/分园", "地址", "高德POI名称", "高德POI地址", "高德经纬度", `距${officeLocation.name}`, "高德搜索链接", "联系电话", "托班计划", "小班计划", "对口居委/招生范围", "招生类型", "置信度", "备注", "主要来源"];
const campusData = campusItems.map((item) => [
  item.id,
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
  ["公办招生主体数", schools.length, "来自2026 PDF。"],
  ["公办园区/分园点位数", publicCampusItems.length, "按公开园部地址拆分；少数历史变更项标B级。"],
  ["民办/私立点位数", privateCampusItems.length, "来自徐汇区民办幼儿园名单公开资料；招生条件、收费和名额需电话确认。"],
  ["全部园区/点位数", campusData.length, "公办园区点位 + 民办/私立点位。"],
  ["高德增强字段", "已接入", "已批量补POI名称、经纬度和到网易上海西岸研发中心的高德直线距离；低置信匹配保留为待核验。"],
  ["小班计划合计", schools.reduce((sum, s) => sum + s.small, 0), "PDF计划班级数合计。"],
  ["托班明确班级数合计", schools.reduce((sum, s) => sum + (/^\d+$/.test(s.toddler) ? Number(s.toddler) : 0), 0), "不含混龄式招生。"],
  ["混龄式招生主体数", schools.filter((s) => `${s.toddler}`.includes("混龄")).length, "混龄式招生不直接等同托班班级数。"],
  ["重点核验", "汇星北园、复旦大学附属徐汇实验幼儿园", "公开资料出现地址/合并后的园部安排差异，择园前应电话确认。"],
];

const markdown = `# 徐汇区幼儿园园区位置与择园参考

## 结论摘要

- 2026 PDF 共列出 ${schools.length} 个公办招生主体，拆分为 ${publicCampusItems.length} 个公办实际园区/分园点位。
- 已补充 ${privateCampusItems.length} 个徐汇区民办/私立点位，包含级别、地址和联系电话；这些点位更适合作为当前暂无居住证情况下的兜底池。
- 已通过高德服务批量增强 POI 名称、经纬度和到${officeLocation.name}的直线距离；未可靠匹配项保留为待核验。
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
const beikeSearch = (keyword) => `https://sh.zu.ke.com/zufang/rs${encodeURIComponent(keyword)}/`;
const itemDistance = (item) => item?.officeDistance && item.officeDistance !== pendingAmapValue ? item.officeDistance : "高德未匹配，待电话/地图确认";
const poiLine = (item) => item?.amapPoiName && item.amapPoiName !== pendingAmapValue
  ? `${item.amapPoiName}${item.amapPoiAddress && item.amapPoiAddress !== pendingAmapValue ? `；${item.amapPoiAddress}` : ""}`
  : "高德POI未可靠匹配";

const decisionRecommendations = [
  {
    rank: "公办 1",
    tag: "首选争取",
    tagClass: "green",
    type: "公办 / 一级",
    item: campusRef("公办", "园南幼儿园", "本部", "龙川北路园南一村27号"),
    rentArea: "园南一村 / 汇成五村 / 上中路",
    beikeKeyword: "园南 长桥 3室 100平 10000",
    price: "贝壳实时筛选：3室、100平以上、10000以内",
    judgement: "租房预算、居住环境和公办争取的平衡最好，是当前主线。",
  },
  {
    rank: "公办 2",
    tag: "备选争取",
    tagClass: "blue",
    type: "公办 / 一级",
    item: campusRef("公办", "徐汇实验幼儿园", "本部", "龙瑞路135号"),
    rentArea: "华沁家园 / 华滨家园 / 龙瑞路",
    beikeKeyword: "华沁家园 华滨家园 3室 100平 10000",
    price: "贝壳实时筛选：华泾北大三房，优先电梯",
    judgement: "距离公司近于长桥，面积更容易做大，但需严格反查居委。",
  },
  {
    rank: "公办 3",
    tag: "近公司",
    tagClass: "amber",
    type: "公办 / 二级",
    item: campusRef("公办", "阳光幼儿园", "小班部", "龙水南路龙南三村7号"),
    rentArea: "龙南三四村 / 樟树苑 / 龙水南路",
    beikeKeyword: "龙南 龙水南路 3室 100平 10000",
    price: "贝壳实时筛选：离公司近，但大户型低预算不稳定",
    judgement: "高德距离最近，适合捡漏；不要作为唯一租房主线。",
  },
  {
    rank: "公办 4",
    tag: "容量好",
    tagClass: "blue",
    type: "公办 / 示范园",
    item: campusRef("公办", "上海幼儿园", "上中园", "上中路402号"),
    rentArea: "上中路 / 长桥 / 植物园",
    beikeKeyword: "上中路 长桥 3室 100平 10000",
    price: "贝壳实时筛选：长桥南侧大户型，优先材料配合",
    judgement: "示范园且片区与租房需求匹配，但公办竞争和材料门槛仍高。",
  },
  {
    rank: "公办 5",
    tag: "通勤好",
    tagClass: "amber",
    type: "公办 / 一级",
    item: campusRef("公办", "龙华幼儿园", "龙恒园", "龙华西路31弄15号"),
    rentArea: "龙华西路 / 龙恒路 / 云锦路",
    beikeKeyword: "龙华西路 龙恒路 3室 100平 10000",
    price: "贝壳实时筛选：通勤优秀，但预算更容易被卡",
    judgement: "距离公司好，但租金和公办顺位压力都更大，作为机会项。",
  },
  {
    rank: "民办 1",
    tag: "主兜底",
    tagClass: "green",
    type: "普惠民办 / 一级",
    item: campusRef("民办", "汇城苑幼稚园", "本部", "百色路汇城五村75号"),
    rentArea: "汇成五村 / 园南 / 百色路",
    beikeKeyword: "汇成五村 园南 3室 100平 10000",
    price: "贝壳实时筛选：和园南主线共用看房区域",
    judgement: "最适合和园南/长桥公办争取线并行，优先电话确认名额。",
  },
  {
    rank: "民办 2",
    tag: "近公司",
    tagClass: "green",
    type: "民办 / 一级",
    item: campusRef("民办", "胡姬港湾幼儿园", "本部", "丰谷路205弄35号"),
    rentArea: "龙华 / 云锦路 / 丰谷路",
    beikeKeyword: "云锦路 丰谷路 3室 100平 10000",
    price: "贝壳实时筛选：近公司，预算需严格筛",
    judgement: "高德距离近；名单地址与高德POI地址有差异，必须电话确认实际园区。",
  },
  {
    rank: "民办 3",
    tag: "长桥兜底",
    tagClass: "blue",
    type: "民办 / 二级",
    item: campusRef("民办", "牛牛幼稚园", "本部", "罗城路700弄95号"),
    rentArea: "罗城路 / 长桥 / 上中西路",
    beikeKeyword: "罗城路 长桥 3室 100平 10000",
    price: "贝壳实时筛选：长桥南侧更容易找大户型",
    judgement: "和长桥租房主线兼容，适合作为第二民办兜底。",
  },
  {
    rank: "民办 4",
    tag: "华泾兜底",
    tagClass: "blue",
    type: "普惠民办 / 二级",
    item: campusRef("民办", "凯琴数码幼儿园", "本部", "华泾路995号"),
    rentArea: "华泾 / 华沁 / 华滨 / 罗秀",
    beikeKeyword: "华泾 华沁 3室 100平 10000",
    price: "贝壳实时筛选：华泾方向面积更友好",
    judgement: "适合华泾北/罗秀租房线，距离公司中等，需确认名额和收费。",
  },
  {
    rank: "民办 5",
    tag: "康健备选",
    tagClass: "amber",
    type: "民办 / 一级",
    item: campusRef("民办", "杜鹃园幼稚园", "本部", "桂林西街9弄57号"),
    rentArea: "桂林西街 / 康健 / 田林南",
    beikeKeyword: "桂林西街 康健 3室 100平 10000",
    price: "贝壳实时筛选：生活成熟，通勤略远于龙华",
    judgement: "一级民办，适合作为康健/田林方向的兜底备选。",
  },
];

const renderDecisionRows = () => decisionRecommendations.map((row) => `
            <tr>
              <td><span class="tag ${row.tagClass}">${escapeHtml(row.rank)}</span><br><span class="sub">${escapeHtml(row.tag)}</span></td>
              <td>${escapeHtml(row.type)}</td>
              <td><strong>${escapeHtml(row.item?.name || "")} ${escapeHtml(row.item?.campus || "")}</strong><br><span class="sub">${escapeHtml(row.item?.address || "")}</span></td>
              <td><strong>${escapeHtml(row.item?.phone || "待电话确认")}</strong><br><span class="sub">${escapeHtml(poiLine(row.item))}</span></td>
              <td><strong>${escapeHtml(row.rentArea)}</strong></td>
              <td><a href="${escapeHtml(beikeSearch(row.beikeKeyword))}" target="_blank" rel="noopener">贝壳实时搜索</a><br><span class="sub">${escapeHtml(row.price)}</span></td>
              <td><span class="distance-list"><span><b>园到公司：</b>${escapeHtml(itemDistance(row.item))}</span><span><b>口径：</b>高德 POI 直线距离</span></span></td>
              <td>${escapeHtml(row.judgement)}</td>
            </tr>
`).join("");

const html = `<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>徐汇区幼儿园筛选方案</title>
  <link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'%3E%3Crect width='64' height='64' rx='12' fill='%23235c9f'/%3E%3Cpath d='M16 42h32M20 42V24l12-8 12 8v18M28 42V30h8v12' stroke='white' stroke-width='4' fill='none' stroke-linecap='round' stroke-linejoin='round'/%3E%3C/svg%3E">
  <style>
    :root {
      color-scheme: light;
      --bg: #f6f7f9;
      --paper: #ffffff;
      --ink: #172033;
      --soft: #5f6f85;
      --line: #dce3ec;
      --blue: #235c9f;
      --green: #1f7a5c;
      --amber: #a16207;
      --red: #b42318;
      --rail: #edf2f7;
      --shadow: 0 18px 50px rgba(23, 32, 51, 0.08);
    }
    * { box-sizing: border-box; }
    html { scroll-behavior: smooth; }
    body {
      margin: 0;
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
      border-bottom: 1px solid rgba(220, 227, 236, 0.9);
      background: rgba(246, 247, 249, 0.94);
      backdrop-filter: blur(12px);
    }
    .topbar .shell {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 18px;
      min-height: 64px;
    }
    .brand { font-weight: 800; font-size: 16px; letter-spacing: 0; }
    .nav { display: flex; gap: 16px; color: var(--soft); font-size: 13px; white-space: nowrap; }
    .nav a { color: inherit; }
    .hero {
      background:
        linear-gradient(135deg, rgba(35, 92, 159, 0.12), transparent 34%),
        linear-gradient(180deg, #ffffff 0%, #f6f7f9 100%);
      border-bottom: 1px solid var(--line);
    }
    .hero-grid {
      display: grid;
      grid-template-columns: minmax(0, 1fr) 420px;
      gap: 44px;
      align-items: center;
      min-height: 470px;
      padding: 56px 24px 48px;
    }
    h1 {
      margin: 0;
      font-size: clamp(34px, 5vw, 64px);
      line-height: 1.06;
      letter-spacing: 0;
    }
    .lead {
      max-width: 760px;
      margin: 22px 0 0;
      color: var(--soft);
      font-size: 18px;
      line-height: 1.72;
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
    .notice {
      display: grid;
      grid-template-columns: 1fr 1fr;
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
    footer {
      border-top: 1px solid var(--line);
      color: var(--soft);
      padding: 24px 0 36px;
    }
    @media (max-width: 1100px) {
      .hero-grid, .workflow, .notice, .personal-grid { grid-template-columns: 1fr; }
      .principles, .areas, .judgement-grid, .priority-grid, .action-steps, .logic-compact, .summary-strip, .profile-grid, .todo-board { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      .toolbar { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      .source-list { grid-template-columns: 1fr; }
    }
    @media (max-width: 720px) {
      .shell { padding: 0 14px; }
      .topbar .shell { align-items: flex-start; flex-direction: column; padding-top: 12px; padding-bottom: 12px; }
      .nav { overflow: auto; width: 100%; padding-bottom: 2px; }
      .hero-grid { padding: 38px 14px 34px; min-height: auto; }
      .scoreboard { padding: 14px; }
      .principles, .areas, .matrix, .toolbar, .judgement-grid, .priority-grid, .action-steps, .logic-compact, .summary-strip, .profile-grid, .todo-board { grid-template-columns: 1fr; }
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
      <div class="brand">徐汇区幼儿园筛选方案</div>
      <nav class="nav" aria-label="页面导航">
        <a href="#profile">基本情况</a>
        <a href="#decision">结论</a>
        <a href="#priority">推荐板块</a>
        <a href="#todos">待办</a>
        <a href="#query">园区查询</a>
        <a href="#sources">来源</a>
      </nav>
    </div>
  </header>

  <section class="hero">
    <div class="shell hero-grid">
      <div>
        <h1>长桥/植物园优先，先把材料链补齐。</h1>
        <p class="lead">基于你刚来上海、暂无居住证、办公点在西岸网易研发中心、需要 3 房 100 平以上租房的情况，当前主线不是冲热门公办园，而是先选更可租、更可办材料、更适合接送的居住板块。</p>
        <div class="hero-actions">
          <a class="button primary" href="#decision">看明确结论</a>
          <a class="button" href="#query">进入园区查询</a>
          <a class="button" href="徐汇区幼儿园园区位置与择园参考.xlsx">打开 Excel 表</a>
        </div>
      </div>
      <aside class="scoreboard" aria-label="数据摘要">
        <div class="score-row"><span>公办招生主体</span><strong>${schools.length}</strong></div>
        <div class="score-row"><span>全部园区/点位</span><strong>${campusItems.length}</strong></div>
        <div class="score-row"><span>小班计划合计</span><strong>${schools.reduce((sum, s) => sum + s.small, 0)}</strong></div>
        <div class="score-row"><span>明确托班计划</span><strong>${schools.reduce((sum, s) => sum + (/^\\d+$/.test(s.toddler) ? Number(s.toddler) : 0), 0)}</strong></div>
      </aside>
    </div>
  </section>

  <main class="shell">
    <section id="profile">
      <div class="section-title">
        <h2>基本情况</h2>
        <p>这些条件是后续所有推荐的前提，尤其是居住证和租房材料链会直接影响公办报名顺位。</p>
      </div>
      <div class="profile-grid">
        <article class="profile-card"><strong>户籍/材料</strong><span>刚来上海，暂无居住证；需要先补居住登记、租赁合同/备案和居住证材料。</span></article>
        <article class="profile-card"><strong>办公位置</strong><span>上海市徐汇区西岸网易研发中心，按龙耀路/云锦路/龙水南路附近估算通勤。</span></article>
        <article class="profile-card"><strong>租房需求</strong><span>至少 3 房、100 平以上、预算 10000 以内，优先电梯、小区环境好、遛娃方便。</span></article>
        <article class="profile-card"><strong>决策目标</strong><span>先选能办材料且通勤可接受的居住板块，再反查居委和可报名幼儿园。</span></article>
      </div>
    </section>

    <section id="decision">
      <div class="section-title">
        <h2>首页结论</h2>
        <p>这里直接给到 5 所公办争取、5 所民办兜底，距离使用高德 POI 到网易上海西岸研发中心的直线距离。</p>
      </div>
      <div class="summary-strip">
        <article class="summary-card urgent">
          <strong>最大风险</strong>
          <span>暂无居住证会明显影响公办报名材料完整性和录取顺位，公办小班概率当前偏低。</span>
        </article>
        <article class="summary-card">
          <strong>主推荐</strong>
          <span>公办争取以园南、徐汇实验、阳光、上海幼儿园、龙华幼儿园为主。</span>
        </article>
        <article class="summary-card">
          <strong>备选</strong>
          <span>民办兜底以汇城苑、胡姬港湾、牛牛、凯琴、杜鹃园并行联系。</span>
        </article>
        <article class="summary-card">
          <strong>第一行动</strong>
          <span>看房前先问房东是否配合居住登记、租赁合同/备案和居住证材料。</span>
        </article>
      </div>
      <div class="recommendation-table">
        <table>
          <thead>
            <tr>
              <th>优先级</th>
              <th>园所类型</th>
              <th>推荐幼儿园/分园</th>
              <th>联系电话/来源</th>
              <th>推荐租房小区/板块</th>
              <th>贝壳租房链接与价格</th>
              <th>距离估算</th>
              <th>判断</th>
            </tr>
          </thead>
          <tbody>
${renderDecisionRows()}
          </tbody>
        </table>
      </div>
    </section>

    <section id="judgement">
      <div class="section-title">
        <h2>你的情况判断</h2>
        <p>这四项决定了为什么不建议一上来就住徐汇滨江核心区冲热门公办。</p>
      </div>
      <div class="judgement-grid">
        <article class="judgement-card"><h3>公办概率</h3><p>暂无居住证时，非沪籍公办录取顺位和材料完整性都不占优；公办只能作为争取线。</p></article>
        <article class="judgement-card"><h3>材料链</h3><p>租房地址、居住登记、居住证、租赁合同/备案要尽量一致；房东配合度是硬条件。</p></article>
        <article class="judgement-card"><h3>普惠民办</h3><p>应提升为并行主线，优先联系汇城苑、凯琴、小神童这类离推荐板块近的普惠民办。</p></article>
        <article class="judgement-card"><h3>租房预算</h3><p>1 万内、3 房、100 平以上、优先电梯和好环境，西岸核心区难度高，南部板块更现实。</p></article>
      </div>
      <div class="logic-compact">
        <article class="logic-card"><strong>资格层</strong><span>先看居住证、居住登记、居委和招生范围是否匹配。</span></article>
        <article class="logic-card"><strong>民办层</strong><span>先问是否接受非沪籍/居住登记凭证、剩余名额、收费和开放日。</span></article>
        <article class="logic-card"><strong>便利层</strong><span>看实际园区地址和接送路线，不只看幼儿园总名。</span></article>
        <article class="logic-card"><strong>偏好层</strong><span>最后比较托班、小班容量、老人接送和遛娃空间。</span></article>
      </div>
    </section>

    <section id="priority">
      <div class="section-title">
        <h2>推荐居住板块</h2>
        <p>按“录取材料可行性、租房预算、通勤、接送与遛娃”综合排序。</p>
      </div>
      <div class="priority-grid">
        <article class="priority-card">
          <header><h3>长桥 / 植物园</h3><span class="tag green">首选</span></header>
          <p class="reason">预算和生活便利性最匹配。</p>
          <p>重点看园南、上中路、汇成、长桥新村周边；可反查长桥第二、长桥第三、园南、上海幼儿园等招生范围。</p>
        </article>
        <article class="priority-card">
          <header><h3>华泾北 / 罗秀</h3><span class="tag blue">备选首选</span></header>
          <p class="reason">面积更容易做大，通勤仍可接受。</p>
          <p>关注龙瑞路、罗秀路、华沁/华滨/徐汇新城周边；可看徐汇实验、星辰、果果、盛华、印象等方向。</p>
        </article>
        <article class="priority-card">
          <header><h3>龙华 / 龙南</h3><span class="tag amber">捡漏</span></header>
          <p class="reason">通勤最好，但预算压力最大。</p>
          <p>如果出现 1 万内、100 平以上、房东配合材料的房源可以看；否则不要把它作为唯一主线。</p>
        </article>
        <article class="priority-card">
          <header><h3>徐家汇 / 衡复</h3><span class="tag red">谨慎</span></header>
          <p class="reason">租金高、竞争强、当前阶段不划算。</p>
          <p>除非预算上调或已有稳定材料优势，否则不建议优先投入看房和报名精力。</p>
        </article>
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
          <label class="todo-item"><input type="checkbox" data-todo="private-backup"><span>同步联系普惠民办/民办园<small>优先电话问汇城苑、凯琴、小神童是否有名额。</small></span></label>
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
          <label class="todo-item"><input type="checkbox" data-todo="policy-read"><span>核对当年徐汇招生政策<small>报名条件、验证时间、录取批次每年可能调整。</small></span></label>
          <label class="todo-item"><input type="checkbox" data-todo="online-register"><span>完成信息登记和报名<small>按官方平台要求提交材料。</small></span></label>
          <label class="todo-item"><input type="checkbox" data-todo="private-deposit"><span>确认民办保位节点<small>问清保位费、退费期限和是否影响公办等待。</small></span></label>
          <label class="todo-item"><input type="checkbox" data-todo="verification"><span>准备现场/线上验证材料<small>居住证、租赁材料、户口本、出生证等。</small></span></label>
          <label class="todo-item"><input type="checkbox" data-todo="after-admission"><span>录取后办理体检和入园材料<small>同步安排老人接送路线和备用方案。</small></span></label>
        </article>
      </div>
    </section>

    <section id="query" class="details-section">
      <div class="section-title">
        <h2>园区查询</h2>
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
          <label>招生类型
            <select id="admission">
              <option value="">全部类型</option>
              <option value="固定对口">固定对口</option>
              <option value="区域自主">区域自主</option>
              <option value="扩招">扩招</option>
              <option value="民办招生">民办招生</option>
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
                data-confidence="${escapeHtml(item.confidence)}"
                data-admission="${escapeHtml(item.admissionType)}"
                data-toddler="${escapeHtml(item.toddlerMode)}"
                data-confirm="${item.needsConfirm ? "yes" : "no"}"
                data-text="${escapeHtml(item.searchText)}">
                <td><div class="school-name">${escapeHtml(item.name)}</div><div class="sub">${escapeHtml(item.campus)}</div></td>
                <td><span class="tag ${item.nature === "民办" ? "amber" : "green"}">${escapeHtml(item.nature)}</span></td>
                <td>${escapeHtml(item.category)}</td>
                <td>${escapeHtml(item.level)}</td>
                <td>${escapeHtml(item.area)}</td>
                <td><span class="tag ${item.admissionType === "固定对口" ? "green" : item.admissionType === "扩招" || item.admissionType === "民办招生" ? "amber" : "blue"}">${escapeHtml(item.admissionType)}</span></td>
                <td>${escapeHtml(item.address)}</td>
                <td><div>${escapeHtml(item.amapPoiName)}</div><div class="sub">${escapeHtml(item.amapPoiAddress)}</div></td>
                <td>${escapeHtml(item.officeDistance)}</td>
                <td>${escapeHtml(item.phone)}</td>
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
        <h2>片区容量概览</h2>
        <p>只展示前 12 个供给相对集中的片区，用于判断大方向；个人录取仍以居委、材料和当年政策为准。</p>
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
      </div>
      <div class="source-list">
        <article class="source-card">
          <strong><a href="https://sh.bendibao.com/news/2026415/305294.shtm" target="_blank" rel="noopener">2026 徐汇区幼儿园招生工作方案</a></strong>
          <span>用于判断户籍优先、外省市户籍排序、录取批次和报名验证要求。</span>
        </article>
        <article class="source-card">
          <strong><a href="https://m.sh.bendibao.com/edu/305384.html" target="_blank" rel="noopener">2026 徐汇幼儿园报名材料</a></strong>
          <span>用于核对非本市户籍幼儿、父母居住证、租赁合同和地址一致要求。</span>
        </article>
        <article class="source-card">
          <strong><a href="https://m.sh.bendibao.com/zffw/285087.html" target="_blank" rel="noopener">上海居住证办理条件</a></strong>
          <span>用于确认居住登记、合法稳定住所、租赁合同/备案等材料链。</span>
        </article>
        <article class="source-card">
          <strong><a href="https://sh.bendibao.com/edu/2024325/284155_2.shtm" target="_blank" rel="noopener">徐汇区民办幼儿园名单</a></strong>
          <span>用于补充民办/私立幼儿园的级别、地址和联系电话，作为公办概率偏低时的兜底池。</span>
        </article>
        <article class="source-card">
          <strong><a href="https://sh.zu.ke.com/" target="_blank" rel="noopener">贝壳上海租房</a></strong>
          <span>用于实时搜索 3 室、100 平以上、10000 元以内的整租房源。</span>
        </article>
        <article class="source-card">
          <strong><a href="https://sh.zu.ke.com/wzdt/" target="_blank" rel="noopener">贝壳徐汇租房板块</a></strong>
          <span>用于快速进入徐汇滨江、华泾、龙华、植物园、长桥等板块。</span>
        </article>
        <article class="source-card">
          <strong><a href="https://www.chooffice.com/1357.html" target="_blank" rel="noopener">西岸网易研发中心办公点参考</a></strong>
          <span>用于确认西岸网易研发中心靠近龙耀路、徐汇滨江一带的通勤判断。</span>
        </article>
        <article class="source-card">
          <strong><a href="https://www.xxh-edu.com/brand/57.html" target="_blank" rel="noopener">汇城苑幼稚园公开资料</a></strong>
          <span>用于核对汇城苑幼稚园地址、电话、托幼一体和普惠民办属性。</span>
        </article>
        <article class="source-card">
          <strong><a href="https://www.021school.cn/schools/12780" target="_blank" rel="noopener">凯琴数码幼儿园公开资料</a></strong>
          <span>用于核对凯琴数码幼儿园地址、电话、托幼一体和普惠民办属性。</span>
        </article>
        <article class="source-card">
          <strong><a href="https://www.021school.cn/schools/12790" target="_blank" rel="noopener">小神童幼儿园公开资料</a></strong>
          <span>用于核对小神童幼儿园地址、电话、收费参考和普惠民办属性。</span>
        </article>
      </div>
    </section>
  </main>

  <footer>
    <div class="shell">主要来源：${escapeHtml(planSource)}；${escapeHtml(publicListSource)}；${escapeHtml(privateListSource)}；${escapeHtml(pdfSource)}；2026 徐汇招生政策、上海居住证办理指南、贝壳/58 等公开租房信息。</div>
  </footer>

  <script>
    const search = document.querySelector("#search");
    const area = document.querySelector("#area");
    const admission = document.querySelector("#admission");
    const toddler = document.querySelector("#toddler");
    const confidence = document.querySelector("#confidence");
    const confirmSelect = document.querySelector("#confirm");
    const resultCount = document.querySelector("#resultCount");
    const rows = Array.from(document.querySelectorAll("#rows tr"));
    const todoInputs = Array.from(document.querySelectorAll("[data-todo]"));

    function applyFilters() {
      const q = search.value.trim().toLowerCase();
      let visible = 0;
      for (const row of rows) {
        const okText = !q || row.dataset.text.toLowerCase().includes(q);
        const okArea = !area.value || row.dataset.area === area.value;
        const okAdmission = !admission.value || row.dataset.admission === admission.value;
        const okToddler = !toddler.value || row.dataset.toddler === toddler.value;
        const okConfidence = !confidence.value || row.dataset.confidence === confidence.value;
        const okConfirm = !confirmSelect.value || row.dataset.confirm === confirmSelect.value;
        const show = okText && okArea && okAdmission && okToddler && okConfidence && okConfirm;
        row.hidden = !show;
        if (show) visible += 1;
      }
      resultCount.textContent = visible;
    }

    for (const control of [search, area, admission, toddler, confidence, confirmSelect]) {
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
  </script>
</body>
</html>`;

const csvEscape = (value) => `"${String(value ?? "").replaceAll('"', '""')}"`;
const toCsv = (rows) => rows.map((row) => row.map(csvEscape).join(",")).join("\n");

await fs.mkdir(outputDir, { recursive: true });
await fs.writeFile(path.join(outputDir, "徐汇区幼儿园园区位置表.csv"), toCsv([campusHeader, ...campusData]), "utf8");
await fs.writeFile(path.join(outputDir, "徐汇区幼儿园择园参考.md"), markdown, "utf8");
await fs.writeFile(path.join(outputDir, "徐汇区幼儿园园区位置与择园参考.html"), html, "utf8");
await fs.writeFile(path.join(outputDir, "index.html"), html, "utf8");

const workbook = Workbook.create();
const summary = workbook.worksheets.add("摘要");
const schoolSheet = workbook.worksheets.add("招生主体");
const campusSheet = workbook.worksheets.add("园区点位");
const areaSheet = workbook.worksheets.add("片区统计");
const sourceSheet = workbook.worksheets.add("来源与核验规则");

summary.getRange(`A1:C${summaryRows.length}`).values = summaryRows;
schoolSheet.getRange(`A1:I${schoolData.length + 1}`).values = [schoolHeader, ...schoolData];
campusSheet.getRange(`A1:U${campusData.length + 1}`).values = [campusHeader, ...campusData];
areaSheet.getRange(`A1:F${areaStats.length + 1}`).values = [["片区", "招生主体数", "园区点位数", "小班计划数", "明确托班班级数", "混龄招生主体数"], ...areaStats];
sourceSheet.getRange("A1:B11").values = [
  ["项目", "说明"],
  ["PDF主表", pdfSource],
  ["2026线上主表", planSource],
  ["公办园部地址主来源", publicListSource],
  ["民办/私立地址与电话来源", privateListSource],
  ["高德MCP接入口径", `目标办公点：${officeLocation.name}；高德匹配为“网易上海西岸研发中心”。用关键词搜索和距离测量补齐高德POI、经纬度和距离。`],
  ["高德检索口径", "表格中的高德链接使用“上海市徐汇区 + 幼儿园名 + 园区名 + 地址”生成，用于逐点打开核验。"],
  ["A级", "园部地址在公办园地址清单中明确列出，且无明显冲突。"],
  ["B级", "公开资料有变更、合并或地址冲突，位置大体可定位，但实际入读园区需电话确认。"],
  ["C级", "仅有第三方或地图信息，缺少官方交叉验证。本表目前未使用C级作为主确认。"],
  ["电话确认重点", "汇星幼儿园北园；复旦大学附属徐汇实验幼儿园；区域内自主招生、扩招园、民办/私立园。"],
];

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
campusSheet.getRange("B:D").format.columnWidthPx = 90;
campusSheet.getRange("E:E").format.columnWidthPx = 190;
campusSheet.getRange("F:F").format.columnWidthPx = 120;
campusSheet.getRange("G:H").format.columnWidthPx = 150;
campusSheet.getRange("I:J").format.columnWidthPx = 160;
campusSheet.getRange("K:L").format.columnWidthPx = 120;
campusSheet.getRange("M:M").format.columnWidthPx = 360;
campusSheet.getRange("N:N").format.columnWidthPx = 120;
campusSheet.getRange("O:P").format.columnWidthPx = 88;
campusSheet.getRange("Q:Q").format.columnWidthPx = 520;
campusSheet.getRange("R:S").format.columnWidthPx = 90;
campusSheet.getRange("T:U").format.columnWidthPx = 360;
campusSheet.getRange(`A1:U${campusData.length + 1}`).format.wrapText = true;
campusSheet.freezePanes.freezeColumns(4);

areaSheet.getRange("A:A").format.columnWidthPx = 160;
areaSheet.getRange("B:F").format.columnWidthPx = 110;

sourceSheet.getRange("A:A").format.columnWidthPx = 150;
sourceSheet.getRange("B:B").format.columnWidthPx = 760;
sourceSheet.getRange("A1:B11").format.wrapText = true;

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
