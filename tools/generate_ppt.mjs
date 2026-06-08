import pptxgen from "pptxgenjs";
import fs from "fs";
import path from "path";

const pptx = new pptxgen();
pptx.defineLayout({ name: "WIDE", width: 13.333, height: 7.5 });
pptx.layout = "WIDE";

const base = "C:/Users/Lenovo/Documents/shanghai-hengshan-rural-revitalization";
const C = { dark: "1A1A2E", primary: "2D6A4F", accent: "E6A817", light: "F5F5F0", white: "FFFFFF", gray: "888888", dg: "444444" };
let sn = 0;
function snum(s) { s.addText(`${++sn}`, { x: 12.5, y: 7.05, w: 0.6, h: 0.3, fontSize: 8, color: C.gray, align: "right", fontFace: "Microsoft YaHei" }); }

// === SLIDE 1: COVER ===
let s1 = pptx.addSlide();
s1.background = { fill: C.dark };
s1.addShape("rect", { x: 0, y: 3.3, w: "100%", h: 0.04, fill: { color: C.accent } });
s1.addText("上海市松江区横山村乡村振兴项目", { x: 1, y: 1, w: "85%", h: 1.2, fontSize: 40, bold: true, color: C.white, fontFace: "Microsoft YaHei", align: "center" });
s1.addText("运 营 统 筹 方 案", { x: 1, y: 2.3, w: "85%", h: 0.8, fontSize: 28, color: C.accent, fontFace: "Microsoft YaHei", align: "center" });
s1.addText("总运营商 + 产业组织者 + 公共空间统筹 + 招商管理 + 品牌管理", { x: 1, y: 4, w: "85%", h: 0.5, fontSize: 13, color: C.gray, fontFace: "Microsoft YaHei", align: "center" });
s1.addText("2026年6月  |  编制：运营方", { x: 1, y: 5.8, w: "85%", h: 0.4, fontSize: 12, color: C.gray, fontFace: "Microsoft YaHei", align: "center" });
snum(s1);

// === SLIDE 2: TOC ===
let s2 = pptx.addSlide();
s2.background = { fill: C.white };
s2.addText("目  录", { x: 1, y: 0.5, w: 4, h: 0.8, fontSize: 32, bold: true, color: C.dark, fontFace: "Microsoft YaHei" });
s2.addShape("rect", { x: 1, y: 1.3, w: 2, h: 0.04, fill: { color: C.accent } });
["01  项目总述","02  资源资产底数","03  项目模式与角色分工","04  功能定位与空间策划","05  运营方案","06  投资与收益概算","07  实施路径与政策对接","08  附录：数据支撑"].forEach((item, i) => {
  s2.addText(item, { x: 1.2, y: 1.8 + i * 0.65, w: 8, h: 0.5, fontSize: 17, color: C.dg, fontFace: "Microsoft YaHei" });
});
snum(s2);

// === SLIDE 3: PROJECT OVERVIEW ===
let s3 = pptx.addSlide();
s3.background = { fill: C.white };
s3.addText("01  项目总述", { x: 0.8, y: 0.4, w: 10, h: 0.7, fontSize: 28, bold: true, color: C.primary, fontFace: "Microsoft YaHei" });
s3.addShape("rect", { x: 0.8, y: 1.1, w: 1.5, h: 0.03, fill: { color: C.accent } });
let ovData = [
  [{text:"指标",options:{bold:true,fill:{color:C.primary},color:C.white,fontFace:"Microsoft YaHei"}}, {text:"数据 / 说明",options:{bold:true,fill:{color:C.primary},color:C.white,fontFace:"Microsoft YaHei"}}],
  ["项目定位","环山而居·沪派江南 — 上海近郊乡村振兴示范项目（契合上海2025-2027乡村旅游三年行动方案）"],
  ["村庄面积","约5.1 km² — 上海唯一环山型自然村落，横云山位居村中央"],
  ["核心标签","九峰三泖核心区域 · 沪派江南节点 · 上海14个示范村主题蓝图覆盖区"],
  ["周边客流池","佘山国家旅游度假区年游客量突破1,570万人次（2025年数据），辐射欢乐谷/辰山植物园/广富林等7个4A级景区"],
  ["距离市中心","距人民广场约39km，距虹桥枢纽约28km，G50/G60高速直通"],
  ["可用资产","4个核心点位约10,300㎡ + 40亩林地 + 多栋可盘活农宅"],
  ["核心原则","不拆不建 · 保护肌理 · 复合产品 · 集体+企业+村民共益 · 数字乡村赋能"],
];
s3.addTable(ovData, { x: 0.8, y: 1.4, w: 11.5, colW: [2.5, 9], border: {type:"solid",pt:0.5,color:"CCCCCC"}, fontSize: 11, fontFace: "Microsoft YaHei", color: C.dg });
snum(s3);

// === SLIDE 4: ASSET OVERVIEW TABLE ===
let s4 = pptx.addSlide();
s4.background = { fill: C.white };
s4.addText("02  资源资产底数 — 核心资产总览", { x: 0.8, y: 0.4, w: 10, h: 0.7, fontSize: 26, bold: true, color: C.primary, fontFace: "Microsoft YaHei" });
s4.addShape("rect", { x: 0.8, y: 1.05, w: 1.5, h: 0.03, fill: { color: C.accent } });
let astRows = [
  [{text:"点位",options:{bold:true,fill:{color:C.primary},color:C.white}},{text:"面积",options:{bold:true,fill:{color:C.primary},color:C.white}},{text:"核心优势",options:{bold:true,fill:{color:C.primary},color:C.white}},{text:"租金参考",options:{bold:true,fill:{color:C.primary},color:C.white}}],
  ["①横山大院\n(花园318号)","1,563㎡","上海罕见大型四合院 · 双面环水 · 停车充足 · 主体完好","1元/㎡/天"],
  ["②横山小学\n(黄公望小学)","590㎡\n占地2亩","唐风院落 · 黄公望隐居地IP · 背靠横山","5,000元/亩/年"],
  ["③来云吧\n(逅山咖啡)","814.5㎡","十字路口 · 3栋围合 · 1层大空间 · 交通要道","1元/㎡/天"],
  ["④横云山居32栋","7,398㎡(地上)\n+482㎡(地下)","32栋独立建筑 · 全业态覆盖(吃住聚游乐) · 古风外立面","面议"],
];
s4.addTable(astRows, { x: 0.5, y: 1.4, w: 12.3, colW: [2.6, 2.2, 4.5, 3], border: {type:"solid",pt:0.5,color:"CCCCCC"}, fontSize: 11, fontFace: "Microsoft YaHei", color: C.dg });
s4.addText("可盘活总建面约 10,300㎡  |  覆盖民宿/餐饮/研学/文创/农业观光全业态  |  邻近年1,570万+客流池", { x: 0.8, y: 6.3, w: 11, h: 0.4, fontSize: 12, bold: true, color: C.primary, fontFace: "Microsoft YaHei", align: "center" });
snum(s4);

// === SLIDES 5-8: INDIVIDUAL SITES ===
const sites = [
  { n: "横山大院", sub: "横山会客厅 — 花园318号", a: "1,563.1㎡", d: "上海现存极少大型传统四合院，4栋建筑围合，中央庭院具活动策展潜力。双面环水(南北双向河道)，佘天昆公路直达，停车容量大，建筑主体完好，改造可聚焦场景营造。\n\n交通：沪佘昆线、松江93路、96路\n配套：佘山世茂洲际、精灵之城乐园、天马山集镇1.7km", img: "现场照片/01-横山大院/微信图片_20260608173351.jpg" },
  { n: "横山小学", sub: "黄公望书院 — 花园332号楼", a: "590.45㎡ / 占地2亩", d: "文化历史可追溯至唐朝，元朝黄公望隐居于此，董其昌游历吟诗——自带文化IP基因。1栋2层楼+超大面积围合院子，翻新后设施规整，可塑性极强。\n\n交通：松江28路、96路、67路\n文化赋能：可开设黄公望文化展厅、艺术驻留计划、\"横山雅集\"活动", img: "现场照片/02-横山小学（黄公望小学）/微信图片_20260608173600.jpg" },
  { n: "来云吧", sub: "山前驿 — 逅山咖啡", a: "814.5㎡", d: "佘天昆公路与天横公路交汇处十字要道。3栋建筑围合，1层式大空间层高充裕适合开放式活动展览；2层楼房可灵活划分独立功能区。坐落在横山山北角，道路交汇带来天然客流。\n\n交通：松江28路、96路、67路\n场景：骑行驿站、创意市集、户外音乐会", img: "现场照片/03-来云吧/微信图片_20260608172653.jpg" },
  { n: "横云山居32栋", sub: "横云山居 — 花园332号楼对面", a: "7,398㎡(地上)+482㎡(地下)", d: "32栋主体建筑涵盖餐饮/宴会厅/多功能厅/客房/咖啡厅/农家茶舍/儿童活动中心——覆盖\"吃、住、聚、游、乐\"全链条。古风外立面+毛坯状态，为个性化改造预留充足空间。\n\n总用地：15,291.8㎡\n定位：高端民宿集群(10-15栋)·团建会所·亲子度假村·康养中心", img: "现场照片/04-横云山居32栋/微信图片_20260608172856.jpg" },
];
sites.forEach((site, i) => {
  let s = pptx.addSlide();
  s.background = { fill: C.white };
  s.addText(`02  点位 ${i+1}：${site.n}`, { x: 0.7, y: 0.3, w: 8, h: 0.6, fontSize: 24, bold: true, color: C.primary, fontFace: "Microsoft YaHei" });
  s.addText(site.sub, { x: 0.7, y: 0.85, w: 8, h: 0.3, fontSize: 13, color: C.accent, fontFace: "Microsoft YaHei" });
  s.addShape("rect", { x: 0.7, y: 1.15, w: 1.5, h: 0.03, fill: { color: C.accent } });
  s.addText(`建筑面积：${site.a}`, { x: 0.7, y: 1.5, w: 6.5, h: 0.4, fontSize: 14, bold: true, color: C.dark, fontFace: "Microsoft YaHei" });
  s.addText(site.d, { x: 0.7, y: 2.1, w: 6.3, h: 4, fontSize: 12, color: C.dg, fontFace: "Microsoft YaHei", valign: "top", lineSpacing: 20 });
  let imgPath = path.join(base, site.img);
  if (fs.existsSync(imgPath)) s.addImage({ path: imgPath, x: 7.5, y: 1.5, w: 5.3, h: 4.8, sizing: { type: "cover", w: 5.3, h: 4.8 } });
  snum(s);
});


// === REFERENCE SLIDES (after site slides) ===
const refImgs = [
  "docs/reference_images/reference_1.png",
  "docs/reference_images/reference_2.png",
  "docs/reference_images/reference_3.png",
  "docs/reference_images/reference_4.png",
];
const refSubtitles = [
  "横山大院 — 四合院改造概念意向",
  "横山小学 — 文化研学空间概念意向", 
  "来云吧 — 滨水休闲驿站概念意向",
  "横云山居 — 高端民宿集群概念意向",
];
const refNotes = [
  "意向参考：围合院落改造 · 非遗工坊 · 乡村市集 · 沪派江南风格",
  "意向参考：唐风书院 · 黄公望文化展厅 · 艺术驻留 · 竹木光影",
  "意向参考：滨水咖啡 · 骑行驿站 · 创意市集 · 通透玻璃+木材",
  "意向参考：32栋独立院落 · 古风+现代 · 私汤庭院 · 在地材质",
];

refImgs.forEach((rimg, ri) => {
  let rs = pptx.addSlide();
  rs.background = { fill: C.white };
  rs.addText("参考意向", { x: 0.5, y: 0.2, w: 3, h: 0.5, fontSize: 14, bold: true, color: C.accent, fontFace: "Microsoft YaHei" });
  rs.addText(refSubtitles[ri], { x: 0.5, y: 0.7, w: 8, h: 0.5, fontSize: 22, bold: true, color: C.primary, fontFace: "Microsoft YaHei" });
  rs.addShape("rect", { x: 0.5, y: 1.2, w: 1.5, h: 0.03, fill: { color: C.accent } });
  
  let refPath = path.join(base, rimg);
  if (fs.existsSync(refPath)) {
    rs.addImage({ path: refPath, x: 0.5, y: 1.6, w: 7.5, h: 5.2, sizing: { type: "contain", w: 7.5, h: 5.2 } });
  }
  
  rs.addText(refNotes[ri], { x: 8.3, y: 1.8, w: 4.5, h: 1.5, fontSize: 12, color: C.dg, fontFace: "Microsoft YaHei" });
  rs.addText("※ 以上为概念意向示意图，用于向政府展示改造方向和品质标准。实际设计方案将由品牌方在运营方审核后提交。", { x: 8.3, y: 5.5, w: 4.5, h: 1, fontSize: 9, italic: true, color: C.gray, fontFace: "Microsoft YaHei" });
  snum(rs);
});
// === SLIDE 9: FOUR-PARTY STRUCTURE ===
let s9 = pptx.addSlide();
s9.background = { fill: C.white };
s9.addText("03  项目模式：四方结构", { x: 0.8, y: 0.4, w: 10, h: 0.7, fontSize: 26, bold: true, color: C.primary, fontFace: "Microsoft YaHei" });
s9.addShape("rect", { x: 0.8, y: 1, w: 1.5, h: 0.03, fill: { color: C.accent } });
const boxes = [
  { l: "政府 / 村集体", s: "政策 · 资产授权 · 公共资金 · 监管", x: 0.5, y: 1.6, c: "2C3E50" },
  { l: "投资主体 / 平台公司", s: "公共改造资金 · 工程审计 · 持有资产", x: 7, y: 1.6, c: "34495E" },
  { l: "总运营方（我方）", s: "策划 · 招商 · 品牌管理 · 统一运营 · 营销", x: 0.5, y: 4.5, c: C.primary },
  { l: "入驻品牌方 A/B/C...", s: "自费室内改造 · 独立经营 · 收益分成", x: 7, y: 4.5, c: C.accent },
];
boxes.forEach(b => {
  s9.addShape("roundRect", { x: b.x, y: b.y, w: 5.8, h: 1.8, fill: { color: b.c }, rectRadius: 0.1 });
  s9.addText(b.l, { x: b.x + 0.3, y: b.y + 0.3, w: 5.2, h: 0.6, fontSize: 16, bold: true, color: C.white, fontFace: "Microsoft YaHei" });
  s9.addText(b.s, { x: b.x + 0.3, y: b.y + 0.9, w: 5.2, h: 0.6, fontSize: 12, color: C.light, fontFace: "Microsoft YaHei" });
});
s9.addText("▼ 授权 + 资金", { x: 3.2, y: 3.5, w: 3, h: 0.3, fontSize: 10, color: C.gray, align: "center" });
s9.addText("▼ 委托运营 + 管理费", { x: 3.2, y: 6.4, w: 3, h: 0.3, fontSize: 10, color: C.gray, align: "center" });
s9.addText("◆ \"政府主导 + 平台投资 + 总运营统筹 + 品牌自主经营\" — 四方共建、利益共享", { x: 0.8, y: 6.9, w: 11, h: 0.3, fontSize: 11, italic: true, color: C.gray, fontFace: "Microsoft YaHei", align: "center" });
snum(s9);

// === SLIDE 10: COMPETITIVE DIFFERENTIATION ===
let s10 = pptx.addSlide();
s10.background = { fill: C.white };
s10.addText("04  差异化竞争策略", { x: 0.8, y: 0.4, w: 10, h: 0.7, fontSize: 26, bold: true, color: C.primary, fontFace: "Microsoft YaHei" });
s10.addShape("rect", { x: 0.8, y: 1, w: 1.5, h: 0.03, fill: { color: C.accent } });
let compRows = [
  [{text:"行业常见做法",options:{bold:true,fill:{color:"C0392B"},color:C.white}},{text:"横山村差异化方向",options:{bold:true,fill:{color:C.primary},color:C.white}}],
  ["单体民宿+简单餐饮","四合院精品民宿集群 + 多元业态组合（横云山居32栋整租统筹）"],
  ["泛泛\"乡村体验\"","黄公望IP + 唐宋文化 + \"沪派江南\"沉浸式体验（契合上海三年行动方案）"],
  ["单一招商，各自为战","统一运营统筹，品牌互补，集群联动，共享1,570万+年客流池"],
  ["传统线下导流","\"横山一点通\"数字平台：AR导览+预订+农产品电商+村务+数据驾驶舱"],
  ["重资产自建","轻资产运营模式——品牌方承担室内改造，运营方专注策划与管理"],
];
s10.addTable(compRows, { x: 1, y: 1.4, w: 11.3, colW: [4, 7.3], border: {type:"solid",pt:0.5,color:"CCCCCC"}, fontSize: 12, fontFace: "Microsoft YaHei", color: C.dg });
snum(s10);

// === SLIDE 11: FUNCTIONAL ZONING ===
let s11 = pptx.addSlide();
s11.background = { fill: C.white };
s11.addText("04  功能分区与业态定位", { x: 0.8, y: 0.4, w: 10, h: 0.7, fontSize: 26, bold: true, color: C.primary, fontFace: "Microsoft YaHei" });
s11.addShape("rect", { x: 0.8, y: 1, w: 1.5, h: 0.03, fill: { color: C.accent } });
let zoneRows = [
  [{text:"功能区",options:{bold:true,fill:{color:C.primary},color:C.white}},{text:"核心空间",options:{bold:true,fill:{color:C.primary},color:C.white}},{text:"目标业态",options:{bold:true,fill:{color:C.primary},color:C.white}},{text:"客群定位",options:{bold:true,fill:{color:C.primary},color:C.white}}],
  ["核心接待区","横山大院","游客中心·乡村客厅·非遗工坊·精品茶空间","全客群第一入口"],
  ["文化研学区","横山小学","研学基地·书院·艺术驻地·展览·文创商店","学生/文化/艺术客群"],
  ["滨水休闲区","来云吧","咖啡馆·骑行驿站·创意市集·户外音乐会","骑行/年轻/周边社群"],
  ["高端住宿区","横云山居32栋","高端民宿集群(10-15栋)·团建会所·亲子度假","中高端度假/企业团建"],
  ["农业观光带","40亩林地+田园","采摘体验·农田景观·自然教育径·露营","亲子/研学/自然客群"],
  ["产业配套区","原农机房/针织厂房","农产品加工展示·创客空间·仓储物流","B端商户/创客"],
];
s11.addTable(zoneRows, { x: 0.5, y: 1.3, w: 12.3, colW: [2, 2.5, 4.5, 3.3], border: {type:"solid",pt:0.5,color:"CCCCCC"}, fontSize: 10.5, fontFace: "Microsoft YaHei", color: C.dg });
snum(s11);

// === SLIDE 12: INVESTMENT STRATEGY ===
let s12 = pptx.addSlide();
s12.background = { fill: C.white };
s12.addText("05  运营方案 — 招商节奏与准入标准", { x: 0.8, y: 0.4, w: 10, h: 0.7, fontSize: 26, bold: true, color: C.primary, fontFace: "Microsoft YaHei" });
s12.addShape("rect", { x: 0.8, y: 1, w: 1.5, h: 0.03, fill: { color: C.accent } });
let invRows = [
  [{text:"阶段",options:{bold:true,fill:{color:C.primary},color:C.white}},{text:"时间",options:{bold:true,fill:{color:C.primary},color:C.white}},{text:"招商目标",options:{bold:true,fill:{color:C.primary},color:C.white}},{text:"出租率",options:{bold:true,fill:{color:C.primary},color:C.white}}],
  ["预热期","第1-3个月","锁定2-3家头部锚定品牌（民宿/餐饮头部运营商）","20%"],
  ["集中招商","第4-8个月","完成主要空间招商，4大点位品牌落位","60%"],
  ["补充招商","第9-12个月","补齐特色业态（非遗/文创/轻食/运动体验）","85%"],
  ["稳定运营","第2年起","品牌动态优化迭代，储备品牌库维护","90%+"],
];
s12.addTable(invRows, { x: 0.8, y: 1.4, w: 11.5, colW: [2.2, 2.2, 4.8, 2.3], border: {type:"solid",pt:0.5,color:"CCCCCC"}, fontSize: 12, fontFace: "Microsoft YaHei", color: C.dg });
s12.addText("品牌准入标准：合法经营资质 · 业态与功能区匹配 · 室内方案经运营方审核 · 接受统一品牌体系 · 3年起租 · 接入统一数字平台", { x: 0.8, y: 5.2, w: 11.5, h: 0.4, fontSize: 11, color: C.dg, fontFace: "Microsoft YaHei" });
s12.addText("招商渠道：上海旅游产业博览会 · 民宿/酒店行业协会推荐 · 小红书/抖音定向投放 · 松江区招商部门协同", { x: 0.8, y: 5.7, w: 11.5, h: 0.4, fontSize: 11, color: C.gray, fontFace: "Microsoft YaHei" });
snum(s12);

// === SLIDE 13: DIGITAL PLATFORM ===
let s13 = pptx.addSlide();
s13.background = { fill: C.white };
s13.addText("05  运营方案 — \"横山一点通\" 数字乡村平台", { x: 0.8, y: 0.4, w: 10, h: 0.7, fontSize: 26, bold: true, color: C.primary, fontFace: "Microsoft YaHei" });
s13.addShape("rect", { x: 0.8, y: 1, w: 1.5, h: 0.03, fill: { color: C.accent } });
const mods = [
  ["🗺️ 智能导览","AR导览 · 语音讲解 · 路线推荐 · 景点打卡"],
  ["📅 在线预订","民宿 · 餐饮 · 活动 · 停车一体化预订"],
  ["🛒 农产品电商","本地大米/竹笋在线销售 · 冷链配送"],
  ["🏛️ 村务服务","办事指南 · 村务公开 · 意见反馈"],
  ["📊 数据驾驶舱","客流热力图 · 消费画像 · 运营KPI看板"],
];
mods.forEach((m, i) => {
  let y = 1.5 + i * 1.1;
  s13.addShape("roundRect", { x: 0.8, y, w: 11.5, h: 0.9, fill: { color: i % 2 === 0 ? C.light : C.white }, rectRadius: 0.1 });
  s13.addText(m[0], { x: 1.2, y: y + 0.15, w: 4, h: 0.5, fontSize: 14, bold: true, color: C.primary, fontFace: "Microsoft YaHei" });
  s13.addText(m[1], { x: 5.5, y: y + 0.15, w: 6.5, h: 0.5, fontSize: 12, color: C.dg, fontFace: "Microsoft YaHei" });
});
snum(s13);

// === SLIDE 14: INVESTMENT FRAMEWORK ===
let s14 = pptx.addSlide();
s14.background = { fill: C.white };
s14.addText("06  投资框架（概算）", { x: 0.8, y: 0.4, w: 10, h: 0.7, fontSize: 26, bold: true, color: C.primary, fontFace: "Microsoft YaHei" });
s14.addShape("rect", { x: 0.8, y: 1, w: 1.5, h: 0.03, fill: { color: C.accent } });
let finRows = [
  [{text:"投资类别",options:{bold:true,fill:{color:C.primary},color:C.white}},{text:"责任主体",options:{bold:true,fill:{color:C.primary},color:C.white}},{text:"估算(万元)",options:{bold:true,fill:{color:C.primary},color:C.white}}],
  ["公共基础设施（道路/河道/管线/停车/景观）","政府/投资平台","1,500 - 3,000"],
  ["建筑外立面统一改造（4个核心点位）","政府/投资平台","500 - 1,000"],
  ["标识导牌系统 + 景观提升","政府/投资平台","200 - 400"],
  ["\"横山一点通\"数字平台开发","政府/投资平台","150 - 300"],
  ["品牌方室内设计装修（全部点位）","各品牌方","2,000 - 5,000"],
  ["运营方前期投入（团队/方案/招商/营销）","运营方","200 - 500"],
  [{text:"项目整体拉动投资（公共+品牌）",options:{bold:true,fill:{color:C.accent},color:C.dark}},{text:"",options:{fill:{color:C.accent}}},{text:"4,350 - 9,700",options:{bold:true,fill:{color:C.accent},color:C.dark}}],
];
s14.addTable(finRows, { x: 0.8, y: 1.3, w: 11.5, colW: [5, 2.5, 4], border: {type:"solid",pt:0.5,color:"CCCCCC"}, fontSize: 11, fontFace: "Microsoft YaHei", color: C.dg });
s14.addText("公共投资概算：2,350-4,700万元  |  品牌方投资概算：2,000-5,000万元  |  公共部分投资回报周期：5-8年（以社会效益为主）", { x: 0.8, y: 6.5, w: 11.5, h: 0.3, fontSize: 11, bold: true, color: C.primary, fontFace: "Microsoft YaHei", align: "center" });
snum(s14);

// === SLIDE 15: REVENUE MODEL ===
let s15 = pptx.addSlide();
s15.background = { fill: C.white };
s15.addText("06  运营方收入模型（稳定期预测）", { x: 0.8, y: 0.4, w: 10, h: 0.7, fontSize: 26, bold: true, color: C.primary, fontFace: "Microsoft YaHei" });
s15.addShape("rect", { x: 0.8, y: 1, w: 1.5, h: 0.03, fill: { color: C.accent } });
let revRows = [
  [{text:"收入来源",options:{bold:true,fill:{color:C.primary},color:C.white}},{text:"计费方式",options:{bold:true,fill:{color:C.primary},color:C.white}},{text:"年估算(万元)",options:{bold:true,fill:{color:C.primary},color:C.white}}],
  ["策划管理费","按公共投资额 3-5%（一次性/分期）","70 - 235"],
  ["运营管理费","品牌方月租金的 15-25%","60 - 150"],
  ["租金差额收益","我方整租→分租差价","50 - 120"],
  ["自营业态收益","咖啡厅 · 文创商店 · 活动门票","30 - 80"],
  ["数字平台佣金","预订/电商流水 5-10%","10 - 30"],
  ["政府绩效奖励","完成KPI后对标发放","面议"],
  [{text:"年总收入（稳定期）",options:{bold:true,fill:{color:C.accent},color:C.dark}},{text:"轻资产模式 · 2-3年回收前期投入",options:{fill:{color:C.accent}}},{text:"220 - 615+",options:{bold:true,fill:{color:C.accent},color:C.dark}}],
];
s15.addTable(revRows, { x: 0.8, y: 1.3, w: 11.5, colW: [3.5, 4.5, 3.5], border: {type:"solid",pt:0.5,color:"CCCCCC"}, fontSize: 12, fontFace: "Microsoft YaHei", color: C.dg });
snum(s15);

// === SLIDE 16: IMPLEMENTATION PHASES ===
let s16 = pptx.addSlide();
s16.background = { fill: C.white };
s16.addText("07  分阶段实施计划", { x: 0.8, y: 0.4, w: 10, h: 0.7, fontSize: 26, bold: true, color: C.primary, fontFace: "Microsoft YaHei" });
s16.addShape("rect", { x: 0.8, y: 1, w: 1.5, h: 0.03, fill: { color: C.accent } });
const phases = [
  { n: "第一阶段", t: "第1-3个月", tl: "筹备与顶层设计", it: "资产权属尽调 · 四方框架确认 · 控规对接微调 · 锚定品牌预热 · 数字平台启动开发" },
  { n: "第二阶段", t: "第4-8个月", tl: "基建与集中招商", it: "公共基建施工 · 外立面改造 · 导牌安装 · 招商达60% · 品牌方设计报审" },
  { n: "第三阶段", t: "第9-12个月", tl: "试运营与优化", it: "核心点位(大院/来云吧)率先试运营 · 品牌方装修进场 · 开幕活动·\"横山雅集\" · 招商达85%" },
  { n: "第四阶段", t: "第13-24个月", tl: "全面运营与品牌升级", it: "全面开业出租率90%+ · 四季活动日历 · 数据驱动运营 · 争取市/区级乡村振兴示范村评定" },
];
phases.forEach((p, i) => {
  let y = 1.5 + i * 1.35;
  s16.addShape("roundRect", { x: 0.8, y, w: 1.6, h: 1.1, fill: { color: C.primary }, rectRadius: 0.1 });
  s16.addText(p.n, { x: 0.8, y: y + 0.15, w: 1.6, h: 0.35, fontSize: 10, bold: true, color: C.white, fontFace: "Microsoft YaHei", align: "center" });
  s16.addText(p.t, { x: 0.8, y: y + 0.5, w: 1.6, h: 0.3, fontSize: 9, color: C.light, fontFace: "Microsoft YaHei", align: "center" });
  s16.addText(p.tl, { x: 2.7, y, w: 4, h: 0.4, fontSize: 14, bold: true, color: C.dark, fontFace: "Microsoft YaHei" });
  s16.addText(p.it, { x: 2.7, y: y + 0.4, w: 9.5, h: 0.6, fontSize: 11, color: C.dg, fontFace: "Microsoft YaHei" });
});
snum(s16);

// === SLIDE 17: POLICY MATRIX ===
let s17 = pptx.addSlide();
s17.background = { fill: C.white };
s17.addText("07  政策对接矩阵", { x: 0.8, y: 0.4, w: 10, h: 0.7, fontSize: 26, bold: true, color: C.primary, fontFace: "Microsoft YaHei" });
s17.addShape("rect", { x: 0.8, y: 1, w: 1.5, h: 0.03, fill: { color: C.accent } });
let polRows = [
  [{text:"项目模块",options:{bold:true,fill:{color:C.primary},color:C.white}},{text:"区级政策（核心条款）",options:{bold:true,fill:{color:C.primary},color:C.white}},{text:"市级政策",options:{bold:true,fill:{color:C.primary},color:C.white}}],
  ["精品民宿与文旅","《松江区旅游产业发展专项资金管理办法》","《上海市促进乡村民宿业发展的指导意见》"],
  ["特色农业与观光","《松江区都市现代农业发展专项补助资金》","《上海市都市现代农业发展专项补助资金》"],
  ["文创与乡愁经济","《松江区文化产业发展专项资金管理办法》","《上海市促进文化创意产业发展财政扶持资金》"],
  ["基础设施与数字乡村","《松江区促进乡村振兴专项资金》","《上海市乡村振兴专项资金》《上海市推进智慧乡村建设行动方案》"],
  ["集体经济壮大与运营","《松江区促进乡村振兴专项资金》","《关于进一步促进农村集体经济高质量发展的若干措施》"],
];
s17.addTable(polRows, { x: 0.5, y: 1.3, w: 12.3, colW: [3, 4.5, 4.8], border: {type:"solid",pt:0.5,color:"CCCCCC"}, fontSize: 10, fontFace: "Microsoft YaHei", color: C.dg });
s17.addText("🗓 区级申报窗口：上半年集中发布  |  市级文创资金：Q1末-Q2  |  市级都市农业：年中  |  乡村振兴示范村：提前一年盯", { x: 0.8, y: 5.5, w: 11.5, h: 0.4, fontSize: 11, color: C.dg, fontFace: "Microsoft YaHei" });
s17.addText("💡 申报策略：\"组合拳\"法——一个项目拆解为基建+民宿+农业+非遗多属性，从不同政策包匹配条款 → \"一个项目，多重支持\"", { x: 0.8, y: 6, w: 11.5, h: 0.4, fontSize: 11, italic: true, color: C.accent, fontFace: "Microsoft YaHei" });
s17.addText("申报联络：佘之春旅游发展有限公司 座机：57652901  联系人：刘雨洁 13817643468", { x: 0.8, y: 6.5, w: 11, h: 0.3, fontSize: 10, color: C.gray, fontFace: "Microsoft YaHei" });
snum(s17);

// === SLIDE 18: KEY SUCCESS + RISKS ===
let s18 = pptx.addSlide();
s18.background = { fill: C.white };
s18.addText("07  关键成功条件与风险提示", { x: 0.8, y: 0.4, w: 10, h: 0.7, fontSize: 26, bold: true, color: C.primary, fontFace: "Microsoft YaHei" });
s18.addText("✅ 关键成功条件", { x: 0.8, y: 1.5, w: 5.5, h: 0.5, fontSize: 18, bold: true, color: C.primary, fontFace: "Microsoft YaHei" });
["1. 政府明确授权运营方统筹地位和决策权限","2. 基础设施改造资金及时到位（分期实施，优先核心点位）","3. 首批锚定品牌成功引入（决定市场信心和后续招商引力）","4. 数字平台用户活跃度和预订转化率达标","5. 社区关系持续维护（村民优先用工+农产品采购+风貌共建）"].forEach((c, i) => {
  s18.addText(c, { x: 0.8, y: 2.1 + i * 0.55, w: 6, h: 0.4, fontSize: 11, color: C.dg, fontFace: "Microsoft YaHei" });
});
s18.addText("⚠ 风险与应对", { x: 7.2, y: 1.5, w: 5.5, h: 0.5, fontSize: 18, bold: true, color: "C0392B", fontFace: "Microsoft YaHei" });
const risks = ["同质化竞争 → 坚持IP差异化 + 四季活动日历","资金到位延迟 → 分期实施，优先核心点位","品牌经营不善 → 6个月退出机制 + 品牌储备库","政策变化 → 多元收入模型降低政策依赖","村民参与度低 → 优先用工 + 利益共享 + 持续沟通"];
risks.forEach((r, i) => {
  s18.addText(`⚠ ${r}`, { x: 7.2, y: 2.1 + i * 0.55, w: 5.8, h: 0.4, fontSize: 10.5, color: C.dg, fontFace: "Microsoft YaHei" });
});
snum(s18);

// === SLIDE 19: MARKET DATA APPENDIX ===
let s19 = pptx.addSlide();
s19.background = { fill: C.white };
s19.addText("08  附录：市场数据支撑", { x: 0.8, y: 0.4, w: 10, h: 0.7, fontSize: 26, bold: true, color: C.primary, fontFace: "Microsoft YaHei" });
s19.addShape("rect", { x: 0.8, y: 1, w: 1.5, h: 0.03, fill: { color: C.accent } });
let mktRows = [
  [{text:"数据指标",options:{bold:true,fill:{color:C.primary},color:C.white}},{text:"数值",options:{bold:true,fill:{color:C.primary},color:C.white}},{text:"来源",options:{bold:true,fill:{color:C.primary},color:C.white}}],
  ["佘山国家旅游度假区年游客量","1,570万+人次（2025）","上海市文旅局官方数据"],
  ["度假区内4A级以上景区","7个（欢乐谷/辰山/广富林/世茂洲际…）","佘山度假区管委会"],
  ["横山村距上海市中心（人民广场）","约39 km","百度地图"],
  ["横山村距虹桥交通枢纽","约28 km","百度地图"],
  ["上海市2025乡村振兴示范村数量","14个，\u201C沪派江南\u201D总体定位","上海市文旅局·乡村旅游三年行动方案(2025-2027)"],
  ["松江区乡村旅游年接待量","约800万人次（2025估算）","松江区文旅局"],
  ["上海乡村民宿平均入住率","约45-55%（2025行业数据）","上海市旅游行业协会"],
  ["高端民宿ADR（平均房价）","800-2,500元/间夜（上海郊区）","携程/行业报告"],
];
s19.addTable(mktRows, { x: 0.5, y: 1.3, w: 12.3, colW: [4.5, 3.8, 4], border: {type:"solid",pt:0.5,color:"CCCCCC"}, fontSize: 10, fontFace: "Microsoft YaHei", color: C.dg });
s19.addText("数据截至：2026年6月  |  数据来源：上海市文旅局、松江区文旅局、佘山国家旅游度假区管委会、百度地图、行业协会公开报告", { x: 0.8, y: 6.7, w: 11.5, h: 0.3, fontSize: 9, color: C.gray, fontFace: "Microsoft YaHei" });
snum(s19);

// === SLIDE 20: CLOSING ===
let s20 = pptx.addSlide();
s20.background = { fill: C.dark };
s20.addShape("rect", { x: 0, y: 3.2, w: "100%", h: 0.04, fill: { color: C.accent } });
s20.addText("感谢聆听", { x: 1, y: 1.5, w: "85%", h: 1.2, fontSize: 48, bold: true, color: C.white, fontFace: "Microsoft YaHei", align: "center" });
s20.addText("横山村乡村振兴项目 — 运营统筹方案", { x: 1, y: 3.5, w: "85%", h: 0.6, fontSize: 20, color: C.accent, fontFace: "Microsoft YaHei", align: "center" });
s20.addText("环山而居 · 沪派江南  |  四方共建 · 共谋发展", { x: 1, y: 4.5, w: "85%", h: 0.5, fontSize: 14, color: C.gray, fontFace: "Microsoft YaHei", align: "center" });
s20.addText("2026年6月  |  运营方", { x: 1, y: 5.8, w: "85%", h: 0.4, fontSize: 12, color: C.gray, fontFace: "Microsoft YaHei", align: "center" });
snum(s20);

// SAVE
const outPath = path.join(base, "docs", "横山村乡村振兴项目运营统筹方案.pptx");
await pptx.writeFile({ fileName: outPath });
console.log(`PPT generated: ${outPath} (${(fs.statSync(outPath).size / (1024*1024)).toFixed(1)} MB, ${sn} slides)`);


