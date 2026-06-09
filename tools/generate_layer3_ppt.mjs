import pptxgen from "pptxgenjs";
import fs from "fs";
import path from "path";

const pptx = new pptxgen();
pptx.defineLayout({ name: "WIDE", width: 13.333, height: 7.5 });
pptx.layout = "WIDE";
const base = "C:/Users/Lenovo/Documents/shanghai-hengshan-rural-revitalization";
const C = { dark: "1A1A2E", pri: "2D6A4F", acc: "E6A817", wh: "FFFFFF", lt: "F5F5F0", gy: "888888", dg: "444444", rd: "C0392B", bl: "2980B9" };
let sn = 0;
const snum = (s) => s.addText(`${++sn}`, { x: 12.5, y: 7.05, w: 0.6, h: 0.3, fontSize: 8, color: C.gy, align: "right", fontFace: "Microsoft YaHei" });

// ═══ SLIDE 1: COVER ═══
let s1 = pptx.addSlide(); s1.background = { fill: C.dark };
s1.addShape("rect", { x: 0, y: 3.3, w: "100%", h: 0.04, fill: { color: C.acc } });
s1.addText("上海市松江区横山村乡村振兴项目", { x: 1, y: 0.8, w: "85%", h: 1, fontSize: 38, bold: true, color: C.wh, fontFace: "Microsoft YaHei", align: "center" });
s1.addText("投资决策支撑报告", { x: 1, y: 2, w: "85%", h: 0.8, fontSize: 28, color: C.acc, fontFace: "Microsoft YaHei", align: "center" });
s1.addText("基础事实 → 定量分析 → 核心判断 → A/B/C三案 → 投资决策", { x: 1, y: 4.2, w: "85%", h: 0.4, fontSize: 13, color: C.gy, fontFace: "Microsoft YaHei", align: "center" });
s1.addText("2026年6月  |  运营方", { x: 1, y: 5.8, w: "85%", h: 0.4, fontSize: 12, color: C.gy, fontFace: "Microsoft YaHei", align: "center" });
snum(s1);

// ═══ SLIDE 2: METHODOLOGY ═══
let s2 = pptx.addSlide(); s2.background = { fill: C.wh };
s2.addText("报告方法论：从数据到决策的证据链", { x: 0.7, y: 0.4, w: 11, h: 0.7, fontSize: 24, bold: true, color: C.pri, fontFace: "Microsoft YaHei" });
s2.addShape("rect", { x: 0.7, y: 1.05, w: 1.5, h: 0.03, fill: { color: C.acc } });
const steps = ["基础事实","定量分析","核心判断","A/B/C三案","投资决策"];
steps.forEach((st, i) => {
  const x = 0.5 + i * 2.5;
  s2.addShape("roundRect", { x, y: 1.8, w: 2.2, h: 1, fill: { color: i < 3 ? C.pri : C.acc }, rectRadius: 0.1 });
  s2.addText(st, { x, y: (0.8 + i * 2.5) + 0.25, w: 2.2, h: 0.5, fontSize: 14, bold: true, color: C.wh, fontFace: "Microsoft YaHei", align: "center" });
  if (i < 4) s2.addText("→", { x: (0.5 + i * 2.5) + 2.2, y: 2, w: 0.3, h: 0.4, fontSize: 20, color: C.gy, align: "center" });
});
s2.addText("本报告所有结论均标注证据等级： ✅已核实数据  🔶推算/估算  ⚠待核实假设", { x: 0.7, y: 4.5, w: 11, h: 0.4, fontSize: 12, color: C.gy, fontFace: "Microsoft YaHei" });
snum(s2);

// ═══ SLIDE 3: LOCATION ═══
let s3 = pptx.addSlide(); s3.background = { fill: C.wh };
s3.addText("基础事实 — 区位与交通可达性", { x: 0.7, y: 0.4, w: 11, h: 0.7, fontSize: 22, bold: true, color: C.pri, fontFace: "Microsoft YaHei" });
s3.addShape("rect", { x: 0.7, y: 1.05, w: 1.5, h: 0.03, fill: { color: C.acc } });
let locData = [
  [{text:"目的地",options:{bold:true,fill:{color:C.pri},color:C.wh}},{text:"距离(km)",options:{bold:true,fill:{color:C.pri},color:C.wh}},{text:"驾车(min)",options:{bold:true,fill:{color:C.pri},color:C.wh}},{text:"证据",options:{bold:true,fill:{color:C.pri},color:C.wh}}],
  ["人民广场","~39","45-55","✅百度地图"],["虹桥枢纽","~28","30-40","✅百度地图"],["松江新城","~12","15-20","✅百度地图"],["佘山站(9号线)","~3.5","5-8","✅百度地图"],["G50高速入口","~8","12-15","🔶估算"],["G60高速入口","~12","18-22","🔶估算"],
];
s3.addTable(locData, { x: 0.5, y: 1.4, w: 5.5, colW: [2,1.5,1.2,0.8], border: {type:"solid",pt:0.5,color:"CCCCCC"}, fontSize: 10, fontFace: "Microsoft YaHei", color: C.dg });
// Right: isochrone
let isoData = [
  [{text:"圈层",options:{bold:true,fill:{color:C.bl},color:C.wh}},{text:"车程",options:{bold:true,fill:{color:C.bl},color:C.wh}},{text:"人口(估算)",options:{bold:true,fill:{color:C.bl},color:C.wh}}],
  ["核心圈","15min","~10万"],["紧密圈","30min","~60万"],["辐射圈","45min","~300万"],["拓展圈","60min","~800万"],["远程圈","90min","~2,500万"],
];
s3.addTable(isoData, { x: 6.5, y: 1.4, w: 6, colW: [1.5,1.5,3], border: {type:"solid",pt:0.5,color:"CCCCCC"}, fontSize: 10, fontFace: "Microsoft YaHei", color: C.dg });
s3.addText("核心优势：距市中心45分钟车程 = 近郊高频消费模型 vs 竞品莫干山/安吉2h+ = 低频远途", { x: 0.7, y: 5.8, w: 11.5, h: 0.3, fontSize: 11, bold: true, color: C.pri, fontFace: "Microsoft YaHei" });
snum(s3);

// ═══ SLIDE 4: POPULATION ═══
let s4 = pptx.addSlide(); s4.background = { fill: C.wh };
s4.addText("定量分析 — 人口与消费市场", { x: 0.7, y: 0.4, w: 11, h: 0.7, fontSize: 22, bold: true, color: C.pri, fontFace: "Microsoft YaHei" });
let popData = [
  [{text:"指标",options:{bold:true,fill:{color:C.pri},color:C.wh}},{text:"数据",options:{bold:true,fill:{color:C.pri},color:C.wh}},{text:"来源/等级",options:{bold:true,fill:{color:C.pri},color:C.wh}}],
  ["松江区常住人口(2025)","196.18万","✅松江统计局"],["外来常住人口","~116万(59%)","✅松江统计局"],["家庭户数","~65万户","🔶估算"],["大学城在校生","~15万","✅公开数据"],["中高收入家庭(>30万)","~25万户","⚠假设"],
];
s4.addTable(popData, { x: 0.5, y: 1.3, w: 6, colW: [3,1.5,1.5], border: {type:"solid",pt:0.5,color:"CCCCCC"}, fontSize: 10, fontFace: "Microsoft YaHei", color: C.dg });
let consData = [
  [{text:"消费场景",options:{bold:true,fill:{color:C.acc},color:C.dark}},{text:"人次/年(万)",options:{bold:true,fill:{color:C.acc},color:C.dark}},{text:"总规模(亿)",options:{bold:true,fill:{color:C.acc},color:C.dark}}],
  ["周末近郊休闲","600","12-30"],["亲子户外/研学","240","7-19"],["乡村餐饮体验","500","7.5-15"],["民宿/短期度假","120","7-30"],["企业团建/场地","80-170","4-51"],
];
s4.addTable(consData, { x: 7, y: 1.3, w: 5.5, colW: [2,1.5,2], border: {type:"solid",pt:0.5,color:"CCCCCC"}, fontSize: 10, fontFace: "Microsoft YaHei", color: C.dg });
s4.addText("佘山国家旅游度假区年客流1,570万+（2025）→ 横山村距度假区核心5km内，是天然分流节点", { x: 0.7, y: 6.2, w: 11, h: 0.3, fontSize: 12, bold: true, color: C.pri, fontFace: "Microsoft YaHei" });
snum(s4);

// ═══ SLIDE 5: COMPETITIVE ═══
let s5 = pptx.addSlide(); s5.background = { fill: C.wh };
s5.addText("定量分析 — 周边供给竞争格局", { x: 0.7, y: 0.4, w: 11, h: 0.7, fontSize: 22, bold: true, color: C.pri, fontFace: "Microsoft YaHei" });
let compData = [
  [{text:"项目",options:{bold:true,fill:{color:C.pri},color:C.wh}},{text:"年客流",options:{bold:true,fill:{color:C.pri},color:C.wh}},{text:"距离",options:{bold:true,fill:{color:C.pri},color:C.wh}}],
  ["欢乐谷","~300万","5km"],["辰山植物园","~200万","6km"],["广富林遗址","~180万","7km"],["佘山森林公园","~500万","4km"],["深坑酒店","~30万(alibaba)","4km"],
];
s5.addTable(compData, { x: 0.5, y: 1.3, w: 5.5, colW: [2.5,1.5,1.5], border: {type:"solid",pt:0.5,color:"CCCCCC"}, fontSize: 10, fontFace: "Microsoft YaHei", color: C.dg });
let gapData = [
  [{text:"品类",options:{bold:true,fill:{color:C.rd},color:C.wh}},{text:"供给密度",options:{bold:true,fill:{color:C.rd},color:C.wh}},{text:"机会评级",options:{bold:true,fill:{color:C.rd},color:C.wh}}],
  ["高端民宿集群","低","★★★★★"],["文化研学产品","低","★★★★★"],["骑行/徒步驿站","低","★★★★"],["乡村亲子体验","中","★★★★"],["精品咖啡馆","中","★★★"],
];
s5.addTable(gapData, { x: 6.5, y: 1.3, w: 5.8, colW: [2.5,1.5,1.8], border: {type:"solid",pt:0.5,color:"CCCCCC"}, fontSize: 10, fontFace: "Microsoft YaHei", color: C.dg });
s5.addText("关键发现：松江乡村板块高端供给稀缺 → 横山可占位\"松江高端乡村休闲首选目的地\"", { x: 0.7, y: 5.5, w: 11, h: 0.3, fontSize: 12, bold: true, color: C.pri, fontFace: "Microsoft YaHei" });
snum(s5);

// ═══ SLIDE 6: SITE ASSETS ═══
let s6 = pptx.addSlide(); s6.background = { fill: C.wh };
s6.addText("定量分析 — 场地资产盘点", { x: 0.7, y: 0.4, w: 11, h: 0.7, fontSize: 22, bold: true, color: C.pri, fontFace: "Microsoft YaHei" });
let astData = [
  [{text:"点位",options:{bold:true,fill:{color:C.pri},color:C.wh}},{text:"面积(㎡)",options:{bold:true,fill:{color:C.pri},color:C.wh}},{text:"栋/层",options:{bold:true,fill:{color:C.pri},color:C.wh}},{text:"适配功能(TOP3)",options:{bold:true,fill:{color:C.pri},color:C.wh}}],
  ["横山大院","1,563","4栋/四合院","游客中心 ★5, 市集 ★5, 民宿 ★4"],
  ["横山小学","590+2亩","1栋2层+大院","研学 ★5, 艺术 ★5, 展览 ★4"],
  ["来云吧","815","3栋","咖啡 ★5, 骑行驿站 ★5, 市集 ★4"],
  ["横云山居","7,881(含地下)","32栋","民宿 ★5, 团建 ★5, 亲子 ★5"],
  [{text:"合计",options:{bold:true,fill:{color:C.acc},color:C.dark}},{text:"~10,849",options:{bold:true,fill:{color:C.acc}}},{text:"40+",options:{bold:true,fill:{color:C.acc}}},{text:"覆盖民宿/餐饮/研学/文创/亲子/团建全业态",options:{bold:true,fill:{color:C.acc}}}],
];
s6.addTable(astData, { x: 0.5, y: 1.3, w: 12.3, colW: [2,2,2,6.3], border: {type:"solid",pt:0.5,color:"CCCCCC"}, fontSize: 10, fontFace: "Microsoft YaHei", color: C.dg });
s6.addText("可盘活总建面~10,849㎡ | 停车容量估83-127辆 | ⚠水电污消防待核实 | ⚠建筑结构安鉴待做", { x: 0.7, y: 5.8, w: 11, h: 0.3, fontSize: 11, color: C.gy, fontFace: "Microsoft YaHei" });
snum(s6);

// ═══ SLIDE 7: COMPETITOR CASES ═══
let s7 = pptx.addSlide(); s7.background = { fill: C.wh };
s7.addText("竞品对标 — 关键借鉴与差异机会", { x: 0.7, y: 0.4, w: 11, h: 0.7, fontSize: 22, bold: true, color: C.pri, fontFace: "Microsoft YaHei" });
let benchData = [
  [{text:"案例",options:{bold:true,fill:{color:C.pri},color:C.wh}},{text:"价格/间夜",options:{bold:true,fill:{color:C.pri},color:C.wh}},{text:"核心借鉴",options:{bold:true,fill:{color:C.pri},color:C.wh}},{text:"不适用横山",options:{bold:true,fill:{color:C.pri},color:C.wh}}],
  ["八十八亩田(松江)","¥800-1,500","稻田IP化+小红书流量","体量小，依赖个人IP"],
  ["计家墩(昆山)","¥1,200-3,500","新村民机制+社群运营","整村搬迁模式不可行"],
  ["莫干山集群","¥800-5,000+","洋家乐高端定位","距上海2h+，已成红海"],
  ["青山村(杭州)","教育为主","设计+教育双引擎","依赖独特IP，难复制"],
  ["安吉桃花源","度假大盘","国际品牌+复合型","占地超大，投入量级不同"],
];
s7.addTable(benchData, { x: 0.5, y: 1.2, w: 12.3, colW: [2.5,2,3.5,4.3], border: {type:"solid",pt:0.5,color:"CCCCCC"}, fontSize: 10, fontFace: "Microsoft YaHei", color: C.dg });
s7.addText("横山差异化：近郊45min(非远途2h+) × 高端供给稀缺 × 黄公望IP唯一性 = \"松江高端乡村休闲首选\"", { x: 0.7, y: 6.2, w: 11, h: 0.3, fontSize: 11, bold: true, color: C.pri, fontFace: "Microsoft YaHei" });
snum(s7);

// ═══ SLIDE 8: CORE JUDGMENTS ═══
let s8 = pptx.addSlide(); s8.background = { fill: C.wh };
s8.addText("核心判断 — 从数据推导出的三个关键结论", { x: 0.7, y: 0.4, w: 11, h: 0.7, fontSize: 22, bold: true, color: C.pri, fontFace: "Microsoft YaHei" });
const judgments = [
  { t: "判断一：走\"近郊高频\"模型，不走\"远郊低频\"路线", d: "45min车程覆盖800万人口，周末高频消费场景(餐饮+咖啡+骑行+亲子)是核心引擎，住宿是提升客单价的手段而非唯一收入来源", c: C.pri },
  { t: "判断二：以\"文化IP+高端供给\"避免同质化竞争", d: "黄公望IP + 上海唯一环山村落 + 松江高端乡村供给稀缺 → 定位\"松江高端乡村休闲首选目的地\"", c: C.bl },
  { t: "判断三：分期推进，先验证高频消费模型再升级住宿产品", d: "基建条件未核实(水电污消防)+品牌投资意向未验证 → 先用最小投资跑通模型，成功后再争取更大政策支持", c: C.acc },
];
judgments.forEach((j, i) => {
  const y = 1.4 + i * 1.9;
  s8.addShape("roundRect", { x: 0.7, y, w: 11.8, h: 1.6, fill: { color: "#FAFAFA" }, rectRadius: 0.1, line: { color: j.c, width: 2 } });
  s8.addText(j.t, { x: 1.2, y: y + 0.15, w: 10.5, h: 0.5, fontSize: 16, bold: true, color: j.c, fontFace: "Microsoft YaHei" });
  s8.addText(j.d, { x: 1.2, y: y + 0.7, w: 10.5, h: 0.7, fontSize: 12, color: C.dg, fontFace: "Microsoft YaHei" });
});
snum(s8);

// ═══ SLIDE 9: ABC COMPARISON ═══
let s9 = pptx.addSlide(); s9.background = { fill: C.wh };
s9.addText("A/B/C 三案对比 — 三种战略路径", { x: 0.7, y: 0.4, w: 11, h: 0.7, fontSize: 22, bold: true, color: C.pri, fontFace: "Microsoft YaHei" });
let abcData = [
  [{text:"维度",options:{bold:true,fill:{color:C.dark},color:C.wh}},{text:"A案:存量提效",options:{bold:true,fill:{color:C.bl},color:C.wh}},{text:"B案:系统重构",options:{bold:true,fill:{color:C.pri},color:C.wh}},{text:"C案:分期进击 ★推荐",options:{bold:true,fill:{color:C.acc},color:C.dark}}],
  ["战略核心","接受框架，做存量提效","重新定义，重构产业系统","B案方向，分期验证推进"],
  ["建筑改造面积","~4,500㎡","~8,500㎡","一期3,500 → 二期5,000㎡"],
  ["建设周期","12个月","24-30个月","一期12 + 二期24个月"],
  ["政府投入","1,500-2,500万","3,500-5,000万","分期3,500-5,000万"],
  ["社会资本拉动","500-1,500万","3,000-6,000万","分期2,500-6,000万"],
  ["年运营收入","300-600万","1,200-2,000万","300 → 1,500万"],
  ["年集体收益","50-100万","150-300万","50 → 200万"],
  ["投资回收期(公共)","5-7年","8-12年","6-10年"],
  ["核心风险","差异化不足","投资回收压力大"," 二期资金不确定性"],
];
s9.addTable(abcData, { x: 0.3, y: 1.1, w: 12.7, colW: [2.5,2.8,2.8,4.6], border: {type:"solid",pt:0.5,color:"CCCCCC"}, fontSize: 9, fontFace: "Microsoft YaHei", color: C.dg });
snum(s9);

// ═══ SLIDE 10: PLAN C MILESTONES ═══
let s10 = pptx.addSlide(); s10.background = { fill: C.wh };
s10.addText("推荐方案C — 分期里程碑", { x: 0.7, y: 0.4, w: 11, h: 0.7, fontSize: 22, bold: true, color: C.acc, fontFace: "Microsoft YaHei" });
const ms = [
  { t: "1-3月", tl: "资产核实+基建检测", d: "完成10项待核实清单→确定改造边界", },
  { t: "4-6月", tl: "一期公共工程", d: "横山大院+来云吧外立面改造，标识系统安装" },
  { t: "7-9月", tl: "一期招商完成", d: "出租率60%+，引入本地餐饮/咖啡/小民宿" },
  { t: "10-12月", tl: "一期试运营", d: "月客流>1万，收集运营数据→验证消费模型" },
  { t: "13-18月", tl: "运营复盘→启动二期", d: "关键指标达标(入住率/客单价/复购率)→政府审批二期" },
  { t: "19-30月", tl: "全面运营", d: "横山小学+横云山居开业，年客流>50万，争取示范村" },
];
ms.forEach((m, i) => {
  const y = 1.3 + i * 0.95;
  s10.addShape("roundRect", { x: 0.5, y, w: 1.6, h: 0.75, fill: { color: C.acc }, rectRadius: 0.08 });
  s10.addText(m.t, { x: 0.5, y: y + 0.15, w: 1.6, h: 0.45, fontSize: 10, bold: true, color: C.dark, fontFace: "Microsoft YaHei", align: "center" });
  s10.addText(m.tl, { x: 2.3, y, w: 4, h: 0.35, fontSize: 13, bold: true, color: C.dark, fontFace: "Microsoft YaHei" });
  s10.addText(m.d, { x: 2.3, y: y + 0.35, w: 10, h: 0.35, fontSize: 10, color: C.dg, fontFace: "Microsoft YaHei" });
});
snum(s10);

// ═══ SLIDE 11: DECISION ITEMS ═══
let s11 = pptx.addSlide(); s11.background = { fill: C.wh };
s11.addText("提请政府决策事项", { x: 0.7, y: 0.4, w: 11, h: 0.7, fontSize: 22, bold: true, color: C.pri, fontFace: "Microsoft YaHei" });
const decs = [
  "1. 确认\"四方结构\"合作框架（政府/平台/运营方/品牌方），明确运营方统筹地位",
  "2. 授权启动资产权属与基建条件实地核查（10项待核实清单）",
  "3. 确认C案分期推进策略，批准一期公共工程预算框架（1,500-2,000万）",
  "4. 协调规资局明确控规边界与生态保护红线",
  "5. 指定对口部门与联络人，建立月度联席会机制",
  "6. 授权运营方启动一期招商预热（接触头部民宿品牌+本地特色商户）",
];
decs.forEach((d, i) => {
  s11.addText(d, { x: 1, y: 1.5 + i * 0.85, w: 11, h: 0.65, fontSize: 14, color: C.dark, fontFace: "Microsoft YaHei" });
});
snum(s11);

// ═══ SLIDE 12: CLOSING ═══
let s12 = pptx.addSlide(); s12.background = { fill: C.dark };
s12.addShape("rect", { x: 0, y: 3.2, w: "100%", h: 0.04, fill: { color: C.acc } });
s12.addText("感谢聆听", { x: 1, y: 1.2, w: "85%", h: 1.2, fontSize: 48, bold: true, color: C.wh, fontFace: "Microsoft YaHei", align: "center" });
s12.addText("横山村乡村振兴项目 — 投资决策支撑报告", { x: 1, y: 3.6, w: "85%", h: 0.6, fontSize: 18, color: C.acc, fontFace: "Microsoft YaHei", align: "center" });
s12.addText("环山而居 · 沪派江南  |  四方共建 · 共谋发展", { x: 1, y: 4.8, w: "85%", h: 0.5, fontSize: 14, color: C.gy, fontFace: "Microsoft YaHei", align: "center" });
snum(s12);

// SAVE
const out = `${base}/docs/layer3-政府汇报/横山村乡村振兴项目-投资决策支撑报告.pptx`;
await pptx.writeFile({ fileName: out });
console.log(`Layer3 PPT: ${out} (${(fs.statSync(out).size/1024).toFixed(0)} KB, ${sn} slides)`);

