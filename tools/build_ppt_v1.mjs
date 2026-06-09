import pptxgen from "pptxgenjs";
import fs from "fs";

const pptx = new pptxgen();
pptx.defineLayout({ name: "WIDE", width: 13.333, height: 7.5 });
pptx.layout = "WIDE";

const C = { bg: "0F1F2F", primary: "1A3A5C", accent: "C9A96E", white: "FFFFFF", light: "E8EDF2", gray: "8899AA", dg: "5A6A7A", h1: "F5F8FC", body: "A8B8C8" };
let sn = 0;
const snum = (s) => s.addText(String(++sn), { x: 12.5, y: 7.05, w: 0.6, h: 0.3, fontSize: 8, color: C.gray, align: "right", fontFace: "Microsoft YaHei" });

// SLIDE 1: COVER
let s1 = pptx.addSlide(); s1.background = { fill: C.bg };
s1.addShape("rect", { x: 1, y: 3.4, w: 4, h: 0.04, fill: { color: C.accent } });
s1.addText("横山不是又一个乡村民宿项目", { x: 1, y: 1.2, w: 10, h: 1.4, fontSize: 42, bold: true, color: C.h1, fontFace: "Microsoft YaHei" });
s1.addText("上海市松江区横山村乡村振兴项目 · 投资决策支撑报告", { x: 1, y: 2.8, w: 10, h: 0.5, fontSize: 18, color: C.body, fontFace: "Microsoft YaHei" });
s1.addText("2026年6月  |  运营方", { x: 1, y: 5.8, w: 10, h: 0.3, fontSize: 12, color: C.gray, fontFace: "Microsoft YaHei" });
snum(s1);

// SLIDE 2: WHY HENGSHAN
let s2 = pptx.addSlide(); s2.background = { fill: C.primary };
s2.addText("01 为什么是横山", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.accent, fontFace: "Microsoft YaHei" });
s2.addShape("rect", { x: 0.8, y: 0.85, w: 1.5, h: 0.03, fill: { color: C.accent } });
s2.addText("上海唯一45分钟可达的近郊高频休闲目的地", { x: 0.8, y: 1.2, w: 11, h: 1, fontSize: 28, bold: true, color: C.h1, fontFace: "Microsoft YaHei" });
s2.addText("莫干山2h+ · 安吉2.5h+ · 横山45min", { x: 0.8, y: 2.5, w: 8, h: 0.4, fontSize: 16, color: C.accent, fontFace: "Microsoft YaHei" });
s2.addText("六个不可复制的条件：环山村落(上海唯一)·黄公望隐居地(700年)·松江大米GI(2014)·佘山1,570万客流(5km)·近郊45min·万米可盘活空间", { x: 0.8, y: 3.5, w: 11.5, h: 0.6, fontSize: 13, color: C.body, fontFace: "Microsoft YaHei" });
s2.addText("黄公望50岁辞官隐居横山，画出《富春山居图》。700年后同一片田、同一品种的米。", { x: 0.8, y: 5.5, w: 11, h: 0.4, fontSize: 12, italic: true, color: C.gray, fontFace: "Microsoft YaHei" });
snum(s2);

// SLIDE 3: SIX CONDITIONS
let s3 = pptx.addSlide(); s3.background = { fill: C.primary };
s3.addText("02 六个不可复制的条件", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.accent, fontFace: "Microsoft YaHei" });
const conds = ["环山型村落 · 上海唯一","黄公望隐居地 · 700年","松江大米GI · 法律壁垒","佘山1,570万客流 · 5km","近郊45分钟 · 高频可达","万米可盘活 · 村落级系统"];
conds.forEach((c, i) => {
  const col = i % 3, row = Math.floor(i / 3);
  s3.addShape("roundRect", { x: 0.5 + col * 4.1, y: 1.5 + row * 2.6, w: 3.8, h: 2.2, fill: { color: i % 2 === 0 ? "162D42" : "1A3550" }, rectRadius: 0.1, line: { color: C.accent, width: 1 } });
  s3.addText(String(i + 1), { x: 0.8 + col * 4.1, y: 1.6 + row * 2.6, w: 0.8, h: 0.6, fontSize: 24, bold: true, color: C.accent, fontFace: "Microsoft YaHei" });
  s3.addText(c, { x: 1.6 + col * 4.1, y: 1.7 + row * 2.6, w: 2.5, h: 1.8, fontSize: 14, color: C.h1, fontFace: "Microsoft YaHei" });
});
snum(s3);

// SLIDE 4: MARKET
let s4 = pptx.addSlide(); s4.background = { fill: C.primary };
s4.addText("03 市场在说话", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.accent, fontFace: "Microsoft YaHei" });
s4.addShape("rect", { x: 0.8, y: 0.85, w: 1.5, h: 0.03, fill: { color: C.accent } });
s4.addText("全球乡村旅游市场约$1,000亿 · 中国约1.5万亿 · 松江乡村高端供给极度稀缺", { x: 0.8, y: 1.3, w: 11, h: 1, fontSize: 22, bold: true, color: C.h1, fontFace: "Microsoft YaHei" });
s4.addText("佘山周边120家酒店民宿 · 缺少集群式高端乡村民宿", { x: 0.8, y: 2.8, w: 8, h: 0.4, fontSize: 14, color: C.accent, fontFace: "Microsoft YaHei" });
let mkData = [["品类","供给密度","机会"],["高端民宿集群","低","★★★★★"],["文化研学","极低","★★★★★"],["骑行配套","极低","★★★★★"],["乡村夜经济","极低","★★★★★"],["森林疗愈","极低","★★★★★"]];
let mkHeader = mkData[0].map(h=>({text:h,options:{bold:true,fill:{color:C.accent},color:C.bg,fontFace:"Microsoft YaHei",fontSize:10}}));
let mkBody = mkData.slice(1).map(r=>r.map(c=>({text:c,options:{fontSize:10,color:C.h1,fontFace:"Microsoft YaHei"}})));
s4.addTable([mkHeader, ...mkBody], { x: 0.5, y: 3.5, w: 12, border: {type:"solid",pt:0.5,color:C.dg}, colW: [4,3,5] });
snum(s4);

// SLIDE 5: CUSTOMER
let s5 = pptx.addSlide(); s5.background = { fill: C.primary };
s5.addText("04 谁在买单", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.accent, fontFace: "Microsoft YaHei" });
s5.addText("四类核心客群 · 年客流20-40万人次", { x: 0.8, y: 1.2, w: 10, h: 0.8, fontSize: 24, bold: true, color: C.h1, fontFace: "Microsoft YaHei" });
const pers = [{label:"亲子家庭 40%",desc:"30-45岁·有车·年4-8次",x:0.5,y:2.5},{label:"中产个人 30%",desc:"28-45岁·可下班来·年6-12次",x:6.8,y:2.5},{label:"大学城师生 15%",desc:"18-25岁·拼车公交·年2-4次",x:0.5,y:4.5},{label:"企业团建 15%",desc:"20-100人·工作日",x:6.8,y:4.5}];
pers.forEach(p => {
  s5.addShape("roundRect", { x: p.x, y: p.y, w: 5.8, h: 1.6, fill: { color: "162D42" }, rectRadius: 0.08 });
  s5.addText(p.label, { x: p.x + 0.3, y: p.y + 0.2, w: 5.2, h: 0.5, fontSize: 16, bold: true, color: C.accent, fontFace: "Microsoft YaHei" });
  s5.addText(p.desc, { x: p.x + 0.3, y: p.y + 0.8, w: 5.2, h: 0.5, fontSize: 12, color: C.body, fontFace: "Microsoft YaHei" });
});
snum(s5);

// SLIDE 6: DIFFERENTIATION
let s6 = pptx.addSlide(); s6.background = { fill: C.primary };
s6.addText("05 凭什么和别人不一样", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.accent, fontFace: "Microsoft YaHei" });
s6.addShape("rect", { x: 0.8, y: 0.85, w: 1.5, h: 0.03, fill: { color: C.accent } });
s6.addText("别人只有一个剧场——横山整个村就是剧场。4个点位环山分布是天然的场景切换。24个节气是天然的时间分层。排列组合本身就是最深的差异化。", { x: 1.5, y: 1.8, w: 10, h: 2, fontSize: 20, italic: true, color: C.accent, fontFace: "Microsoft YaHei", align: "center" });
const compRows = [
  [{text:"",options:{fontSize:10,fill:{color:"162D42"}}},{text:"莫干山",options:{bold:true,fontSize:11,fill:{color:"162D42"},color:C.accent}},{text:"八十八亩田",options:{bold:true,fontSize:11,fill:{color:"162D42"},color:C.accent}},{text:"横山",options:{bold:true,fontSize:11,fill:{color:C.accent},color:C.bg}}],
  [{text:"车程",options:{fontSize:10,fill:{color:"162D42"},color:C.h1}},{text:"2h+",options:{fontSize:10,color:C.h1}},{text:"50min",options:{fontSize:10,color:C.h1}},{text:"45min",options:{fontSize:10,color:C.h1}}],
  [{text:"文化IP",options:{fontSize:10,fill:{color:"162D42"},color:C.h1}},{text:"无",options:{fontSize:10,color:C.h1}},{text:"无",options:{fontSize:10,color:C.h1}},{text:"黄公望·700年",options:{fontSize:10,color:C.h1}}],
  [{text:"地理标志",options:{fontSize:10,fill:{color:"162D42"},color:C.h1}},{text:"无",options:{fontSize:10,color:C.h1}},{text:"无",options:{fontSize:10,color:C.h1}},{text:"松江大米GI",options:{fontSize:10,color:C.h1}}],
];
s6.addTable(compRows, { x: 0.8, y: 4, w: 11.5, colW: [2.5,3,3,3], border: {type:"solid",pt:0.5,color:C.dg} });
snum(s6);

// SLIDE 7: PRODUCT LINES
let s7 = pptx.addSlide(); s7.background = { fill: C.primary };
s7.addText("06 五大差异化业态线", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.accent, fontFace: "Microsoft YaHei" });
const lines = [
  {t:"米文旅线",s:"松江大米GI·700年从种子到餐桌·插秧→收割→碾米→品鉴→米酒封坛→稻田晚宴"},
  {t:"森林疗愈线",s:"唐代疗愈四法·日本森林浴·皮质醇降低12-16%"},
  {t:"黄公望文化线",s:"50岁辞官隐居·AI导览·画疗·艺术驻留"},
  {t:"环山户外线",s:"8km环山闭环·骑行驿站·夜骑LED·上海50-80万骑行人口"},
  {t:"乡村夜经济线",s:"夜骑·暗夜星空(中国首个)·皮影戏·围炉夜话·深夜食堂"},
];
lines.forEach((l, i) => {
  s7.addShape("roundRect", { x: 0.5, y: 1.2 + i * 1.2, w: 12.3, h: 1, fill: { color: i % 2 === 0 ? "162D42" : "1A3550" }, rectRadius: 0.06 });
  s7.addText(l.t, { x: 0.8, y: 1.25 + i * 1.2, w: 2.5, h: 0.5, fontSize: 13, bold: true, color: C.accent, fontFace: "Microsoft YaHei" });
  s7.addText(l.s, { x: 3.5, y: 1.25 + i * 1.2, w: 9, h: 0.9, fontSize: 12, color: C.body, fontFace: "Microsoft YaHei" });
});
snum(s7);

// SLIDE 8: A+B+C
let s8 = pptx.addSlide(); s8.background = { fill: C.primary };
s8.addText("07 A+B+C投资组合", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.accent, fontFace: "Microsoft YaHei" });
s8.addText("最优启动包：A+C = ¥93-103万 = 15条差异化业态", { x: 0.8, y: 1.2, w: 11, h: 0.8, fontSize: 22, bold: true, color: C.h1, fontFace: "Microsoft YaHei" });
const abcRows = [
  [{text:"级别",options:{bold:true,fill:{color:C.accent},color:C.bg}},{text:"定义",options:{bold:true,fill:{color:C.accent},color:C.bg}},{text:"条数",options:{bold:true,fill:{color:C.accent},color:C.bg}},{text:"投资",options:{bold:true,fill:{color:C.accent},color:C.bg}}],
  [{text:"A:基础盘",options:{fontSize:10,color:C.h1}},{text:"立即启动·3个月上线",options:{fontSize:10,color:C.h1}},{text:"8条",options:{fontSize:10,color:C.h1}},{text:"43-47万",options:{fontSize:10,color:C.h1}}],
  [{text:"C:最优组合",options:{fontSize:10,color:C.h1}},{text:"边际成本趋零",options:{fontSize:10,color:C.h1}},{text:"7条",options:{fontSize:10,color:C.h1}},{text:"49-56万",options:{fontSize:10,color:C.h1}}],
  [{text:"A+C启动包",options:{bold:true,fontSize:10,fill:{color:"162D42"},color:C.accent}},{text:"15条业态·完整体验",options:{fontSize:10,fill:{color:"162D42"},color:C.h1}},{text:"15条",options:{fontSize:10,fill:{color:"162D42"},color:C.h1}},{text:"93-103万",options:{bold:true,fontSize:10,fill:{color:"162D42"},color:C.accent}}],
  [{text:"B:完整版",options:{fontSize:10,color:C.h1}},{text:"能级跃升·A+C数据驱动",options:{fontSize:10,color:C.h1}},{text:"7条",options:{fontSize:10,color:C.h1}},{text:"155-270万",options:{fontSize:10,color:C.h1}}],
];
s8.addTable(abcRows, { x: 0.8, y: 2.5, w: 11.5, colW: [3,4,1.5,3], border: {type:"solid",pt:0.5,color:C.dg} });
snum(s8);

// SLIDE 9: IMPLEMENTATION
let s9 = pptx.addSlide(); s9.background = { fill: C.primary };
s9.addText("08 分阶段实施", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.accent, fontFace: "Microsoft YaHei" });
const phases = [
  {t:"第1-3个月",tl:"筹备与顶层设计",d:"资产权属摸底·四方框架确认·控规边界·锚定品牌预热"},
  {t:"第4-8个月",tl:"基建与招商",d:"公共基建·外立面改造·招商60%"},
  {t:"第9-12个月",tl:"试运营",d:"核心点位开业·开幕活动·招商85%"},
  {t:"第13-24个月",tl:"全面运营",d:"全点位开业·年客流>20万·争取示范村评定"},
];
phases.forEach((p, i) => {
  const y = 1.5 + i * 1.4;
  s9.addShape("roundRect", { x: 0.5, y, w: 2.5, h: 1.1, fill: { color: "162D42" }, rectRadius: 0.1 });
  s9.addText(p.t, { x: 0.5, y: y + 0.1, w: 2.5, h: 0.4, fontSize: 11, bold: true, color: C.accent, fontFace: "Microsoft YaHei", align: "center" });
  s9.addText(p.tl, { x: 0.5, y: y + 0.5, w: 2.5, h: 0.4, fontSize: 10, color: C.h1, fontFace: "Microsoft YaHei", align: "center" });
  s9.addText(p.d, { x: 3.3, y: y + 0.2, w: 9, h: 0.7, fontSize: 12, color: C.body, fontFace: "Microsoft YaHei" });
});
snum(s9);

// SLIDE 10: DECISION
let s10 = pptx.addSlide(); s10.background = { fill: C.primary };
s10.addText("09 提请政府决策", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.accent, fontFace: "Microsoft YaHei" });
const decs = ["确认四方结构合作框架，明确运营方统筹地位","授权启动资产权属与基建条件实地核查","确认A+C组合方案（93-103万·15条业态）","协调规资局明确控规边界与生态保护红线","指定对口部门与联络人，建立月度联席会机制","授权运营方启动一期招商预热"];
decs.forEach((d, i) => { s10.addText((i+1) + ". " + d, { x: 1.2, y: 1.3 + i * 0.75, w: 10, h: 0.5, fontSize: 15, color: C.h1, fontFace: "Microsoft YaHei" }); });
snum(s10);

// SLIDE 11: CLOSING
let s11 = pptx.addSlide(); s11.background = { fill: C.bg };
s11.addShape("rect", { x: 1, y: 3.2, w: 4, h: 0.04, fill: { color: C.accent } });
s11.addText("感谢聆听", { x: 1, y: 1.5, w: 11, h: 1.2, fontSize: 44, bold: true, color: C.h1, fontFace: "Microsoft YaHei", align: "center" });
s11.addText("环山而居 · 沪派江南  |  四方共建 · 共谋发展", { x: 1, y: 3.8, w: 11, h: 0.5, fontSize: 16, color: C.accent, fontFace: "Microsoft YaHei", align: "center" });
snum(s11);

const out = "output/Hengshan_PPT_v1.pptx";
await pptx.writeFile({ fileName: out });
const size = fs.statSync(out).size;
console.log("PPT v1: " + out + " (" + Math.round(size/1024) + " KB, " + sn + " slides)");
