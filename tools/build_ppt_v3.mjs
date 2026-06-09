import pptxgen from "pptxgenjs";
import fs from "fs";

const pptx = new pptxgen();
pptx.defineLayout({ name: "WIDE", width: 13.333, height: 7.5 });
pptx.layout = "WIDE";

const C = { bg: "0F1F2F", pri: "1A3A5C", acc: "C9A96E", wh: "FFFFFF", lt: "E8EDF2", gy: "8899AA", dg: "5A6A7A", h1: "F5F8FC", bd: "A8B8C8", red: "E74C3C", grn: "27AE60" };
let sn = 0;
const snum = (s) => s.addText(String(++sn), { x: 12.5, y: 7.05, w: 0.6, h: 0.3, fontSize: 8, color: C.gy, align: "right", fontFace: "Microsoft YaHei" });

// ═══ S1: COVER ═══
let s1 = pptx.addSlide(); s1.background = { fill: C.bg };
s1.addShape("rect", { x: 1, y: 3.2, w: 4, h: 0.05, fill: { color: C.acc } });
s1.addText("横山不是又一个乡村民宿项目", { x: 1, y: 1.2, w: 10, h: 1.4, fontSize: 44, bold: true, color: C.h1, fontFace: "Microsoft YaHei" });
s1.addText("上海市松江区横山村乡村振兴项目 · 投资决策支撑报告", { x: 1, y: 2.8, w: 10, h: 0.5, fontSize: 18, color: C.bd, fontFace: "Microsoft YaHei" });
s1.addText("2026年6月  |  运营方", { x: 1, y: 5.8, w: 10, h: 0.3, fontSize: 12, color: C.gy, fontFace: "Microsoft YaHei" });
snum(s1);

// ═══ S2: WHY ═══
let s2 = pptx.addSlide(); s2.background = { fill: C.pri };
s2.addText("01  为什么是横山", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
s2.addShape("rect", { x: 0.8, y: 0.85, w: 1.5, h: 0.03, fill: { color: C.acc } });
s2.addText("上海唯一45分钟可达的近郊高频休闲目的地", { x: 0.8, y: 1.5, w: 11, h: 1, fontSize: 30, bold: true, color: C.h1, fontFace: "Microsoft YaHei" });
s2.addText("莫干山2h+ · 安吉2.5h+ · 横山45min  |  车程缩短2.7倍  |  频次差5-10倍", { x: 0.8, y: 2.8, w: 10, h: 0.4, fontSize: 16, color: C.acc, fontFace: "Microsoft YaHei" });
s2.addText("六个不可复制条件：环山村落(上海唯一)·黄公望隐居地(700年·五条证据)·松江大米GI(2014·法律保护)·佘山1,570万客流(5km·必经之路)·近郊45min·万米可盘活空间(10,849㎡·40+栋·40亩)", { x: 0.8, y: 3.8, w: 11.5, h: 0.8, fontSize: 12, color: C.bd, fontFace: "Microsoft YaHei" });
s2.addText("黄公望50岁辞官隐居横山，画出《富春山居图》。700年后——同一片田、同一品种的米。你来横山，吃的就是他吃过的米，看的就是他画过的山。", { x: 0.8, y: 5.5, w: 11, h: 0.5, fontSize: 12, italic: true, color: C.gy, fontFace: "Microsoft YaHei" });
snum(s2);

// ═══ S3: SIX CONDITIONS ═══
let s3 = pptx.addSlide(); s3.background = { fill: C.pri };
s3.addText("02  六个不可复制的条件", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
const conds = [{n:"1",t:"环山型村落",s:"上海唯一"}, {n:"2",t:"黄公望隐居地",s:"700年"}, {n:"3",t:"松江大米GI",s:"法律保障"}, {n:"4",t:"佘山客流",s:"1,570万/年"}, {n:"5",t:"近郊45min",s:"高频可达"}, {n:"6",t:"万米空间",s:"10,849㎡"}];
conds.forEach((c, i) => {
  const col = i % 3, row = Math.floor(i / 3);
  s3.addShape("roundRect", { x: 0.5 + col * 4.1, y: 1.5 + row * 2.6, w: 3.8, h: 2.2, fill: { color: "162D42" }, rectRadius: 0.1, line: { color: C.acc, width: 1.5 } });
  s3.addText(c.n, { x: 0.8 + col * 4.1, y: 1.7 + row * 2.6, w: 0.7, h: 0.6, fontSize: 28, bold: true, color: C.acc, fontFace: "Microsoft YaHei" });
  s3.addText(c.t, { x: 0.8 + col * 4.1, y: 2.3 + row * 2.6, w: 3.2, h: 0.5, fontSize: 16, bold: true, color: C.h1, fontFace: "Microsoft YaHei" });
  s3.addText(c.s, { x: 0.8 + col * 4.1, y: 2.8 + row * 2.6, w: 3.2, h: 0.5, fontSize: 12, color: C.acc, fontFace: "Microsoft YaHei" });
});
snum(s3);

// ═══ S4: MARKET ═══
let s4 = pptx.addSlide(); s4.background = { fill: C.pri };
s4.addText("03  市场在说话", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
s4.addText("全球乡村旅游约$1,000亿 · 中国约1.5万亿 · 上海年接待约2,000万人次", { x: 0.8, y: 1.2, w: 11, h: 0.7, fontSize: 20, bold: true, color: C.h1, fontFace: "Microsoft YaHei" });
s4.addText("上海2025年发布《乡村旅游三年行动方案》· 沪派江南定位 · 14个示范村", { x: 0.8, y: 2, w: 10, h: 0.3, fontSize: 12, color: C.acc, fontFace: "Microsoft YaHei" });
const mkRows = [
  [{text:"品类",options:{bold:true,fill:{color:C.acc},color:C.bg,fontSize:10}},{text:"供给密度",options:{bold:true,fill:{color:C.acc},color:C.bg,fontSize:10}},{text:"品质",options:{bold:true,fill:{color:C.acc},color:C.bg,fontSize:10}},{text:"机会",options:{bold:true,fill:{color:C.acc},color:C.bg,fontSize:10}},{text:"横山可填补",options:{bold:true,fill:{color:C.acc},color:C.bg,fontSize:10}}],
  ["高端民宿集群","低","中","★★★★★","横云山居32栋·10-15栋集群"],
  ["文化研学","极低","低","★★★★★","黄公望IP+横山小学+米旅"],
  ["骑行配套","极低","低","★★★★★","8km环山+驿站+夜骑"],
  ["乡村夜经济","极低","极低","★★★★★","夜骑+暗夜星空+皮影+深夜食堂"],
  ["森林疗愈","极低","极低","★★★★★","40亩林地+唐代疗愈四法"],
];
let mH = mkRows[0]; let mB = mkRows.slice(1).map(r=>r.map(c=>typeof c==="string"?{text:c,options:{fontSize:9,color:C.h1,fontFace:"Microsoft YaHei"}}:c));
s4.addTable([mH,...mB], { x:0.3, y:2.6, w:12.7, colW:[2.5,1.5,1.2,2,5.5], border:{type:"solid",pt:0.5,color:C.dg} });
snum(s4);

// ═══ S5: CUSTOMER ═══
let s5 = pptx.addSlide(); s5.background = { fill: C.pri };
s5.addText("04  谁在买单", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
s5.addText("四类核心客群 · 稳定期年客流20-40万人次", { x: 0.8, y: 1.2, w: 10, h: 0.8, fontSize: 22, bold: true, color: C.h1, fontFace: "Microsoft YaHei" });
const pers = [
  {label:"亲子家庭 40%", icon:"👨‍👩‍👧", data:"30-45岁·有车·年4-8次\n松江约2.2万高收入有车家庭\n市场规模约3.5-13.2亿元/年", x:0.3, y:2.5, w:3.1},
  {label:"中产个人 30%", icon:"🧑‍💼", data:"28-45岁·单身/情侣·有车\n可下班来·年6-12次\n约7.5万潜在人群", x:3.5, y:2.5, w:3.1},
  {label:"大学城师生 15%", icon:"🎓", data:"18-25岁·拼车/公交\n年2-4次·约2.7万人\n社交传播力极强", x:6.7, y:2.5, w:3.1},
  {label:"企业团建 15%", icon:"🏢", data:"20-100人·工作日\n松江约1.3万家有预算\n市场规模约3.3-19.5亿元/年", x:10, y:2.5, w:3.1},
];
pers.forEach(p => {
  s5.addShape("roundRect", { x: p.x, y: p.y, w: p.w, h: 3.5, fill: { color: "162D42" }, rectRadius: 0.1 });
  s5.addText(p.icon, { x: p.x + 0.1, y: p.y + 0.2, w: 0.6, h: 0.5, fontSize: 22, align: "center" });
  s5.addText(p.label, { x: p.x + 0.8, y: p.y + 0.2, w: p.w - 1, h: 0.5, fontSize: 13, bold: true, color: C.acc, fontFace: "Microsoft YaHei" });
  s5.addText(p.data, { x: p.x + 0.3, y: p.y + 1, w: p.w - 0.6, h: 2.2, fontSize: 10, color: C.bd, fontFace: "Microsoft YaHei" });
});
snum(s5);

// ═══ S6: COMPETITIVE ═══
let s6 = pptx.addSlide(); s6.background = { fill: C.pri };
s6.addText("05  竞争者做不到的事", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
const compRows = [
  [{text:"维度",options:{bold:true,fill:{color:C.acc},color:C.bg,fontSize:10}},{text:"莫干山",options:{bold:true,fill:{color:C.acc},color:C.bg,fontSize:10}},{text:"计家墩",options:{bold:true,fill:{color:C.acc},color:C.bg,fontSize:10}},{text:"八十八亩田",options:{bold:true,fill:{color:C.acc},color:C.bg,fontSize:10}},{text:"横山村",options:{bold:true,fill:{color:C.acc},color:C.bg,fontSize:10}}],
  [{text:"车程",options:{fontSize:10,fill:{color:"162D42"},color:C.h1}},{text:"2h+",options:{fontSize:10,color:C.h1}},{text:"1h",options:{fontSize:10,color:C.h1}},{text:"50min",options:{fontSize:10,color:C.h1}},{text:"45min ⭐",options:{fontSize:10,color:C.acc,bold:true}}],
  [{text:"文化IP",options:{fontSize:10,fill:{color:"162D42"},color:C.h1}},{text:"无",options:{fontSize:10,color:C.h1}},{text:"设计师社群",options:{fontSize:10,color:C.h1}},{text:"无",options:{fontSize:10,color:C.h1}},{text:"黄公望·700年 ⭐",options:{fontSize:10,color:C.acc,bold:true}}],
  [{text:"地理标志",options:{fontSize:10,fill:{color:"162D42"},color:C.h1}},{text:"无",options:{fontSize:10,color:C.h1}},{text:"无",options:{fontSize:10,color:C.h1}},{text:"无",options:{fontSize:10,color:C.h1}},{text:"松江大米GI ⭐",options:{fontSize:10,color:C.acc,bold:true}}],
  [{text:"山体格局",options:{fontSize:10,fill:{color:"162D42"},color:C.h1}},{text:"不在村内",options:{fontSize:10,color:C.h1}},{text:"无",options:{fontSize:10,color:C.h1}},{text:"无",options:{fontSize:10,color:C.h1}},{text:"山居村中央 ⭐",options:{fontSize:10,color:C.acc,bold:true}}],
  [{text:"客流基础",options:{fontSize:10,fill:{color:"162D42"},color:C.h1}},{text:"红海",options:{fontSize:10,color:C.h1}},{text:"自建",options:{fontSize:10,color:C.h1}},{text:"自建",options:{fontSize:10,color:C.h1}},{text:"佘山1,570万·5km ⭐",options:{fontSize:10,color:C.acc,bold:true}}],
];
let cH = compRows[0]; let cB = compRows.slice(1).map(r=>r.map(c=>typeof c==="string"?{text:c,options:{fontSize:10,color:C.h1,fontFace:"Microsoft YaHei"}}:c));
s6.addTable([cH,...cB], { x:0.3, y:1.5, w:12.7, colW:[2,2.3,2.3,2.5,3.6], border:{type:"solid",pt:0.5,color:C.dg} });
s6.addText("别人只有一个剧场——横山整个村就是剧场。排列组合本身就是最深的差异化。", { x: 0.8, y: 6.2, w: 11, h: 0.4, fontSize: 13, italic: true, color: C.acc, fontFace: "Microsoft YaHei", align: "center" });
snum(s6);

// ═══ S7: GLOBAL ═══
let s7 = pptx.addSlide(); s7.background = { fill: C.pri };
s7.addText("06  全球对标：四个\"别人有、横山可以做得更好\"", { x: 0.8, y: 0.3, w: 8, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
const globs = [
  {t:"日本·稻田艺术",d:"田舍馆村年15-35万游客，需建24m观景塔(门票¥300)。横山有横云山——半山腰就是天然观景台，零成本。从半山腰看下去的视角，和700年前黄公望画《富春山居图》的视角是同一个。"},
  {t:"国际暗夜协会",d:"全球约40个认证暗夜社区(Dark Sky Community)，中国0个。横山位于佘山森林公园缓冲区内，光污染极少，冬季肉眼可见猎户座。若认证成功——中国首个暗夜社区。"},
  {t:"日本·Forest Adventure",d:"全国40+据点，利用现有森林做树上关卡(不砍树/不硬化/不建固定建筑)。横山40亩林地可做上海第一个。"},
  {t:"挪威·多场景米其林",d:"把一顿饭拆成五幕——船头/船尾/水底。横山把\"米的旅程\"拆成五幕——水边/田间/灶前/桌前/火边。同样的拆解逻辑，完全不同的文化内核。"},
];
globs.forEach((g, i) => {
  s7.addShape("roundRect", { x: 0.3, y: 1.2 + i * 1.5, w: 12.7, h: 1.2, fill: { color: i % 2 === 0 ? "162D42" : "1A3550" }, rectRadius: 0.06 });
  s7.addText(g.t, { x: 0.6, y: 1.25 + i * 1.5, w: 3.5, h: 0.4, fontSize: 12, bold: true, color: C.acc, fontFace: "Microsoft YaHei" });
  s7.addText(g.d, { x: 4.3, y: 1.25 + i * 1.5, w: 8.5, h: 1, fontSize: 10, color: C.bd, fontFace: "Microsoft YaHei" });
});
snum(s7);

// ═══ S8: PRODUCT LINES ═══
let s8 = pptx.addSlide(); s8.background = { fill: C.pri };
s8.addText("07  五大差异化业态线", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
const lines = [
  {t:"米文旅线",s:"松江大米GI·700年从种子到餐桌·插秧→收割→碾米→品鉴→米酒封坛→稻田晚宴。参考日本道の駅和棚田Owner制度。"},
  {t:"森林疗愈线",s:"唐代疗愈四法(香道/茶疗/画疗/声疗)·日本森林浴·皮质醇降低12-16%。上海首个基于真实森林的康养产品。"},
  {t:"黄公望文化线",s:"50岁辞官隐居横山·AI导览·画疗·艺术驻留。他不是一个700岁的古人——是每一个想转型的中年人的精神镜像。"},
  {t:"环山户外线",s:"8km环山闭环·骑行驿站·夜骑LED·日本Forest Adventure模式。上海50-80万骑行人口·最近有组织路线在淀山湖(1h+)。"},
  {t:"乡村夜经济线",s:"夜骑·暗夜星空·皮影戏·围炉夜话·深夜食堂。可申请中国首个暗夜社区认证。上海乡村晚上就睡了——横山不睡。"},
];
lines.forEach((l, i) => {
  s8.addShape("roundRect", { x: 0.3, y: 1.1 + i * 1.2, w: 12.7, h: 1, fill: { color: i % 2 === 0 ? "162D42" : "1A3550" }, rectRadius: 0.06 });
  s8.addText(l.t, { x: 0.6, y: 1.15 + i * 1.2, w: 2.5, h: 0.4, fontSize: 13, bold: true, color: C.acc, fontFace: "Microsoft YaHei" });
  s8.addText(l.s, { x: 3.3, y: 1.15 + i * 1.2, w: 9.5, h: 0.9, fontSize: 11, color: C.bd, fontFace: "Microsoft YaHei" });
});
snum(s8);

// ═══ S9: FIVE SENSES ═══
let s9 = pptx.addSlide(); s9.background = { fill: C.pri };
s9.addText("08  五感地图：横山可以用身体记住", { x: 0.8, y: 0.3, w: 8, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
const senses = [
  {icon:"👂",t:"听觉·7点",d:"竹林风声·稻田蛙鸣·碾米机嗡鸣·河道水流·小学钟声·清晨鸟鸣·雨打瓦片"},
  {icon:"👃",t:"嗅觉·4点",d:"气味银行提取四季气味——春之兰笋(4月)·夏之稻花(7月)·秋之新米·冬之松木"},
  {icon:"✋",t:"触觉·1线",d:"赤脚步道200米——碎石→草地→泥土→溪水。四种材质在脚底的变化。"},
  {icon:"👅",t:"味觉·2点",d:"品米室——北坡偏甜·南坡颗粒更大。同一座山，相距500米，味道不同。"},
  {icon:"👁",t:"视觉·4窗",d:"富春山居图四时窗——春窗(小学·4月)·夏窗(来云吧·7月)·秋窗(山居·10月)·冬窗(大院·12月)"},
];
senses.forEach((s, i) => {
  s9.addShape("roundRect", { x: 0.3, y: 1.1 + i * 1.2, w: 12.7, h: 1, fill: { color: i % 2 === 0 ? "162D42" : "1A3550" }, rectRadius: 0.06 });
  s9.addText(s.icon, { x: 0.5, y: 1.15 + i * 1.2, w: 0.5, h: 0.5, fontSize: 18 });
  s9.addText(s.t, { x: 1.1, y: 1.15 + i * 1.2, w: 2.5, h: 0.4, fontSize: 12, bold: true, color: C.acc, fontFace: "Microsoft YaHei" });
  s9.addText(s.d, { x: 3.8, y: 1.15 + i * 1.2, w: 9, h: 0.9, fontSize: 10, color: C.bd, fontFace: "Microsoft YaHei" });
});
snum(s9);

// ═══ S10: DINING ═══
let s10 = pptx.addSlide(); s10.background = { fill: C.pri };
s10.addText("09  横山的一顿饭：五幕·4小时·2.5km", { x: 0.8, y: 0.3, w: 8, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
const acts = [
  {n:"①",t:"水边·来云吧 17:00",d:"欢迎酒「黄公望的黄昏」。米酒×冷萃·杯底一粒糯米。看横云山在夕阳里变色。"},
  {n:"②",t:"田间·稻田中央 17:30",d:"手食——兰笋刺身+横山豆腐+米糕。站着吃，用手。风吹过稻田。"},
  {n:"③",t:"灶前·碾米坊 18:00",d:"看着今晚的米从稻谷变白米。碾米机嗡鸣。鲜碾米香只能在空气中停留几分钟。",hl:"你手心有一粒温热的生米——甜的。"},
  {n:"④",t:"桌前·横云山居 18:30",d:"七道米食料理——每道配不同田块/碾米度的米。第四道：你看着碾的那碗饭，距碾出2小时17分钟。",hl:"主厨说：「这是你看着碾的。」"},
  {n:"⑤",t:"火边·围炉 20:30",d:"米布丁+烤红薯+热米酒。火在响。不用说话。带走一袋标签印着「北坡第三块田·碾于今晚18:12」的米。"},
];
acts.forEach((a, i) => {
  s10.addShape("roundRect", { x: 0.3, y: 1.1 + i * 1.15, w: 12.7, h: 1, fill: { color: a.hl ? "1A3550" : (i % 2 === 0 ? "162D42" : "1A3550") }, rectRadius: 0.06, line: a.hl ? { color: C.acc, width: 1 } : undefined });
  s10.addText(a.n, { x: 0.5, y: 1.15 + i * 1.15, w: 0.5, h: 0.35, fontSize: 11, bold: true, color: C.acc, fontFace: "Microsoft YaHei" });
  s10.addText(a.t, { x: 1.1, y: 1.15 + i * 1.15, w: 3.5, h: 0.35, fontSize: 11, bold: true, color: C.h1, fontFace: "Microsoft YaHei" });
  s10.addText(a.d, { x: 4.8, y: 1.15 + i * 1.15, w: 8, h: 0.9, fontSize: 10, color: a.hl ? C.acc : C.bd, fontFace: "Microsoft YaHei" });
});
snum(s10);

// ═══ S11: SPATIAL ═══
let s11 = pptx.addSlide(); s11.background = { fill: C.pri };
s11.addText("10  空间规划：一心·一轴·四区·多节点", { x: 0.8, y: 0.3, w: 8, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
const zones = [
  {t:"核心接待区",s:"横山大院·1,563㎡+庭院800㎡\n到达第一站·文化展示·市集·邮驿·可停车约30辆",x:0.3,y:1.3,w:3.1},
  {t:"文化研学区",s:"横山小学·590㎡+庭院2亩\n研学基地·顾绣教室·版画工坊·皮影剧场·可停车约15辆",x:3.5,y:1.3,w:3.2},
  {t:"滨水休闲区",s:"来云吧·815㎡\n咖啡·骑行驿站·深夜食堂·元代驿站遗址·可停车约20辆",x:6.8,y:1.3,w:3.1},
  {t:"高端住宿区",s:"横云山居32栋·7,881㎡·总用地15,292㎡\n高端民宿集群·品米室·米酒工坊·围炉·暗夜星空·可停车约100辆",x:10.1,y:1.3,w:3.1},
];
zones.forEach(z => {
  s11.addShape("roundRect", { x: z.x, y: z.y, w: z.w, h: 4, fill: { color: "162D42" }, rectRadius: 0.08 });
  s11.addText(z.t, { x: z.x + 0.2, y: z.y + 0.2, w: z.w - 0.4, h: 0.4, fontSize: 12, bold: true, color: C.acc, fontFace: "Microsoft YaHei" });
  s11.addText(z.s, { x: z.x + 0.2, y: z.y + 0.7, w: z.w - 0.4, h: 3, fontSize: 9, color: C.bd, fontFace: "Microsoft YaHei" });
});
s11.addText("三条核心游线：米之旅 4km(步行/骑行) · 山之行 8km(骑行) · 林之愈 2km(步行)  |  停车总容量约165辆", { x: 0.8, y: 6.3, w: 11, h: 0.3, fontSize: 11, bold: true, color: C.acc, fontFace: "Microsoft YaHei", align: "center" });
snum(s11);

// ═══ S12: A+B+C ═══
let s12 = pptx.addSlide(); s12.background = { fill: C.pri };
s12.addText("11  A+B+C投资组合：基于边际成本效率", { x: 0.8, y: 0.3, w: 8, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
s12.addText("最优启动包：A+C = ¥93-103万 = 15条差异化业态", { x: 0.8, y: 1.2, w: 11, h: 0.8, fontSize: 22, bold: true, color: C.h1, fontFace: "Microsoft YaHei" });
s12.addText("不是先后顺序——是边际成本效率。A和C同时上线，开业即有15种体验。B的投入由A+C运营数据决定。", { x: 0.8, y: 2, w: 11, h: 0.3, fontSize: 11, italic: true, color: C.gy, fontFace: "Microsoft YaHei" });
const abcRows = [
  [{text:"级别",options:{bold:true,fill:{color:C.acc},color:C.bg}},{text:"投资",options:{bold:true,fill:{color:C.acc},color:C.bg}},{text:"条数",options:{bold:true,fill:{color:C.acc},color:C.bg}},{text:"包含",options:{bold:true,fill:{color:C.acc},color:C.bg}}],
  [{text:"A·基础盘",options:{fontSize:11,color:C.h1}},{text:"¥43-47万",options:{fontSize:11,color:C.h1}},{text:"8条",options:{fontSize:11,color:C.h1}},{text:"咖啡·品米·骑行·邮驿·邮花·地图·气味·集章",options:{fontSize:10,color:C.h1}}],
  [{text:"C·最优组合",options:{fontSize:11,color:C.h1}},{text:"¥49-56万",options:{fontSize:11,color:C.h1}},{text:"7条",options:{fontSize:11,color:C.h1}},{text:"雾境·夜骑·封坛·碳账户·物种·期货·驻留",options:{fontSize:10,color:C.h1}}],
  [{text:"A+C启动包",options:{bold:true,fontSize:12,fill:{color:"162D42"},color:C.acc}},{text:"¥93-103万",options:{bold:true,fontSize:12,fill:{color:"162D42"},color:C.acc}},{text:"15条",options:{bold:true,fontSize:12,fill:{color:"162D42"},color:C.acc}},{text:"开业即完整体验",options:{bold:true,fontSize:10,fill:{color:"162D42"},color:C.h1}}],
  [{text:"B·完整版",options:{fontSize:11,color:C.h1}},{text:"¥155-270万",options:{fontSize:11,color:C.h1}},{text:"7条",options:{fontSize:11,color:C.h1}},{text:"五幕晚餐·疗愈·AI导览·暗夜社区·野营·稻田晚宴·城市驿站",options:{fontSize:10,color:C.h1}}],
];
let aH = abcRows[0]; let aB = abcRows.slice(1).map(r=>r.map(c=>typeof c==="string"?{text:c,options:{fontSize:10,color:C.h1,fontFace:"Microsoft YaHei"}}:c));
s12.addTable([aH,...aB], { x:0.3, y:2.8, w:12.7, colW:[2.5,2.5,1.5,6.2], border:{type:"solid",pt:0.5,color:C.dg} });
snum(s12);

// ═══ S13: SOCIAL BENEFIT ═══
let s13 = pptx.addSlide(); s13.background = { fill: C.pri };
s13.addText("12  社会效益量化", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
const benRows = [
  [{text:"效益维度",options:{bold:true,fill:{color:C.acc},color:C.bg}},{text:"稳定期量化指标",options:{bold:true,fill:{color:C.acc},color:C.bg}}],
  ["直接就业","80-150个岗位（A+C阶段30-50岗·B阶段新增50-100岗：民宿管家/厨师/疗愈师/骑行向导/品米师/非遗工坊师傅）"],
  ["间接就业","约200-400个岗位（农产品供应/交通接驳/物流配送/周边零售/自由职业：摄影师/导游/内容创作者）"],
  ["村集体增收","年50-300万元（租金收入+碳账户管理费+自营业态分红·用于村公共设施维护和村民福利）"],
  ["村民增收","年人均增收5,000-15,000元（优先本地用工50-80岗+农产品采购松江大米/兰笋/蔬菜年采购额30-60万）"],
  ["非遗传承","培训5-10名传承人（竹编·王师傅68岁+顾绣传承人入驻）·夜校课程每年20-50名学员"],
  ["税收贡献","稳定期年税收约100-300万"],
  ["生态效益","40亩林地保持·碳汇约50-100吨/年·暗夜社区灯光改造·减少墙面硬化"],
];
let bH = benRows[0]; let bB = benRows.slice(1).map(r=>r.map(c=>({text:c,options:{fontSize:9,color:C.h1,fontFace:"Microsoft YaHei"}})));
s13.addTable([bH,...bB], { x:0.3, y:1.3, w:12.7, colW:[3,9.7], border:{type:"solid",pt:0.5,color:C.dg} });
snum(s13);

// ═══ S14: IMPLEMENTATION ═══
let s14 = pptx.addSlide(); s14.background = { fill: C.pri };
s14.addText("13  分阶段实施（24个月）", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
const phases = [
  {n:"1",t:"第1-3个月",tl:"筹备与顶层设计",d:"资产权属摸底·四方框架确认·控规边界·基建实测·锚定品牌预热·数字平台启动"},
  {n:"2",t:"第4-8个月",tl:"基建与集中招商",d:"公共基建施工·外立面改造·导牌安装·招商达60%·品牌方设计报审"},
  {n:"3",t:"第9-12个月",tl:"试运营与优化",d:"核心点位(大院/来云吧)率先试运营·品牌方装修进场·开幕活动·招商达85%"},
  {n:"4",t:"第13-24个月",tl:"全面运营与品牌升级",d:"全点位开业·四季活动日历·暗夜社区认证·争取示范村评定·B投入决策"},
];
phases.forEach((p, i) => {
  const y = 1.5 + i * 1.4;
  s14.addShape("roundRect", { x: 0.3, y, w: 1.5, h: 1.1, fill: { color: "162D42" }, rectRadius: 0.1 });
  s14.addText(p.n, { x: 0.3, y: y + 0.15, w: 1.5, h: 0.4, fontSize: 18, bold: true, color: C.acc, fontFace: "Microsoft YaHei", align: "center" });
  s14.addText(p.t, { x: 0.3, y: y + 0.55, w: 1.5, h: 0.3, fontSize: 8, color: C.bd, fontFace: "Microsoft YaHei", align: "center" });
  s14.addText(p.tl, { x: 2, y, w: 3.5, h: 0.4, fontSize: 13, bold: true, color: C.h1, fontFace: "Microsoft YaHei" });
  s14.addText(p.d, { x: 2, y: y + 0.4, w: 11, h: 0.6, fontSize: 10, color: C.bd, fontFace: "Microsoft YaHei" });
});
snum(s14);

// ═══ S15: OPERATIONS ═══
let s15 = pptx.addSlide(); s15.background = { fill: C.pri };
s15.addText("14  运营体系：横山邮驿 + 碳账户", { x: 0.8, y: 0.3, w: 8, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
s15.addText("横山邮驿（横山大院东厢房·60-80㎡）", { x: 0.8, y: 1.2, w: 8, h: 0.4, fontSize: 13, bold: true, color: C.h1, fontFace: "Microsoft YaHei" });
s15.addText("三大功能：到达台（取五感地图+激活碳账户）· 活动预订板（每日手写更新）· 积分兑换窗（碳积分→消费抵用券，面值¥5/10/20/50，所有商户通用，月结95%兑付）", { x: 0.8, y: 1.7, w: 11, h: 0.6, fontSize: 11, color: C.bd, fontFace: "Microsoft YaHei" });
s15.addText("碳账户（行为驱动复购引擎）", { x: 0.8, y: 2.6, w: 8, h: 0.4, fontSize: 13, bold: true, color: C.h1, fontFace: "Microsoft YaHei" });
s15.addText("骑车8km替代开车 → 减碳1.8kg → 18积分 → 咖啡减¥5。吃本地食材替代冷链 → 5积分 → 市集减¥3。拼车来替代独驾 → 减碳3.1kg → 31积分 → 住宿减¥30。积分只能在横山用——不是补贴，是复购。", { x: 0.8, y: 3.1, w: 11, h: 0.8, fontSize: 11, color: C.bd, fontFace: "Microsoft YaHei" });
s15.addText("获客渠道矩阵：小红书UGC(月200+篇)·抖音视频(月60条)·OTA套餐(30%预订)·私域节气内容(1万+用户)·佘山引流牌(30%客流)·事件营销(暗夜社区·横山雅集·稻田晚宴)", { x: 0.8, y: 4.2, w: 11, h: 0.6, fontSize: 11, color: C.bd, fontFace: "Microsoft YaHei" });
s15.addText("全年内容日历：春·清明兰笋猎人(全年最高光15天) → 夏·夏至夜骑(全年最长白昼+凉夏卡) → 秋·秋分收割品米(全年第二高光+稻田晚宴) → 冬·冬至围炉热米酒(全年最长的夜+暗夜社区认证)", { x: 0.8, y: 5.3, w: 11, h: 0.6, fontSize: 11, color: C.acc, fontFace: "Microsoft YaHei" });
snum(s15);

// ═══ S16: RISK ═══
let s16 = pptx.addSlide(); s16.background = { fill: C.pri };
s16.addText("15  风险与应对", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
const riskRows = [
  [{text:"风险",options:{bold:true,fill:{color:C.acc},color:C.bg}},{text:"影响",options:{bold:true,fill:{color:C.acc},color:C.bg}},{text:"应对",options:{bold:true,fill:{color:C.acc},color:C.bg}}],
  ["同质化竞争","客流不足","IP差异化+四季活动"],
  ["资金延迟","工期拖后","A+C仅¥93万即可开业"],
  ["品牌经营不善","空置率上升","6月退出+品牌储备库"],
  ["政策变化","补贴下降","多元收入模型"],
  ["村民参与低","关系紧张","优先用工+利益共享"],
];
let rH = riskRows[0]; let rB = riskRows.slice(1).map(r=>r.map(c=>({text:c,options:{fontSize:10,color:C.h1,fontFace:"Microsoft YaHei"}})));
s16.addTable([rH,...rB], { x:0.5, y:1.5, w:12.3, colW:[3,2.5,6.8], border:{type:"solid",pt:0.5,color:C.dg} });
snum(s16);

// ═══ S17: POLICY ═══
let s17 = pptx.addSlide(); s17.background = { fill: C.pri };
s17.addText("16  政策申报时间线", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
s17.addText("区级 · 松江区各专项资金", { x: 1, y: 1.2, w: 6, h: 0.3, fontSize: 12, bold: true, color: C.acc, fontFace: "Microsoft YaHei" });
s17.addText("旅游产业(最高200万)·现代农业(30-50%补贴)·文化产业(最高300万)·乡村振兴(按项目) → Q1-Q2集中申报", { x: 1, y: 1.6, w: 11, h: 0.4, fontSize: 11, color: C.bd, fontFace: "Microsoft YaHei" });
s17.addText("市级 · 上海市专项资金", { x: 1, y: 2.3, w: 6, h: 0.3, fontSize: 12, bold: true, color: C.acc, fontFace: "Microsoft YaHei" });
s17.addText("文创产业(最高500万·Q1末-Q2)·都市农业(年中)·乡村振兴示范村(提前一年盯·第2年认定)", { x: 1, y: 2.7, w: 11, h: 0.4, fontSize: 11, color: C.bd, fontFace: "Microsoft YaHei" });
s17.addText("申报策略：「组合拳」法——一个项目拆解为基建+民宿+农业+非遗多属性，从不同政策包匹配条款 → 一个项目·多重支持", { x: 0.8, y: 3.5, w: 11, h: 0.4, fontSize: 12, italic: true, color: C.acc, fontFace: "Microsoft YaHei", align: "center" });
s17.addText("申报联络：佘之春旅游发展有限公司 · 座机：57652901 · 联系人：刘雨洁 13817643468", { x: 0.8, y: 5.5, w: 11, h: 0.3, fontSize: 11, color: C.gy, fontFace: "Microsoft YaHei" });
snum(s17);

// ═══ S18: DECISION ═══
let s18 = pptx.addSlide(); s18.background = { fill: C.pri };
s18.addText("17  提请政府决策", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
const decs = ["确认四方结构合作框架（政府/平台/运营方/品牌方），明确运营方统筹地位","授权启动资产权属与基建条件实地核查（10项待核实事项）","确认A+C组合方案（¥93-103万·15条业态·3-12个月分批呈现）","协调规资局明确控规边界与生态保护红线","指定对口部门与联络人，建立月度联席会机制","授权运营方启动一期招商预热"];
decs.forEach((d, i) => { s18.addText((i+1) + ". " + d, { x: 1.2, y: 1.3 + i * 0.85, w: 10, h: 0.6, fontSize: 15, color: C.h1, fontFace: "Microsoft YaHei" }); });
snum(s18);

// ═══ S19: INVESTMENT ═══
let s19 = pptx.addSlide(); s19.background = { fill: C.pri };
s19.addText("18  投资概算", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
const invRows = [
  [{text:"投资类别",options:{bold:true,fill:{color:C.acc},color:C.bg}},{text:"责任主体",options:{bold:true,fill:{color:C.acc},color:C.bg}},{text:"估算(万元)",options:{bold:true,fill:{color:C.acc},color:C.bg}}],
  ["公共基础设施","政府/投资平台","1,500-3,000"],
  ["外立面改造","政府/投资平台","500-1,000"],
  ["数字平台","政府/投资平台","150-300"],
  ["A+C启动包","运营方","93-103"],
  ["品牌方室内改造","各品牌方","2,000-5,000"],
  [{text:"项目整体拉动投资",options:{bold:true,fill:{color:"162D42"},color:C.acc}},{text:"—",options:{fill:{color:"162D42"},color:C.h1}},{text:"4,350-9,700",options:{bold:true,fill:{color:"162D42"},color:C.acc}}],
];
let iH = invRows[0]; let iB = invRows.slice(1).map(r => r.map(c => typeof c === "string" ? {text:c,options:{fontSize:10,color:C.h1,fontFace:"Microsoft YaHei"}} : c));
s19.addTable([iH,...iB], { x:0.5, y:1.5, w:12.3, colW:[5,3.5,3.8], border:{type:"solid",pt:0.5,color:C.dg} });
snum(s19);

// ═══ S20: CLOSING ═══
let s20 = pptx.addSlide(); s20.background = { fill: C.bg };
s20.addShape("rect", { x: 1, y: 3.2, w: 4, h: 0.05, fill: { color: C.acc } });
s20.addText("感谢聆听", { x: 1, y: 1.5, w: 11, h: 1.2, fontSize: 48, bold: true, color: C.h1, fontFace: "Microsoft YaHei", align: "center" });
s20.addText("环山而居 · 沪派江南  |  四方共建 · 共谋发展", { x: 1, y: 3.8, w: 11, h: 0.5, fontSize: 16, color: C.acc, fontFace: "Microsoft YaHei", align: "center" });
snum(s20);

const out = "output/Hengshan_PPT_v3.pptx";
await pptx.writeFile({ fileName: out });
const size = fs.statSync(out).size;
console.log("PPT v3: " + out + " (" + Math.round(size/1024) + " KB, " + sn + " slides)");
