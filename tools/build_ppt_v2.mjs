import pptxgen from "pptxgenjs";
import fs from "fs";

const pptx = new pptxgen();
pptx.defineLayout({ name: "WIDE", width: 13.333, height: 7.5 });
pptx.layout = "WIDE";

const C = { bg: "0F1F2F", pri: "1A3A5C", acc: "C9A96E", wh: "FFFFFF", lt: "E8EDF2", gy: "8899AA", dg: "5A6A7A", h1: "F5F8FC", bd: "A8B8C8" };
let sn = 0;
const snum = (s) => s.addText(String(++sn), { x: 12.5, y: 7.05, w: 0.6, h: 0.3, fontSize: 8, color: C.gy, align: "right", fontFace: "Microsoft YaHei" });

// ═══ S1: COVER ═══
let s1 = pptx.addSlide(); s1.background = { fill: C.bg };
s1.addShape("rect", { x: 1, y: 3.4, w: 4, h: 0.04, fill: { color: C.acc } });
s1.addText("横山不是又一个乡村民宿项目", { x: 1, y: 1.2, w: 10, h: 1.4, fontSize: 42, bold: true, color: C.h1, fontFace: "Microsoft YaHei" });
s1.addText("上海市松江区横山村乡村振兴项目 · 投资决策支撑报告", { x: 1, y: 2.8, w: 10, h: 0.5, fontSize: 18, color: C.bd, fontFace: "Microsoft YaHei" });
s1.addText("2026年6月  |  运营方", { x: 1, y: 5.8, w: 10, h: 0.3, fontSize: 12, color: C.gy, fontFace: "Microsoft YaHei" });
snum(s1);

// ═══ S2: WHY HENGSHAN ═══
let s2 = pptx.addSlide(); s2.background = { fill: C.pri };
s2.addText("01  为什么是横山", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
s2.addShape("rect", { x: 0.8, y: 0.85, w: 1.5, h: 0.03, fill: { color: C.acc } });
s2.addText("上海唯一45分钟可达的近郊高频休闲目的地", { x: 0.8, y: 1.2, w: 11, h: 1, fontSize: 28, bold: true, color: C.h1, fontFace: "Microsoft YaHei" });
s2.addText("莫干山2h+ · 安吉2.5h+ · 横山45min  |  频次差5-10倍", { x: 0.8, y: 2.5, w: 8, h: 0.4, fontSize: 16, color: C.acc, fontFace: "Microsoft YaHei" });
s2.addText("黄公望50岁辞官隐居横山，画出《富春山居图》。700年后同一片田、同一品种的米。你来横山，吃的就是他吃过的米，看的就是他画过的山。", { x: 0.8, y: 5.5, w: 11, h: 0.4, fontSize: 12, italic: true, color: C.gy, fontFace: "Microsoft YaHei" });
snum(s2);

// ═══ S3: SIX CONDITIONS ═══
let s3 = pptx.addSlide(); s3.background = { fill: C.pri };
s3.addText("02  六个不可复制的条件", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
const conds = ["环山型村落\n上海唯一","黄公望隐居地\n700年","松江大米GI\n法律壁垒","佘山1,570万客流\n5km","近郊45分钟\n高频可达","万米可盘活\n村落级系统"];
conds.forEach((c, i) => {
  const col = i % 3, row = Math.floor(i / 3);
  s3.addShape("roundRect", { x: 0.5 + col * 4.1, y: 1.5 + row * 2.6, w: 3.8, h: 2.2, fill: { color: i % 2 === 0 ? "162D42" : "1A3550" }, rectRadius: 0.1, line: { color: C.acc, width: 1 } });
  s3.addText(String(i + 1), { x: 0.8 + col * 4.1, y: 1.6 + row * 2.6, w: 0.8, h: 0.6, fontSize: 24, bold: true, color: C.acc, fontFace: "Microsoft YaHei" });
  s3.addText(c, { x: 1.6 + col * 4.1, y: 1.7 + row * 2.6, w: 2.5, h: 1.8, fontSize: 13, color: C.h1, fontFace: "Microsoft YaHei" });
});
snum(s3);

// ═══ S4: MARKET ═══
let s4 = pptx.addSlide(); s4.background = { fill: C.pri };
s4.addText("03  市场在说话", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
s4.addShape("rect", { x: 0.8, y: 0.85, w: 1.5, h: 0.03, fill: { color: C.acc } });
s4.addText("全球乡村旅游约$1,000亿 · 中国约1.5万亿 · 松江乡村高端供给极度稀缺", { x: 0.8, y: 1.3, w: 11, h: 0.8, fontSize: 20, bold: true, color: C.h1, fontFace: "Microsoft YaHei" });
const mkRows = [
  [{text:"品类",options:{bold:true,fill:{color:C.acc},color:C.bg,fontSize:10}},{text:"供给密度",options:{bold:true,fill:{color:C.acc},color:C.bg,fontSize:10}},{text:"机会",options:{bold:true,fill:{color:C.acc},color:C.bg,fontSize:10}},{text:"横山可填补",options:{bold:true,fill:{color:C.acc},color:C.bg,fontSize:10}}],
  ["高端民宿集群","低","★★★★★","横云山居32栋·10-15栋集群"],
  ["文化研学","极低","★★★★★","黄公望IP+横山小学+米旅"],
  ["骑行配套","极低","★★★★★","8km环山闭环+驿站+夜骑"],
  ["乡村夜经济","极低","★★★★★","夜骑+暗夜星空+皮影+深夜食堂"],
  ["森林疗愈","极低","★★★★★","40亩林地+唐代疗愈四法"],
];
let mkHd = mkRows[0]; let mkBd = mkRows.slice(1).map(r=>r.map(c=>typeof c==="string"?{text:c,options:{fontSize:9,color:C.h1,fontFace:"Microsoft YaHei"}}:c));
s4.addTable([mkHd,...mkBd], { x:0.5, y:2.5, w:12.3, colW:[3,2,2,5.3], border:{type:"solid",pt:0.5,color:C.dg} });
snum(s4);

// ═══ S5: CUSTOMER ═══
let s5 = pptx.addSlide(); s5.background = { fill: C.pri };
s5.addText("04  谁在买单", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
s5.addText("四类核心客群 · 年客流20-40万人次", { x: 0.8, y: 1.2, w: 10, h: 0.8, fontSize: 22, bold: true, color: C.h1, fontFace: "Microsoft YaHei" });
const pers = [
  {label:"亲子家庭 40%",desc:"30-45岁·有车\n年4-8次·人均¥200-500\n松江约2.2万高收入有车家庭",x:0.5,y:2.5},
  {label:"中产个人 30%",desc:"28-45岁·单身/情侣\n可下班来·年6-12次\n约7.5万潜在人群",x:6.8,y:2.5},
  {label:"大学城师生 15%",desc:"18-25岁·拼车/公交\n年2-4次·约2.7万人\n社交传播力极强",x:0.5,y:4.5},
  {label:"企业团建 15%",desc:"20-100人·工作日\n松江约1.3万家有预算\n客单价高·非高频",x:6.8,y:4.5},
];
pers.forEach(p => {
  s5.addShape("roundRect", { x: p.x, y: p.y, w: 5.8, h: 1.6, fill: { color: "162D42" }, rectRadius: 0.08 });
  s5.addText(p.label, { x: p.x + 0.3, y: p.y + 0.2, w: 5.2, h: 0.5, fontSize: 15, bold: true, color: C.acc, fontFace: "Microsoft YaHei" });
  s5.addText(p.desc, { x: p.x + 0.3, y: p.y + 0.7, w: 5.2, h: 0.8, fontSize: 11, color: C.bd, fontFace: "Microsoft YaHei" });
});
snum(s5);

// ═══ S6: DIFFERENTIATION ═══
let s6 = pptx.addSlide(); s6.background = { fill: C.pri };
s6.addText("05  差异化·竞品对比", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
const compRows = [
  [{text:"",options:{fontSize:10,fill:{color:"162D42"}}},{text:"莫干山",options:{bold:true,fontSize:11,fill:{color:"162D42"},color:C.acc}},{text:"计家墩",options:{bold:true,fontSize:11,fill:{color:"162D42"},color:C.acc}},{text:"八十八亩田",options:{bold:true,fontSize:11,fill:{color:"162D42"},color:C.acc}},{text:"横山",options:{bold:true,fontSize:11,fill:{color:C.acc},color:C.bg}}],
  [{text:"车程",options:{fontSize:10,fill:{color:"162D42"},color:C.h1}},{text:"2h+",options:{fontSize:10,color:C.h1}},{text:"1h",options:{fontSize:10,color:C.h1}},{text:"50min",options:{fontSize:10,color:C.h1}},{text:"45min",options:{fontSize:10,color:C.h1}}],
  [{text:"文化IP",options:{fontSize:10,fill:{color:"162D42"},color:C.h1}},{text:"无",options:{fontSize:10,color:C.h1}},{text:"设计师社群",options:{fontSize:10,color:C.h1}},{text:"无",options:{fontSize:10,color:C.h1}},{text:"黄公望·700年",options:{fontSize:10,color:C.h1}}],
  [{text:"地理标志",options:{fontSize:10,fill:{color:"162D42"},color:C.h1}},{text:"无",options:{fontSize:10,color:C.h1}},{text:"无",options:{fontSize:10,color:C.h1}},{text:"无",options:{fontSize:10,color:C.h1}},{text:"松江大米GI",options:{fontSize:10,color:C.h1}}],
  [{text:"山体",options:{fontSize:10,fill:{color:"162D42"},color:C.h1}},{text:"不在村内",options:{fontSize:10,color:C.h1}},{text:"无",options:{fontSize:10,color:C.h1}},{text:"无",options:{fontSize:10,color:C.h1}},{text:"山居村中央",options:{fontSize:10,color:C.h1}}],
  [{text:"客流基础",options:{fontSize:10,fill:{color:"162D42"},color:C.h1}},{text:"红海饱和",options:{fontSize:10,color:C.h1}},{text:"自建",options:{fontSize:10,color:C.h1}},{text:"自建",options:{fontSize:10,color:C.h1}},{text:"佘山1,570万·5km",options:{fontSize:10,color:C.h1}}],
];
s6.addTable(compRows, { x: 0.5, y: 1.5, w: 12.3, colW:[2,2.5,2.5,2.5,2.8], border:{type:"solid",pt:0.5,color:C.dg} });
s6.addText("别人只有一个剧场——横山整个村就是剧场。排列组合本身就是最深的差异化。", { x: 0.8, y: 6.2, w: 11, h: 0.4, fontSize: 13, italic: true, color: C.acc, fontFace: "Microsoft YaHei", align: "center" });
snum(s6);

// ═══ S7: GLOBAL BENCHMARKS ═══
let s7 = pptx.addSlide(); s7.background = { fill: C.pri };
s7.addText("06  全球对标", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
const globs = [
  {t:"日本·稻田艺术",d:"田舍馆村年15-35万游客，需建24m观景塔。横山有横云山——天然观景台，零成本。"},
  {t:"国际暗夜协会",d:"全球40个认证暗夜社区，中国0个。横山可申请成为中国第一个——全国级媒体事件。"},
  {t:"日本·Forest Adventure",d:"全国40+据点，利用现有森林做树上关卡。横山40亩林地可做上海第一个。"},
  {t:"日本·道の駅",d:"全国1,000+个农产品直卖所。横山邮驿就是横山版的道の駅——当天鲜碾+品米师推荐。"},
];
globs.forEach((g, i) => {
  s7.addShape("roundRect", { x: 0.5, y: 1.3 + i * 1.5, w: 12.3, h: 1.2, fill: { color: i % 2 === 0 ? "162D42" : "1A3550" }, rectRadius: 0.06 });
  s7.addText(g.t, { x: 0.8, y: 1.35 + i * 1.5, w: 3, h: 0.4, fontSize: 13, bold: true, color: C.acc, fontFace: "Microsoft YaHei" });
  s7.addText(g.d, { x: 4, y: 1.35 + i * 1.5, w: 8.5, h: 1, fontSize: 12, color: C.bd, fontFace: "Microsoft YaHei" });
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
  s8.addShape("roundRect", { x: 0.5, y: 1.2 + i * 1.2, w: 12.3, h: 1, fill: { color: i % 2 === 0 ? "162D42" : "1A3550" }, rectRadius: 0.06 });
  s8.addText(l.t, { x: 0.8, y: 1.25 + i * 1.2, w: 2.5, h: 0.5, fontSize: 13, bold: true, color: C.acc, fontFace: "Microsoft YaHei" });
  s8.addText(l.s, { x: 3.5, y: 1.25 + i * 1.2, w: 9, h: 0.9, fontSize: 11, color: C.bd, fontFace: "Microsoft YaHei" });
});
snum(s8);

// ═══ S9: FIVE SENSES ═══
let s9 = pptx.addSlide(); s9.background = { fill: C.pri };
s9.addText("08  五感地图：16个感官点位", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
const senses = [
  {icon:"👂",t:"听觉·7点",d:"竹林风声(北坡)·稻田蛙鸣(6-8月黄昏)·碾米机嗡鸣(秋收季)·河道水流(雨季)·钟声(小学)·鸟鸣(清晨)·雨打瓦片(大院)"},
  {icon:"👃",t:"嗅觉·4点",d:"气味银行提取四季气味——春之兰笋(4月)·夏之稻花(7月)·秋之新米(10月)·冬之松木(12-2月)·5ml小瓶装·¥398/套"},
  {icon:"✋",t:"触觉·1线",d:"赤脚步道200米——碎石(微痛)→草地(柔软)→泥土(湿润)→溪水(冰凉)。四种材质在脚底的变化。"},
  {icon:"👅",t:"味觉·2点",d:"品米室——北坡第三块田的米偏甜，南坡第一块田的米颗粒更大。同一座山，相距500米，味道不同。"},
  {icon:"👁",t:"视觉·4窗",d:"富春山居图四时窗——春窗(小学东窗·4月清晨)·夏窗(来云吧北面露台·7月黄昏)·秋窗(山居南坡·10月下午)·冬窗(大院西厢·12月午后)"},
];
senses.forEach((s, i) => {
  s9.addShape("roundRect", { x: 0.5, y: 1.1 + i * 1.2, w: 12.3, h: 1, fill: { color: i % 2 === 0 ? "162D42" : "1A3550" }, rectRadius: 0.06 });
  s9.addText(s.icon, { x: 0.7, y: 1.15 + i * 1.2, w: 0.5, h: 0.5, fontSize: 18 });
  s9.addText(s.t, { x: 1.3, y: 1.15 + i * 1.2, w: 2.5, h: 0.4, fontSize: 12, bold: true, color: C.acc, fontFace: "Microsoft YaHei" });
  s9.addText(s.d, { x: 4, y: 1.15 + i * 1.2, w: 8.5, h: 0.9, fontSize: 10, color: C.bd, fontFace: "Microsoft YaHei" });
});
snum(s9);

// ═══ S10: SPATIAL ═══
let s10 = pptx.addSlide(); s10.background = { fill: C.pri };
s10.addText("09  空间规划：一心·一轴·四区", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
const zones = [
  {t:"核心接待区",s:"横山大院·1,563㎡+庭院800㎡\n到达第一站·文化展示·市集·邮驿",x:0.5,y:1.5},
  {t:"文化研学区",s:"横山小学·590㎡+庭院2亩\n研学基地·顾绣教室·版画工坊·皮影剧场",x:6.8,y:1.5},
  {t:"滨水休闲区",s:"来云吧·815㎡\n咖啡·骑行驿站·深夜食堂·三岔路口元代驿站遗址",x:0.5,y:3.5},
  {t:"高端住宿区",s:"横云山居32栋·7,881㎡·总用地15,292㎡\n高端民宿集群·品米室·米酒工坊·围炉·暗夜星空",x:6.8,y:3.5},
];
zones.forEach(z => {
  s10.addShape("roundRect", { x: z.x, y: z.y, w: 5.8, h: 1.6, fill: { color: "162D42" }, rectRadius: 0.08 });
  s10.addText(z.t, { x: z.x + 0.3, y: z.y + 0.15, w: 5.2, h: 0.4, fontSize: 14, bold: true, color: C.acc, fontFace: "Microsoft YaHei" });
  s10.addText(z.s, { x: z.x + 0.3, y: z.y + 0.6, w: 5.2, h: 0.9, fontSize: 10, color: C.bd, fontFace: "Microsoft YaHei" });
});
s10.addText("三条游线：米之旅(4km)·山之行(8km骑行)·林之愈(2km步行)", { x: 0.8, y: 6.3, w: 11, h: 0.3, fontSize: 12, bold: true, color: C.acc, fontFace: "Microsoft YaHei", align: "center" });
snum(s10);

// ═══ S11: DISTRIBUTED DINING ═══
let s11 = pptx.addSlide(); s11.background = { fill: C.pri };
s11.addText("10  横山的一顿饭：五幕分布式晚餐", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
const acts = [
  {n:"第一幕",t:"水边·来云吧",d:"欢迎酒「黄公望的黄昏」。米酒×冷萃·杯底一粒糯米。看横云山在夕阳里变色。"},
  {n:"第二幕",t:"田间·稻田中央",d:"手食——兰笋刺身+横山豆腐+新鲜米糕。站着吃，用手。风吹过稻田。"},
  {n:"第三幕",t:"灶前·碾米坊",d:"看着今晚的米从稻谷变白米。碾米机嗡鸣。鲜碾米香。","hl":"你手心有一粒温热的生米——甜的。"},
  {n:"第四幕",t:"桌前·横云山居",d:"七道米食料理。第四道——你看着碾的那碗饭，距碾出2小时17分钟。","hl":"主厨说：这是你看着碾的。"},
  {n:"第五幕",t:"火边·围炉",d:"米布丁+烤红薯+热米酒。火在响。不用说话。带走一袋标签印着「北坡第三块田·碾于今晚18:12」的米。"},
];
acts.forEach((a, i) => {
  s11.addShape("roundRect", { x: 0.5, y: 1.1 + i * 1.2, w: 12.3, h: 1, fill: { color: i === 2 ? "1A3550" : (i % 2 === 0 ? "162D42" : "1A3550") }, rectRadius: 0.06, line: a.hl ? { color: C.acc, width: 1 } : undefined });
  s11.addText(a.n, { x: 0.7, y: 1.15 + i * 1.2, w: 1, h: 0.35, fontSize: 10, bold: true, color: C.acc, fontFace: "Microsoft YaHei" });
  s11.addText(a.t, { x: 1.8, y: 1.15 + i * 1.2, w: 2.5, h: 0.35, fontSize: 11, bold: true, color: C.h1, fontFace: "Microsoft YaHei" });
  s11.addText((a.hl ? "★ " : "") + a.d, { x: 4.5, y: 1.15 + i * 1.2, w: 8, h: 0.9, fontSize: 10, color: a.hl ? C.acc : C.bd, fontFace: "Microsoft YaHei" });
});
snum(s11);

// ═══ S12: A+B+C ═══
let s12 = pptx.addSlide(); s12.background = { fill: C.pri };
s12.addText("11  A+B+C投资组合", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
s12.addText("最优启动包：A+C = ¥93-103万 = 15条差异化业态", { x: 0.8, y: 1.2, w: 11, h: 0.8, fontSize: 22, bold: true, color: C.h1, fontFace: "Microsoft YaHei" });
const abcRows = [
  [{text:"级别",options:{bold:true,fill:{color:C.acc},color:C.bg}},{text:"定义",options:{bold:true,fill:{color:C.acc},color:C.bg}},{text:"条数",options:{bold:true,fill:{color:C.acc},color:C.bg}},{text:"投资",options:{bold:true,fill:{color:C.acc},color:C.bg}}],
  [{text:"A·基础盘",options:{fontSize:10,color:C.h1}},{text:"立即启动·<¥50万/项·3个月上线",options:{fontSize:10,color:C.h1}},{text:"8条",options:{fontSize:10,color:C.h1}},{text:"¥43-47万",options:{fontSize:10,color:C.h1}}],
  [{text:"C·最优组合",options:{fontSize:10,color:C.h1}},{text:"边际成本趋零·A铺好的路",options:{fontSize:10,color:C.h1}},{text:"7条",options:{fontSize:10,color:C.h1}},{text:"¥49-56万",options:{fontSize:10,color:C.h1}}],
  [{text:"A+C启动包",options:{bold:true,fontSize:11,fill:{color:"162D42"},color:C.acc}},{text:"15条业态·开业即完整体验",options:{fontSize:10,fill:{color:"162D42"},color:C.h1}},{text:"15条",options:{fontSize:10,fill:{color:"162D42"},color:C.h1}},{text:"¥93-103万",options:{bold:true,fontSize:11,fill:{color:"162D42"},color:C.acc}}],
  [{text:"B·完整版",options:{fontSize:10,color:C.h1}},{text:"能级跃升·A+C数据驱动",options:{fontSize:10,color:C.h1}},{text:"7条",options:{fontSize:10,color:C.h1}},{text:"¥155-270万",options:{fontSize:10,color:C.h1}}],
];
s12.addTable(abcRows, { x: 0.8, y: 2.5, w: 11.5, colW: [3,4,1.5,3], border: {type:"solid",pt:0.5,color:C.dg} });
s12.addText("不是先后顺序——是边际成本效率。A和C同时上线。", { x: 0.8, y: 6.2, w: 11, h: 0.3, fontSize: 12, italic: true, color: C.gy, fontFace: "Microsoft YaHei", align: "center" });
snum(s12);

// ═══ S13: SOCIAL BENEFIT ═══
let s13 = pptx.addSlide(); s13.background = { fill: C.pri };
s13.addText("12  社会效益量化", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
const benRows = [
  [{text:"效益维度",options:{bold:true,fill:{color:C.acc},color:C.bg}},{text:"稳定期量化指标",options:{bold:true,fill:{color:C.acc},color:C.bg}}],
  ["直接就业","80-150个岗位（A+C阶段30-50岗·B阶段新增50-100岗）"],
  ["间接就业","约200-400个岗位（农产品供应/交通/物流/周边服务业）"],
  ["村集体增收","年50-300万元（租金+碳账户管理费+分红）"],
  ["村民增收","年人均增收5,000-15,000元（优先本地用工+农产品采购）"],
  ["非遗传承","培训5-10名传承人·夜校课程每年20-50名学员"],
  ["税收贡献","稳定期年税收约100-300万"],
  ["生态效益","40亩林地保持+碳汇约50-100吨/年+暗夜社区光污染控制"],
];
let bh = benRows[0]; let bb = benRows.slice(1).map(r=>r.map(c=>typeof c==="string"?{text:c,options:{fontSize:10,color:C.h1,fontFace:"Microsoft YaHei"}}:c));
s13.addTable([bh,...bb], { x:0.5, y:1.5, w:12.3, colW:[3.5,8.8], border:{type:"solid",pt:0.5,color:C.dg} });
snum(s13);

// ═══ S14: IMPLEMENTATION ═══
let s14 = pptx.addSlide(); s14.background = { fill: C.pri };
s14.addText("13  分阶段实施", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
const phases = [
  {t:"第一阶段\n第1-3个月",tl:"筹备与顶层设计",d:"资产权属摸底·四方框架确认·控规边界确认·基建条件实测·锚定品牌预热·数字平台启动"},
  {t:"第二阶段\n第4-8个月",tl:"基建与集中招商",d:"公共基建施工·外立面改造·导牌安装·招商达60%·品牌方设计报审"},
  {t:"第三阶段\n第9-12个月",tl:"试运营与优化",d:"核心点位(大院/来云吧)率先试运营·品牌方装修进场·开幕活动·招商达85%·运营数据积累"},
  {t:"第四阶段\n第13-24个月",tl:"全面运营与品牌升级",d:"全点位开业·四季活动日历运行·暗夜社区认证申请·争取市/区级示范村评定·B阶段投入决策"},
];
phases.forEach((p, i) => {
  const y = 1.5 + i * 1.4;
  s14.addShape("roundRect", { x: 0.5, y, w: 2.8, h: 1.1, fill: { color: "162D42" }, rectRadius: 0.1 });
  s14.addText(p.t, { x: 0.5, y: y + 0.1, w: 2.8, h: 0.7, fontSize: 10, bold: true, color: C.acc, fontFace: "Microsoft YaHei", align: "center" });
  s14.addText(p.tl, { x: 3.5, y, w: 3, h: 0.4, fontSize: 12, bold: true, color: C.h1, fontFace: "Microsoft YaHei" });
  s14.addText(p.d, { x: 3.5, y: y + 0.4, w: 9.3, h: 0.6, fontSize: 10, color: C.bd, fontFace: "Microsoft YaHei" });
});
snum(s14);

// ═══ S15: POLICY ═══
let s15 = pptx.addSlide(); s15.background = { fill: C.pri };
s15.addText("14  政策申报时间线", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
s15.addText("区级", { x: 1, y: 1.3, w: 2, h: 0.4, fontSize: 12, bold: true, color: C.acc, fontFace: "Microsoft YaHei" });
s15.addText("松江区旅游产业专项资金（民宿/文旅·最高200万）→ Q1-Q2集中申报\n松江区都市现代农业专项资金（农业/观光·投资30-50%补贴）→ 年中\n松江区文化产业发展专项资金（文创/非遗·最高300万）→ Q1\n松江区乡村振兴专项资金（基础设施·按项目）→ 滚动申报", { x: 1, y: 1.8, w: 11, h: 1.8, fontSize: 11, color: C.bd, fontFace: "Microsoft YaHei" });
s15.addText("市级", { x: 1, y: 3.8, w: 2, h: 0.4, fontSize: 12, bold: true, color: C.acc, fontFace: "Microsoft YaHei" });
s15.addText("上海市文创产业扶持资金（最高500万）→ Q1末-Q2\n上海市都市农业专项资金 → 年中\n上海市乡村振兴示范村计划 → 提前一年盯·第2年认定", { x: 1, y: 4.3, w: 11, h: 1.3, fontSize: 11, color: C.bd, fontFace: "Microsoft YaHei" });
s15.addText("申报策略：「组合拳」法——一个项目拆解为基建+民宿+农业+非遗多属性，从不同政策包匹配条款 → 一个项目·多重支持", { x: 0.8, y: 6.2, w: 11, h: 0.3, fontSize: 12, italic: true, color: C.acc, fontFace: "Microsoft YaHei", align: "center" });
snum(s15);

// ═══ S16: RISK ═══
let s16 = pptx.addSlide(); s16.background = { fill: C.pri };
s16.addText("15  风险与应对", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
const riskRows = [
  [{text:"风险",options:{bold:true,fill:{color:C.acc},color:C.bg}},{text:"影响",options:{bold:true,fill:{color:C.acc},color:C.bg}},{text:"应对措施",options:{bold:true,fill:{color:C.acc},color:C.bg}}],
  ["同质化竞争加剧","招商困难·客流不足","坚持IP差异化(黄公望+松江大米)+四季活动日历"],
  ["资金到位延迟","基建工期拖后","分期实施·A+C仅¥93-103万即可开业"],
  ["品牌方经营不善","空置率上升","6个月退出机制+品牌储备库维护"],
  ["政策变化","补贴或支持力度下降","多元收入模型·不依赖单一政策"],
  ["村民参与度低","社区关系紧张","优先本地用工+农产品采购+利益共享"],
];
let rh = riskRows[0]; let rb = riskRows.slice(1).map(r=>r.map(c=>({text:c,options:{fontSize:10,color:C.h1,fontFace:"Microsoft YaHei"}})));
s16.addTable([rh,...rb], { x:0.5, y:1.5, w:12.3, colW:[3.5,3,5.8], border:{type:"solid",pt:0.5,color:C.dg} });
snum(s16);

// ═══ S17-18: DECISION ═══
let s17 = pptx.addSlide(); s17.background = { fill: C.pri };
s17.addText("16  提请政府决策", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
const decs = ["确认四方结构合作框架，明确运营方统筹地位","授权启动资产权属与基建条件实地核查","确认A+C组合方案（¥93-103万·15条业态）","协调规资局明确控规边界与生态保护红线","指定对口部门与联络人，建立月度联席会机制","授权运营方启动一期招商预热"];
decs.forEach((d, i) => { s17.addText((i+1) + ". " + d, { x: 1.2, y: 1.3 + i * 0.85, w: 10, h: 0.6, fontSize: 15, color: C.h1, fontFace: "Microsoft YaHei" }); });
snum(s17);

// ═══ S19: INVESTMENT OVERVIEW ═══
let s18 = pptx.addSlide(); s18.background = { fill: C.pri };
s18.addText("17  投资概算", { x: 0.8, y: 0.3, w: 6, h: 0.5, fontSize: 14, color: C.acc, fontFace: "Microsoft YaHei" });
const invRows = [
  [{text:"投资类别",options:{bold:true,fill:{color:C.acc},color:C.bg}},{text:"责任主体",options:{bold:true,fill:{color:C.acc},color:C.bg}},{text:"估算(万元)",options:{bold:true,fill:{color:C.acc},color:C.bg}}],
  ["公共基础设施","政府/投资平台","1,500-3,000"],
  ["外立面统一改造","政府/投资平台","500-1,000"],
  ["横山一点通数字平台","政府/投资平台","150-300"],
  ["A+C运营启动包","运营方","93-103"],
  ["品牌方室内改造","各品牌方","2,000-5,000"],
  [{text:"项目整体拉动投资",options:{bold:true,fill:{color:"162D42"},color:C.acc}},{text:"—",options:{fill:{color:"162D42"},color:C.h1}},{text:"4,350-9,700",options:{bold:true,fill:{color:"162D42"},color:C.acc}}],
];
let ih = invRows[0]; let ib = invRows.slice(1).map(r => r.map(c => typeof c === "string" ? {text:c,options:{fontSize:10,color:C.h1,fontFace:"Microsoft YaHei"}} : c));
s18.addTable([ih,...ib], { x:0.5, y:1.5, w:12.3, colW:[5,3.5,3.8], border:{type:"solid",pt:0.5,color:C.dg} });
snum(s18);

// ═══ S20: CLOSING ═══
let s20 = pptx.addSlide(); s20.background = { fill: C.bg };
s20.addShape("rect", { x: 1, y: 3.2, w: 4, h: 0.04, fill: { color: C.acc } });
s20.addText("感谢聆听", { x: 1, y: 1.5, w: 11, h: 1.2, fontSize: 44, bold: true, color: C.h1, fontFace: "Microsoft YaHei", align: "center" });
s20.addText("环山而居 · 沪派江南  |  四方共建 · 共谋发展", { x: 1, y: 3.8, w: 11, h: 0.5, fontSize: 16, color: C.acc, fontFace: "Microsoft YaHei", align: "center" });
snum(s20);

const out = "output/Hengshan_PPT_v2.pptx";
await pptx.writeFile({ fileName: out });
const size = fs.statSync(out).size;
console.log("PPT v2: " + out + " (" + Math.round(size/1024) + " KB, " + sn + " slides)");
