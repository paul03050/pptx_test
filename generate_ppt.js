const pptxgen = require("pptxgenjs");

async function generate() {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";

  const C = {
    primary: "028090",
    secondary: "00A896",
    accent: "02C39A",
    dark: "024950",
    white: "FFFFFF",
    lightBg: "F0FCFA",
    text: "1A1A1A",
    muted: "6B7280",
    border: "E2E8F0",
  };

  // ===================== SLIDE 1: COVER =====================
  const s1 = pptx.addSlide();
  s1.background = { color: C.dark };

  s1.addShape(pptx.ShapeType.ellipse, { x: 9.5, y: -2, w: 6, h: 6, fill: { color: C.primary, transparency: 50 } });
  s1.addShape(pptx.ShapeType.ellipse, { x: -2.5, y: 4.5, w: 7, h: 7, fill: { color: C.primary, transparency: 60 } });
  s1.addShape(pptx.ShapeType.rect, { x: 0, y: 6.5, w: 13.33, h: 1, fill: { color: C.primary, transparency: 40 } });

  s1.addText("許博宇", {
    x: 1.2, y: 1.5, w: 10.93, h: 1.4,
    fontSize: 56, fontFace: "Arial Black", color: C.white, bold: true, align: "left", valign: "middle",
  });
  s1.addText("Paul Hsu", {
    x: 1.2, y: 2.7, w: 10.93, h: 0.8,
    fontSize: 28, fontFace: "Arial", color: C.accent, align: "left", valign: "middle",
  });
  s1.addText("財務會計  |  Financial Accountant", {
    x: 1.2, y: 3.6, w: 10.93, h: 0.7,
    fontSize: 20, fontFace: "Arial", color: C.secondary, align: "left", valign: "middle",
  });
  s1.addShape(pptx.ShapeType.rect, { x: 1.2, y: 4.4, w: 2.5, h: 0.04, fill: { color: C.secondary } });

  s1.addText("Email:  paul830305@gmail.com", {
    x: 1.2, y: 4.7, w: 10.93, h: 0.4,
    fontSize: 13, fontFace: "Arial", color: C.white, align: "left", valign: "middle",
  });
  s1.addText("Phone:  0975-470-790", {
    x: 1.2, y: 5.1, w: 10.93, h: 0.4,
    fontSize: 13, fontFace: "Arial", color: C.white, align: "left", valign: "middle",
  });
  s1.addText("Location:  台北市中山區", {
    x: 1.2, y: 5.5, w: 10.93, h: 0.4,
    fontSize: 13, fontFace: "Arial", color: C.white, align: "left", valign: "middle",
  });

  // ===================== SLIDE 2: ABOUT ME =====================
  function addLightSlide(title, cb) {
    const s = pptx.addSlide();
    s.background = { color: C.white };
    s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.35, h: 7.5, fill: { color: C.primary } });
    s.addText(title, {
      x: 0.8, y: 0.4, w: 12, h: 0.7,
      fontSize: 30, fontFace: "Arial Black", color: C.primary, bold: true, align: "left", valign: "middle",
    });
    s.addShape(pptx.ShapeType.rect, { x: 0.8, y: 1.1, w: 2, h: 0.05, fill: { color: C.secondary } });
    if (cb) cb(s);
    return s;
  }

  function card(s, x, y, w, h) {
    s.addShape(pptx.ShapeType.roundRect, { x, y, w, h, fill: { color: C.lightBg }, rectRadius: 0.15, line: { color: C.border, width: 0.5 } });
  }

  const s2 = addLightSlide("關於我  |  About Me");

  // Personal info card
  card(s2, 0.8, 1.5, 5.5, 2.2);
  s2.addText("個人資料", { x: 1.1, y: 1.65, w: 4.9, h: 0.4, fontSize: 16, fontFace: "Arial Black", color: C.primary, bold: true, align: "left" });
  const infoData = [["年齡", "32 歲"], ["役別", "免役"], ["居住地", "台北市中山區"], ["英文姓名", "Paul"], ["駕照", "輕型機車、普通小型車"]];
  infoData.forEach((d, i) => {
    const iy = 2.1 + i * 0.3;
    s2.addText(d[0], { x: 1.1, y: iy, w: 1.2, h: 0.3, fontSize: 11, fontFace: "Arial", color: C.muted, align: "left", valign: "middle" });
    s2.addText(d[1], { x: 2.3, y: iy, w: 3.7, h: 0.3, fontSize: 11, fontFace: "Arial", color: C.text, bold: true, align: "left", valign: "middle" });
  });

  // Education card
  card(s2, 0.8, 4.0, 5.5, 3.0);
  s2.addText("學歷", { x: 1.1, y: 4.15, w: 4.9, h: 0.4, fontSize: 16, fontFace: "Arial Black", color: C.primary, bold: true, align: "left" });
  const edu = [
    ["2022 - 2024", "台北商業大學", "會計資訊系 大學畢業"],
    ["2012 - 2016", "臺北市立大學", "資訊科學學系 肄業"],
    ["2009 - 2012", "市立大直高中", "普通科 高中畢業"],
  ];
  edu.forEach((e, i) => {
    const ey = 4.7 + i * 0.7;
    s2.addShape(pptx.ShapeType.ellipse, { x: 1.25, y: ey + 0.1, w: 0.14, h: 0.14, fill: { color: C.primary } });
    if (i < edu.length - 1) {
      s2.addShape(pptx.ShapeType.rect, { x: 1.305, y: ey + 0.24, w: 0.02, h: 0.55, fill: { color: C.border } });
    }
    s2.addText(e[0], { x: 1.6, y: ey - 0.05, w: 4.5, h: 0.2, fontSize: 9, fontFace: "Arial", color: C.muted, align: "left" });
    s2.addText(e[1], { x: 1.6, y: ey + 0.12, w: 4.5, h: 0.22, fontSize: 12, fontFace: "Arial", color: C.text, bold: true, align: "left", valign: "middle" });
    s2.addText(e[2], { x: 1.6, y: ey + 0.35, w: 4.5, h: 0.2, fontSize: 10, fontFace: "Arial", color: C.muted, align: "left", valign: "middle" });
  });

  // Languages card (right column)
  card(s2, 6.8, 1.5, 5.5, 2.2);
  s2.addText("語言能力  |  Languages", { x: 7.1, y: 1.65, w: 4.9, h: 0.4, fontSize: 16, fontFace: "Arial Black", color: C.primary, bold: true, align: "left" });
  const lang = [
    ["中文", "母語 / Native"],
    ["英文 (English)", "聽精通 說中等 讀精通 寫中等"],
    ["台語", "中等"],
  ];
  lang.forEach((l, i) => {
    const ly = 2.15 + i * 0.45;
    s2.addText(l[0], { x: 7.1, y: ly, w: 4.9, h: 0.2, fontSize: 12, fontFace: "Arial", color: C.text, bold: true, align: "left" });
    s2.addText(l[1], { x: 7.1, y: ly + 0.22, w: 4.9, h: 0.2, fontSize: 10, fontFace: "Arial", color: C.muted, align: "left" });
  });

  // Employment status card
  card(s2, 6.8, 4.0, 5.5, 1.5);
  s2.addText("求職狀態", { x: 7.1, y: 4.15, w: 4.9, h: 0.4, fontSize: 16, fontFace: "Arial Black", color: C.primary, bold: true, align: "left" });
  s2.addText("狀態：待業中，可立即上班", { x: 7.1, y: 4.6, w: 4.9, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.text, align: "left", valign: "middle" });
  s2.addText("經歷：3-4 年財務會計經驗", { x: 7.1, y: 4.95, w: 4.9, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.text, align: "left", valign: "middle" });

  // ===================== SLIDE 3: WORK EXPERIENCE =====================
  const s3 = addLightSlide("工作經歷  |  Work Experience");

  const jobs = [
    { title: "財務助理", company: "捷鵬國際股份有限公司", date: "2026/2 - 2026/6", dur: "5 個月", desc: "應付/應收帳款立沖帳、傳票編製、毛利率分析、VBA 優化" },
    { title: "主辦會計", company: "赫士盟食品企業有限公司", date: "2025/12", dur: "1 個月", desc: "營業稅申報、成本分析、科目餘額表與調節表編製" },
    { title: "會計專員", company: "仲駿股份有限公司", date: "2025/8", dur: "1 個月", desc: "營業稅申報、AP/AR沖帳、協助IPO跨部門協作" },
    { title: "主辦會計", company: "海峽會_陸台會股份有限公司", date: "2024/7 - 2025/6", dur: "1 年", desc: "收入稽核、POS系統導入、財務報表編制、401調節表" },
    { title: "會計審計", company: "永輝啟佳聯合會計師事務所", date: "2023/9 - 2024/2", dur: "6 個月", desc: "記帳、查帳、所得扣繳申報、營業稅申報" },
    { title: "會計專員", company: "尚楊財稅記帳士事務所", date: "2023/1 - 2023/8", dur: "8 個月", desc: "客戶帳務處理、營業稅及各類所得申報" },
    { title: "店員", company: "三商家購股份有限公司", date: "2021/10 - 2022/12", dur: "1 年 3 個月", desc: "收銀結帳、賣場陳列、進貨驗收與商品管理" },
  ];

  const tlX = 1.8;
  const tlY = 1.5;
  const tlSpacing = 0.78;

  s3.addShape(pptx.ShapeType.rect, { x: tlX, y: tlY, w: 0.03, h: jobs.length * tlSpacing, fill: { color: C.border } });

  jobs.forEach((job, i) => {
    const y = tlY + i * tlSpacing;
    s3.addShape(pptx.ShapeType.ellipse, { x: tlX - 0.06, y: y + 0.08, w: 0.16, h: 0.16, fill: { color: C.primary } });

    const cX = 2.3;
    const cH = 0.7;
    s3.addShape(pptx.ShapeType.roundRect, { x: cX, y, w: 10.2, h: cH, fill: { color: i % 2 === 0 ? C.white : C.lightBg }, rectRadius: 0.05, line: { color: C.border, width: 0.5 } });
    s3.addShape(pptx.ShapeType.rect, { x: cX, y, w: 0.06, h: cH, fill: { color: C.secondary } });

    s3.addText(job.title, { x: 2.55, y: y + 0.02, w: 1.8, h: 0.32, fontSize: 11, fontFace: "Arial Black", color: C.primary, bold: true, align: "left", valign: "middle" });
    s3.addText(job.company, { x: 2.55, y: y + 0.35, w: 3.3, h: 0.3, fontSize: 9, fontFace: "Arial", color: C.muted, align: "left", valign: "middle" });
    s3.addText(`${job.date}  (${job.dur})`, { x: 6.1, y: y + 0.02, w: 2.2, h: 0.32, fontSize: 9, fontFace: "Arial", color: C.muted, align: "left", valign: "middle" });
    s3.addText(job.desc, { x: 6.1, y: y + 0.35, w: 6.0, h: 0.3, fontSize: 9, fontFace: "Arial", color: C.text, align: "left", valign: "middle" });
  });

  // ===================== SLIDE 4: FINANCIAL ACCOUNTING SKILLS =====================
  const s4 = addLightSlide("專業能力  |  財務會計");

  const finSkills = [
    { title: "稅務申報", items: ["營業稅申報 & 401 調節表", "各類所得扣繳申報 & 調節表", "營所稅結算申報", "配合會計師查帳作業", "熟悉相關稅務法規"] },
    { title: "總帳 / ERP 系統", items: ["文中系統", "鼎新 Cosmos 平台", "金旭總帳系統", "POS 系統導入經驗", "ERP 轉換與重新開帳"] },
    { title: "財務會計核心", items: ["應付/應收帳款管理", "傳票編製 & 毛利率分析", "科目餘額表 & 調節表", "財務報表編制", "收入稽核 & 內控審核"] },
  ];

  const sw = 3.6, sh = 4.8, sy = 1.5, sg = 0.4, sx = 0.8;
  finSkills.forEach((sk, i) => {
    const x = sx + i * (sw + sg);
    card(s4, x, sy, sw, sh);
    s4.addShape(pptx.ShapeType.rect, { x, y: sy, w: sw, h: 0.06, fill: { color: C.primary } });

    s4.addShape(pptx.ShapeType.ellipse, { x: x + 1.3, y: sy + 0.3, w: 0.8, h: 0.8, fill: { color: C.primary } });
    s4.addText(`${i + 1}`, { x: x + 1.3, y: sy + 0.3, w: 0.8, h: 0.8, fontSize: 16, fontFace: "Arial Black", color: C.white, bold: true, align: "center", valign: "middle" });
    s4.addText(sk.title, { x: x + 0.2, y: sy + 1.3, w: sw - 0.4, h: 0.4, fontSize: 14, fontFace: "Arial Black", color: C.primary, bold: true, align: "center", valign: "middle" });

    sk.items.forEach((item, j) => {
      const iy = sy + 1.9 + j * 0.5;
      s4.addShape(pptx.ShapeType.ellipse, { x: x + 0.3, y: iy + 0.12, w: 0.07, h: 0.07, fill: { color: C.secondary } });
      s4.addText(item, { x: x + 0.5, y: iy, w: sw - 0.8, h: 0.35, fontSize: 11, fontFace: "Arial", color: C.text, align: "left", valign: "middle" });
    });
  });

  // ===================== SLIDE 5: IT SKILLS =====================
  const s5 = addLightSlide("專業能力  |  資訊應用");

  const itSkills = [
    { title: "Excel / VBA", items: ["函數應用 (SUM, XLOOKUP)", "樞紐分析 & 圖表製作", "VBA 自動化流程", "數據分析 & 報告製作"] },
    { title: "Python", items: ["基礎程式設計", "資料處理與分析", "自動化腳本撰寫"] },
    { title: "AI 工具應用", items: ["ChatGPT / Gemini / Claude", "提示詞工程", "AI Agent (OpenCode, Hermes)", "LLM 輔助工作流程"] },
    { title: "文書處理", items: ["Word、PowerPoint、Outlook", "中文打字 75~100 字/分", "CBAA Level I 內部稽核"] },
  ];

  const iw = 2.75, ih = 4.2, iy2 = 1.5, ig = 0.3, ix = 0.8;
  itSkills.forEach((sk, i) => {
    const x = ix + i * (iw + ig);
    card(s5, x, iy2, iw, ih);
    s5.addShape(pptx.ShapeType.rect, { x, y: iy2, w: iw, h: 0.06, fill: { color: C.primary } });
    s5.addText(sk.title, { x: x + 0.2, y: iy2 + 0.3, w: iw - 0.4, h: 0.4, fontSize: 13, fontFace: "Arial Black", color: C.primary, bold: true, align: "left", valign: "middle" });

    sk.items.forEach((item, j) => {
      const iy = iy2 + 0.9 + j * 0.55;
      s5.addShape(pptx.ShapeType.ellipse, { x: x + 0.3, y: iy + 0.12, w: 0.07, h: 0.07, fill: { color: C.secondary } });
      s5.addText(item, { x: x + 0.5, y: iy, w: iw - 0.8, h: 0.35, fontSize: 10, fontFace: "Arial", color: C.text, align: "left", valign: "middle" });
    });
  });

  // ===================== SLIDE 6: CAREER SUMMARY =====================
  const s6 = addLightSlide("自傳與職涯目標  |  Career Summary");

  // Core positioning
  card(s6, 0.8, 1.5, 5.8, 2.2);
  s6.addText("核心定位", { x: 1.1, y: 1.65, w: 5.2, h: 0.35, fontSize: 14, fontFace: "Arial Black", color: C.primary, bold: true, align: "left" });
  s6.addText("專精財務會計領域，持續培養資訊技術能力，以數位工具提升會計與財務管理效率。", {
    x: 1.1, y: 2.05, w: 5.2, h: 0.5, fontSize: 12, fontFace: "Arial", color: C.text, align: "left", valign: "top",
  });
  s6.addText("Specialized in financial accounting, with continuous IT capability development to leverage digital tools for enhanced efficiency.", {
    x: 1.1, y: 2.55, w: 5.2, h: 0.6, fontSize: 10, fontFace: "Arial", color: C.muted, align: "left", valign: "top",
  });

  // Key achievements
  card(s6, 0.8, 4.0, 5.8, 3.0);
  s6.addText("關鍵成就", { x: 1.1, y: 4.15, w: 5.2, h: 0.35, fontSize: 14, fontFace: "Arial Black", color: C.primary, bold: true, align: "left" });

  const achievements = [
    "ERP、POS 與總帳系統導入及轉換專案，確保系統順利銜接與帳務正常運作（餐飲業）",
    "協助建置創櫃公司內部控制制度（採購付款循環），完善傳票、憑證、用印等作業流程",
    "運用 AI 撰寫 VBA 優化 Excel 流程，提升資料處理效率",
  ];
  achievements.forEach((a, i) => {
    const ay = 4.65 + i * 0.7;
    s6.addShape(pptx.ShapeType.rect, { x: 1.3, y: ay + 0.02, w: 0.04, h: 0.55, fill: { color: C.secondary } });
    s6.addText(a, { x: 1.55, y: ay, w: 4.7, h: 0.7, fontSize: 11, fontFace: "Arial", color: C.text, align: "left", valign: "top" });
  });

  // Career goals (dark card, right)
  s6.addShape(pptx.ShapeType.roundRect, { x: 7.0, y: 1.5, w: 5.5, h: 5.5, fill: { color: C.dark }, rectRadius: 0.15 });
  s6.addText("職涯規劃", { x: 7.3, y: 1.7, w: 4.9, h: 0.4, fontSize: 16, fontFace: "Arial Black", color: C.accent, bold: true, align: "left" });
  s6.addText("Career Goal", { x: 7.3, y: 2.1, w: 4.9, h: 0.3, fontSize: 11, fontFace: "Arial", color: C.secondary, align: "left" });
  s6.addShape(pptx.ShapeType.rect, { x: 7.3, y: 2.5, w: 1.5, h: 0.03, fill: { color: C.secondary } });
  s6.addText(
    "未來希望持續深化財務分析與數位能力，朝 FP&A（財務規劃與分析）方向發展。\n\n短期目標是運用 Python 與 AI 工具串接 ERP 資料，建立自動化財務流程，降低人工整理與維護 Excel 報表的成本。\n\n中長期則希望投入資金管理、專案成本分析及財務模型建構，協助企業優化資源配置、提升營運效率。",
    { x: 7.3, y: 2.7, w: 4.9, h: 2.5, fontSize: 11, fontFace: "Arial", color: C.white, align: "left", valign: "top", lineSpacing: 16 }
  );
  s6.addText(
    "Future development: FP&A\nAutomate financial workflows with Python & AI\nCapital management, cost analysis & financial modeling",
    { x: 7.3, y: 5.3, w: 4.9, h: 1.4, fontSize: 10, fontFace: "Arial", color: C.secondary, align: "left", valign: "top", lineSpacing: 14 }
  );

  // ===================== SLIDE 7: CLOSING =====================
  const s7 = pptx.addSlide();
  s7.background = { color: C.dark };

  s7.addShape(pptx.ShapeType.ellipse, { x: -2, y: -2, w: 5, h: 5, fill: { color: C.primary, transparency: 50 } });
  s7.addShape(pptx.ShapeType.ellipse, { x: 10, y: 4, w: 5, h: 5, fill: { color: C.primary, transparency: 60 } });

  s7.addText("求職條件  |  Job Preferences", {
    x: 1.2, y: 0.5, w: 10.93, h: 0.8,
    fontSize: 32, fontFace: "Arial Black", color: C.white, bold: true, align: "left", valign: "middle",
  });
  s7.addShape(pptx.ShapeType.rect, { x: 1.2, y: 1.3, w: 2.5, h: 0.04, fill: { color: C.secondary } });

  // Left column
  const leftPrefs = [
    { label: "希望職稱", value: "財務會計" },
    { label: "希望職類", value: "稅務人員、財務會計助理\n記帳/出納/一般會計、查帳/審計人員" },
    { label: "希望產業", value: "金融機構、電子資訊/軟體/半導體\n一般製造業" },
  ];

  leftPrefs.forEach((p, i) => {
    const y = 1.7 + i * 1.5;
    s7.addShape(pptx.ShapeType.roundRect, { x: 1.2, y, w: 5.2, h: 1.2, fill: { color: C.primary, transparency: 50 }, rectRadius: 0.1 });
    s7.addText(p.label, { x: 1.5, y: y + 0.05, w: 4.6, h: 0.35, fontSize: 11, fontFace: "Arial", color: C.accent, bold: true, align: "left", valign: "middle" });
    s7.addText(p.value, { x: 1.5, y: y + 0.4, w: 4.6, h: 0.7, fontSize: 12, fontFace: "Arial", color: C.white, align: "left", valign: "top" });
  });

  // Right column
  const rightPrefs = [
    { label: "希望地點", value: "新北市、台北市" },
    { label: "上班時段", value: "日班、晚班、08:00~18:00" },
    { label: "遠端工作", value: "有意願" },
    { label: "待遇", value: "面議" },
  ];

  rightPrefs.forEach((p, i) => {
    const y = 1.7 + i * 1.2;
    s7.addShape(pptx.ShapeType.roundRect, { x: 6.8, y, w: 5.2, h: 0.95, fill: { color: C.primary, transparency: 50 }, rectRadius: 0.1 });
    s7.addText(p.label, { x: 7.1, y: y + 0.05, w: 4.6, h: 0.3, fontSize: 11, fontFace: "Arial", color: C.accent, bold: true, align: "left", valign: "middle" });
    s7.addText(p.value, { x: 7.1, y: y + 0.35, w: 4.6, h: 0.5, fontSize: 12, fontFace: "Arial", color: C.white, align: "left", valign: "middle" });
  });

  // Bottom bar
  s7.addShape(pptx.ShapeType.rect, { x: 0, y: 6.5, w: 13.33, h: 1, fill: { color: C.primary } });
  s7.addText("許博宇 Paul Hsu  |  paul830305@gmail.com  |  0975-470-790  |  台北市中山區", {
    x: 1.2, y: 6.5, w: 10.93, h: 1,
    fontSize: 13, fontFace: "Arial", color: C.white, align: "center", valign: "middle",
  });

  await pptx.writeFile({ fileName: "Paul_Resume.pptx" });
  console.log("✅ Generated: Paul_Resume.pptx");
}

generate().catch(console.error);
