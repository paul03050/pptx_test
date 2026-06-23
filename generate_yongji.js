const pptxgen = require("pptxgenjs");

async function generate() {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_16x9";
  pptx.author = "許博宇";
  pptx.title = "永吉電腦股份有限公司 業務介紹與ERP會計應用";
  pptx.subject = "面試簡報";

  // ═══════════════════════════════
  // COLOR PALETTE
  // ═══════════════════════════════
  const C = {
    navy: "0D4A7A",
    teal: "1B998B",
    amber: "F18F01",
    darkNavy: "0A2647",
    white: "FFFFFF",
    lightBg: "F5F7FA",
    text: "2D3748",
    muted: "718096",
    border: "E2E8F0",
    cardBg: "FFFFFF",
    darkText: "1A202C",
    lightMuted: "A0AEC0",
  };

  // ═══════════════════════════════
  // HELPER FUNCTIONS
  // ═══════════════════════════════

  // Creates a light content slide with title bar on left
  function addContentSlide(title, subtitle, cb) {
    const s = pptx.addSlide();
    s.background = { color: C.lightBg };

    // Left accent bar
    s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.35, h: 5.625, fill: { color: C.navy } });
    // Top right decorative accent
    s.addShape(pptx.ShapeType.rect, { x: 0.35, y: 0, w: 9.65, h: 0.04, fill: { color: C.teal } });

    // Title
    s.addText(title, {
      x: 0.7, y: 0.25, w: 8.5, h: 0.55,
      fontSize: 28, fontFace: "Microsoft JhengHei", color: C.navy, bold: true, align: "left", valign: "middle",
      margin: 0,
    });
    // Subtitle
    if (subtitle) {
      s.addText(subtitle, {
        x: 0.7, y: 0.75, w: 8.5, h: 0.35,
        fontSize: 14, fontFace: "Microsoft JhengHei", color: C.muted, align: "left", valign: "middle",
        margin: 0,
      });
    }
    // Underline accent
    s.addShape(pptx.ShapeType.rect, { x: 0.7, y: 1.05, w: 1.8, h: 0.04, fill: { color: C.teal } });

    if (cb) cb(s);
    return s;
  }

  function addCard(slide, x, y, w, h) {
    slide.addShape(pptx.ShapeType.roundRect, {
      x, y, w, h,
      fill: { color: C.cardBg },
      rectRadius: 0.1,
      line: { color: C.border, width: 0.5 },
      shadow: { type: "outer", color: "000000", blur: 4, offset: 1, angle: 135, opacity: 0.08 },
    });
  }

  // Add a card with left accent bar
  function addCardWithAccent(slide, x, y, w, h, accentColor) {
    addCard(slide, x, y, w, h);
    slide.addShape(pptx.ShapeType.rect, { x, y, w: 0.06, h, fill: { color: accentColor || C.teal } });
  }

  // Shortcut for creating bullet item text
  function bulletItem(text, opts) {
    return { text, options: { bullet: true, breakLine: true, ...opts } };
  }

  // ════════════════════════════════════════════
  // SLIDE 1: 封面
  // ════════════════════════════════════════════
  const s1 = pptx.addSlide();
  s1.background = { color: C.darkNavy };

  // Decorative circles
  s1.addShape(pptx.ShapeType.ellipse, { x: 7.5, y: -1.5, w: 4.5, h: 4.5, fill: { color: C.navy, transparency: 40 } });
  s1.addShape(pptx.ShapeType.ellipse, { x: -1.5, y: 3.5, w: 4, h: 4, fill: { color: C.navy, transparency: 50 } });

  // Top and bottom accent lines
  s1.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.teal } });
  s1.addShape(pptx.ShapeType.rect, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.teal } });

  // Brand tag
  s1.addText("KEYDEX", {
    x: 0.8, y: 0.4, w: 8.4, h: 0.45,
    fontSize: 14, fontFace: "Arial", color: C.teal, bold: true, align: "left", valign: "middle",
    charSpacing: 6,
    margin: 0,
  });

  // Main title
  s1.addText("永吉電腦股份有限公司", {
    x: 0.8, y: 1.1, w: 8.4, h: 1.0,
    fontSize: 40, fontFace: "Microsoft JhengHei", color: C.white, bold: true, align: "left", valign: "middle",
    margin: 0,
  });

  // Subtitle
  s1.addText("業務介紹 與 鼎新ERP會計應用", {
    x: 0.8, y: 2.1, w: 8.4, h: 0.6,
    fontSize: 22, fontFace: "Microsoft JhengHei", color: C.teal, align: "left", valign: "middle",
    margin: 0,
  });

  // Accent bar
  s1.addShape(pptx.ShapeType.rect, { x: 0.8, y: 2.85, w: 2.5, h: 0.04, fill: { color: C.amber } });

  // Presenter info
  s1.addText("主講人：許博宇  Paul Hsu", {
    x: 0.8, y: 3.2, w: 8.4, h: 0.4,
    fontSize: 16, fontFace: "Microsoft JhengHei", color: C.white, align: "left", valign: "middle",
    margin: 0,
  });
  s1.addText("財務會計  |  會計資訊系畢業", {
    x: 0.8, y: 3.6, w: 8.4, h: 0.35,
    fontSize: 13, fontFace: "Microsoft JhengHei", color: C.lightMuted, align: "left", valign: "middle",
    margin: 0,
  });
  s1.addText("2026", {
    x: 0.8, y: 4.6, w: 8.4, h: 0.35,
    fontSize: 14, fontFace: "Arial", color: C.lightMuted, align: "left", valign: "middle",
    margin: 0,
  });

  // ════════════════════════════════════════════
  // SLIDE 2: 公司概覽
  // ════════════════════════════════════════════
  addContentSlide("公司概覽", "Company Overview", (s) => {
    // Left info block
    addCardWithAccent(s, 0.7, 1.4, 4.2, 3.6, C.navy);

    const infoItems = [
      ["公司全名", "永吉電腦股份有限公司"],
      ["統一編號", "04916011"],
      ["成立時間", "1983 年（成立逾 40 年）"],
      ["資本總額", "NT$ 2,800 萬"],
      ["董事長", "范國英"],
      ["品牌", "KEYDEX（標竿）"],
      ["地址", "台北市信義區莊敬路"],
    ];

    infoItems.forEach((item, i) => {
      const iy = 1.6 + i * 0.43;
      s.addText(item[0], {
        x: 1.0, y: iy, w: 1.5, h: 0.35,
        fontSize: 11, fontFace: "Microsoft JhengHei", color: C.muted, align: "left", valign: "middle",
        margin: 0,
      });
      s.addText(item[1], {
        x: 2.5, y: iy, w: 2.3, h: 0.35,
        fontSize: 11, fontFace: "Microsoft JhengHei", color: C.text, bold: true, align: "left", valign: "middle",
        margin: 0,
      });
    });

    // Right info - relationship enterprises
    addCardWithAccent(s, 5.3, 1.4, 4.2, 1.6, C.teal);
    s.addText("關係企業", {
      x: 5.6, y: 1.5, w: 3.6, h: 0.35,
      fontSize: 15, fontFace: "Microsoft JhengHei", color: C.navy, bold: true, align: "left", valign: "middle",
      margin: 0,
    });

    const subs = [
      ["Ever-Shining", "複合式工廠（模具、射出、線材）"],
      ["iMate", "網路銷售公司（電腦週邊）"],
    ];
    subs.forEach((sub, i) => {
      const sy = 1.95 + i * 0.45;
      s.addShape(pptx.ShapeType.ellipse, { x: 5.7, y: sy + 0.08, w: 0.1, h: 0.1, fill: { color: C.teal } });
      s.addText(sub[0], {
        x: 5.95, y: sy, w: 2, h: 0.25,
        fontSize: 12, fontFace: "Microsoft JhengHei", color: C.text, bold: true, align: "left", valign: "middle",
        margin: 0,
      });
      s.addText(sub[1], {
        x: 5.95, y: sy + 0.22, w: 3.3, h: 0.22,
        fontSize: 10, fontFace: "Microsoft JhengHei", color: C.muted, align: "left", valign: "middle",
        margin: 0,
      });
    });

    // Right info - business items
    addCardWithAccent(s, 5.3, 3.3, 4.2, 1.7, C.amber);
    s.addText("營業項目", {
      x: 5.6, y: 3.4, w: 3.6, h: 0.35,
      fontSize: 15, fontFace: "Microsoft JhengHei", color: C.navy, bold: true, align: "left", valign: "middle",
      margin: 0,
    });

    const bizItems = [
      "電腦週邊設備生產銷售",
      "電腦配線 / Cable 製造",
      "KEYDEX NFC 智慧指環",
      "資訊處理與軟體設計",
      "進出口貿易",
    ];
    bizItems.forEach((item, i) => {
      const by = 3.85 + i * 0.24;
      s.addShape(pptx.ShapeType.ellipse, { x: 5.7, y: by + 0.06, w: 0.08, h: 0.08, fill: { color: C.muted } });
      s.addText(item, {
        x: 5.95, y: by, w: 3.3, h: 0.24,
        fontSize: 10, fontFace: "Microsoft JhengHei", color: C.text, align: "left", valign: "middle",
        margin: 0,
      });
    });
  });

  // ════════════════════════════════════════════
  // SLIDE 3: 主要業務 (一)
  // ════════════════════════════════════════════
  addContentSlide("主要業務（一）", "電腦週邊批發 與 電腦配線製造", (s) => {
    // Card 1: 電腦週邊批發
    addCardWithAccent(s, 0.7, 1.4, 4.2, 3.6, C.navy);
    s.addText("電腦週邊設備批發", {
      x: 1.0, y: 1.55, w: 3.6, h: 0.4,
      fontSize: 18, fontFace: "Microsoft JhengHei", color: C.navy, bold: true, align: "left", valign: "middle",
      margin: 0,
    });
    s.addShape(pptx.ShapeType.rect, { x: 1.0, y: 1.95, w: 1.2, h: 0.03, fill: { color: C.teal } });

    const wholesaleItems = [
      "記憶卡、隨身碟、顯示器",
      "滑鼠、鍵盤、光碟機",
      "硬碟、掃瞄器、印表機",
      "印表機用碳粉匣等耗材",
      "電腦文具用品（報表紙、磁片盒）",
    ];
    wholesaleItems.forEach((item, i) => {
      const iy = 2.15 + i * 0.42;
      s.addShape(pptx.ShapeType.ellipse, { x: 1.15, y: iy + 0.08, w: 0.08, h: 0.08, fill: { color: C.navy } });
      s.addText(item, {
        x: 1.4, y: iy, w: 3.3, h: 0.35,
        fontSize: 12, fontFace: "Microsoft JhengHei", color: C.text, align: "left", valign: "middle",
        margin: 0,
      });
    });

    // Card 2: Cable製造
    addCardWithAccent(s, 5.3, 1.4, 4.2, 3.6, C.teal);
    s.addText("電腦配線 / Cable 製造", {
      x: 5.6, y: 1.55, w: 3.6, h: 0.4,
      fontSize: 18, fontFace: "Microsoft JhengHei", color: C.teal, bold: true, align: "left", valign: "middle",
      margin: 0,
    });
    s.addShape(pptx.ShapeType.rect, { x: 5.6, y: 1.95, w: 1.2, h: 0.03, fill: { color: C.teal } });

    const cableItems = [
      "網路線 / 高速傳輸線",
      "Category 8 高速網路線",
      "電腦連接線、插座",
      "USB / HDMI 等配線",
      "OEM / ODM 客製化生產",
    ];
    cableItems.forEach((item, i) => {
      const iy = 2.15 + i * 0.42;
      s.addShape(pptx.ShapeType.ellipse, { x: 5.75, y: iy + 0.08, w: 0.08, h: 0.08, fill: { color: C.teal } });
      s.addText(item, {
        x: 6.0, y: iy, w: 3.3, h: 0.35,
        fontSize: 12, fontFace: "Microsoft JhengHei", color: C.text, align: "left", valign: "middle",
        margin: 0,
      });
    });

    // Bottom note
    s.addShape(pptx.ShapeType.roundRect, {
      x: 0.7, y: 5.1, w: 8.8, h: 0.4,
      fill: { color: C.navy },
      rectRadius: 0.05,
    });
    s.addText("品牌 KEYDEX（標竿）— 專注電腦週邊與配線領域逾 40 年", {
      x: 0.9, y: 5.1, w: 8.4, h: 0.4,
      fontSize: 12, fontFace: "Microsoft JhengHei", color: C.white, align: "center", valign: "middle",
      margin: 0,
    });
  });

  // ════════════════════════════════════════════
  // SLIDE 4: 主要業務 (二)
  // ════════════════════════════════════════════
  addContentSlide("主要業務（二）", "KEYDEX NFC 智慧指環 與 一條龍生產製造", (s) => {
    // Card 1: NFC智慧指環
    addCardWithAccent(s, 0.7, 1.4, 4.2, 3.6, C.amber);
    s.addText("KEYDEX NFC 智慧指環", {
      x: 1.0, y: 1.55, w: 3.6, h: 0.4,
      fontSize: 18, fontFace: "Microsoft JhengHei", color: C.amber, bold: true, align: "left", valign: "middle",
      margin: 0,
    });
    s.addShape(pptx.ShapeType.rect, { x: 1.0, y: 1.95, w: 1.2, h: 0.03, fill: { color: C.amber } });

    const nfcItems = [
      "NFC 個資保護指環",
      "鎖屏指環（電腦/手機解鎖）",
      "門禁指環（門禁感應）",
      "行動支付指環",
      "防走失救援指環",
      "i-PASS 一卡通支付",
    ];
    nfcItems.forEach((item, i) => {
      const iy = 2.15 + i * 0.38;
      s.addShape(pptx.ShapeType.ellipse, { x: 1.15, y: iy + 0.07, w: 0.08, h: 0.08, fill: { color: C.amber } });
      s.addText(item, {
        x: 1.4, y: iy, w: 3.3, h: 0.3,
        fontSize: 12, fontFace: "Microsoft JhengHei", color: C.text, align: "left", valign: "middle",
        margin: 0,
      });
    });

    // Card 2: 生產製造
    addCardWithAccent(s, 5.3, 1.4, 4.2, 3.6, C.navy);
    s.addText("一條龍生產製造", {
      x: 5.6, y: 1.55, w: 3.6, h: 0.4,
      fontSize: 18, fontFace: "Microsoft JhengHei", color: C.navy, bold: true, align: "left", valign: "middle",
      margin: 0,
    });
    s.addShape(pptx.ShapeType.rect, { x: 5.6, y: 1.95, w: 1.2, h: 0.03, fill: { color: C.navy } });

    // Process flow
    const processes = [
      ["01", "模具設計", "模具廠專業設計開發"],
      ["02", "塑膠射出", "射出廠量產成型"],
      ["03", "抽線加工", "抽線廠線材製造"],
      ["04", "成品組裝", "線材加工廠組裝檢驗"],
    ];
    processes.forEach((p, i) => {
      const py = 2.15 + i * 0.55;
      s.addShape(pptx.ShapeType.roundRect, {
        x: 5.7, y: py, w: 0.32, h: 0.32,
        fill: { color: C.navy },
        rectRadius: 0.04,
      });
      s.addText(p[0], {
        x: 5.7, y: py, w: 0.32, h: 0.32,
        fontSize: 10, fontFace: "Arial", color: C.white, bold: true, align: "center", valign: "middle",
        margin: 0,
      });
      s.addText(p[1], {
        x: 6.2, y: py, w: 1.3, h: 0.18,
        fontSize: 12, fontFace: "Microsoft JhengHei", color: C.text, bold: true, align: "left", valign: "middle",
        margin: 0,
      });
      s.addText(p[2], {
        x: 6.2, y: py + 0.18, w: 3.0, h: 0.16,
        fontSize: 10, fontFace: "Microsoft JhengHei", color: C.muted, align: "left", valign: "middle",
        margin: 0,
      });
      if (i < processes.length - 1) {
        s.addShape(pptx.ShapeType.rect, { x: 5.85, y: py + 0.34, w: 0.02, h: 0.19, fill: { color: C.border } });
      }
    });

    // Bottom note
    s.addShape(pptx.ShapeType.roundRect, {
      x: 0.7, y: 5.1, w: 8.8, h: 0.4,
      fill: { color: C.teal },
      rectRadius: 0.05,
    });
    s.addText("自有工廠：模具廠 → 塑膠射出廠 → 抽線廠 → 線材加工廠，一條龍服務", {
      x: 0.9, y: 5.1, w: 8.4, h: 0.4,
      fontSize: 12, fontFace: "Microsoft JhengHei", color: C.white, align: "center", valign: "middle",
      margin: 0,
    });
  });

  // ════════════════════════════════════════════
  // SLIDE 5: 業務版圖與關係企業
  // ════════════════════════════════════════════
  addContentSlide("業務版圖與關係企業", "Business Landscape & Subsidiaries", (s) => {
    // Three cards for the group structure
    const groupInfo = [
      {
        title: "永吉電腦股份有限公司",
        items: ["品牌 KEYDEX 經營", "電腦週邊批發/配線", "NFC 智慧指環研發", "進出口貿易"],
        accent: C.navy,
      },
      {
        title: "Ever-Shining",
        subtitle: "複合式工廠",
        items: ["模具設計開發", "塑膠射出成型", "抽線及線材加工", "成品檢驗出廠"],
        accent: C.teal,
      },
      {
        title: "iMate",
        subtitle: "網路銷售",
        items: ["電腦週邊商品銷售", "台灣官網營運", "中國市場推廣", "線上通路經營"],
        accent: C.amber,
      },
    ];

    groupInfo.forEach((g, i) => {
      const gx = 0.7 + i * 3.15;

      // Card with accent
      addCardWithAccent(s, gx, 1.4, 2.95, 3.0, g.accent);

      // Title
      s.addText(g.title, {
        x: gx + 0.2, y: 1.55, w: 2.55, h: 0.35,
        fontSize: 15, fontFace: "Microsoft JhengHei", color: g.accent, bold: true, align: "left", valign: "middle",
        margin: 0,
      });
      if (g.subtitle) {
        s.addText(g.subtitle, {
          x: gx + 0.2, y: 1.85, w: 2.55, h: 0.25,
          fontSize: 11, fontFace: "Microsoft JhengHei", color: C.muted, align: "left", valign: "middle",
          margin: 0,
        });
      }

      // Items
      g.items.forEach((item, j) => {
        const iy = 2.3 + j * 0.4;
        s.addShape(pptx.ShapeType.ellipse, { x: gx + 0.35, y: iy + 0.08, w: 0.07, h: 0.07, fill: { color: g.accent } });
        s.addText(item, {
          x: gx + 0.55, y: iy, w: 2.2, h: 0.32,
          fontSize: 11, fontFace: "Microsoft JhengHei", color: C.text, align: "left", valign: "middle",
          margin: 0,
        });
      });
    });

    // Bottom: market info
    addCardWithAccent(s, 0.7, 4.6, 8.8, 0.85, C.navy);
    s.addText("市場與客戶", {
      x: 1.0, y: 4.65, w: 3, h: 0.3,
      fontSize: 14, fontFace: "Microsoft JhengHei", color: C.navy, bold: true, align: "left", valign: "middle",
      margin: 0,
    });

    const marketItems = [
      "台灣/中國/日本市場  |  電腦週邊、電子元件、OEM/ODM  |  通路：批發、零售、網路銷售",
    ];
    marketItems.forEach((item) => {
      s.addText(item, {
        x: 1.0, y: 4.95, w: 8.3, h: 0.35,
        fontSize: 10, fontFace: "Microsoft JhengHei", color: C.muted, align: "left", valign: "middle",
        margin: 0,
      });
    });
  });

  // ════════════════════════════════════════════
  // SLIDE 6: 財務會計架構與鼎新ERP
  // ════════════════════════════════════════════
  addContentSlide("財務會計架構 與 鼎新ERP導入", "Financial Accounting & Dingxin ERP Cosmos", (s) => {
    // Left: 主辦會計職掌
    addCardWithAccent(s, 0.7, 1.4, 4.2, 3.6, C.navy);
    s.addText("主辦會計職掌", {
      x: 1.0, y: 1.55, w: 3.6, h: 0.4,
      fontSize: 18, fontFace: "Microsoft JhengHei", color: C.navy, bold: true, align: "left", valign: "middle",
      margin: 0,
    });
    s.addShape(pptx.ShapeType.rect, { x: 1.0, y: 1.95, w: 1.2, h: 0.03, fill: { color: C.navy } });

    const duties = [
      "總帳管理與傳票編製",
      "營業稅申報（401 調節表）",
      "各類所得扣繳申報",
      "營所稅結算申報",
      "成本分析與毛利率計算",
      "財務報表編製（損益表/資產負債表）",
      "配合會計師查帳作業",
    ];
    duties.forEach((d, i) => {
      const dy = 2.15 + i * 0.36;
      s.addShape(pptx.ShapeType.ellipse, { x: 1.15, y: dy + 0.07, w: 0.08, h: 0.08, fill: { color: C.navy } });
      s.addText(d, {
        x: 1.4, y: dy, w: 3.3, h: 0.3,
        fontSize: 11, fontFace: "Microsoft JhengHei", color: C.text, align: "left", valign: "middle",
        margin: 0,
      });
    });

    // Right: 鼎新ERP模組
    addCardWithAccent(s, 5.3, 1.4, 4.2, 3.6, C.teal);
    s.addText("鼎新 ERP Cosmos 模組", {
      x: 5.6, y: 1.55, w: 3.6, h: 0.4,
      fontSize: 18, fontFace: "Microsoft JhengHei", color: C.teal, bold: true, align: "left", valign: "middle",
      margin: 0,
    });
    s.addShape(pptx.ShapeType.rect, { x: 5.6, y: 1.95, w: 1.2, h: 0.03, fill: { color: C.teal } });

    const modules = [
      ["總帳模組 (GL)", "傳票作業、科目管理、財務報表"],
      ["應收/應付 (AR/AP)", "帳款管理、沖帳作業、帳齡分析"],
      ["固定資產 (FA)", "資產登錄、折舊計算、盤點管理"],
      ["票據管理 (PM)", "票據收付、兌現管理、調整作業"],
      ["成本計算 (CC)", "成本結轉、毛利分析、部門損益"],
      ["稅務管理 (TAX)", "營業稅申報、401調節表、扣繳申報"],
    ];
    modules.forEach((m, i) => {
      const my = 2.15 + i * 0.44;
      s.addShape(pptx.ShapeType.roundRect, {
        x: 5.7, y: my, w: 0.28, h: 0.28,
        fill: { color: C.teal },
        rectRadius: 0.04,
      });
      s.addText(String(i + 1), {
        x: 5.7, y: my, w: 0.28, h: 0.28,
        fontSize: 10, fontFace: "Arial", color: C.white, bold: true, align: "center", valign: "middle",
        margin: 0,
      });
      s.addText(m[0], {
        x: 6.15, y: my, w: 3.1, h: 0.18,
        fontSize: 11, fontFace: "Microsoft JhengHei", color: C.text, bold: true, align: "left", valign: "middle",
        margin: 0,
      });
      s.addText(m[1], {
        x: 6.15, y: my + 0.18, w: 3.1, h: 0.2,
        fontSize: 9, fontFace: "Microsoft JhengHei", color: C.muted, align: "left", valign: "middle",
        margin: 0,
      });
    });
  });

  // ════════════════════════════════════════════
  // SLIDE 7: 總帳管理
  // ════════════════════════════════════════════
  addContentSlide("總帳管理 (GL)", "傳票編製 → 過帳 → 科目餘額 → 財務報表", (s) => {
    // Process flow - horizontal steps
    const steps = [
      ["傳票輸入", "審核憑證\n編製傳票\n分錄輸入"],
      ["過帳作業", "批次過帳\n科目更新\n期間控制"],
      ["科目餘額", "科目餘額表\n試算表\n調節核對"],
      ["財務報表", "損益表\n資產負債表\n現金流量表"],
    ];

    steps.forEach((step, i) => {
      const sx = 0.7 + i * 2.35;

      // Step card
      s.addShape(pptx.ShapeType.roundRect, {
        x: sx, y: 1.4, w: 2.15, h: 2.6,
        fill: { color: C.cardBg },
        rectRadius: 0.08,
        line: { color: C.border, width: 0.5 },
      });

      // Step number
      s.addShape(pptx.ShapeType.roundRect, {
        x: sx + 0.3, y: 1.55, w: 1.55, h: 0.4,
        fill: { color: C.navy },
        rectRadius: 0.05,
      });
      s.addText(`Step ${i + 1}`, {
        x: sx + 0.3, y: 1.55, w: 1.55, h: 0.4,
        fontSize: 12, fontFace: "Arial", color: C.white, bold: true, align: "center", valign: "middle",
        margin: 0,
      });

      // Step title
      s.addText(step[0], {
        x: sx + 0.15, y: 2.1, w: 1.85, h: 0.35,
        fontSize: 14, fontFace: "Microsoft JhengHei", color: C.navy, bold: true, align: "center", valign: "middle",
        margin: 0,
      });
      s.addShape(pptx.ShapeType.rect, { x: sx + 0.4, y: 2.45, w: 1.35, h: 0.02, fill: { color: C.teal } });

      // Step details
      s.addText(step[1], {
        x: sx + 0.2, y: 2.6, w: 1.75, h: 1.2,
        fontSize: 11, fontFace: "Microsoft JhengHei", color: C.muted, align: "center", valign: "top",
        margin: 0,
        lineSpacingMultiple: 1.3,
      });

      // Arrow between steps
      if (i < steps.length - 1) {
        s.addText("\u25B6", {
          x: sx + 2.15, y: 2.3, w: 0.2, h: 0.5,
          fontSize: 16, fontFace: "Arial", color: C.teal, align: "center", valign: "middle",
          margin: 0,
        });
      }
    });

    // Bottom: Key features
    addCardWithAccent(s, 0.7, 4.2, 8.8, 1.2, C.teal);
    s.addText("總帳模組關鍵功能", {
      x: 1.0, y: 4.3, w: 3, h: 0.3,
      fontSize: 14, fontFace: "Microsoft JhengHei", color: C.navy, bold: true, align: "left", valign: "middle",
      margin: 0,
    });

    const glFeatures = [
      "會計科目彈性設定（部門/專案輔助核算）  |  沖帳/立沖帳管理  |  多期間比較報表  |  年度結轉與開帳作業",
    ];
    glFeatures.forEach((f) => {
      s.addText(f, {
        x: 1.0, y: 4.65, w: 8.3, h: 0.35,
        fontSize: 10, fontFace: "Microsoft JhengHei", color: C.muted, align: "left", valign: "middle",
        margin: 0,
      });
    });

    s.addText("產出報表：資產負債表、損益表、現金流量表、科目餘額表、試算表、日記帳", {
      x: 1.0, y: 4.95, w: 8.3, h: 0.35,
      fontSize: 10, fontFace: "Microsoft JhengHei", color: C.muted, align: "left", valign: "middle",
      margin: 0,
    });
  });

  // ════════════════════════════════════════════
  // SLIDE 8: 營業稅申報
  // ════════════════════════════════════════════
  addContentSlide("營業稅申報 (401)", "加值型營業稅申報實務流程", (s) => {
    // Left: 流程
    addCardWithAccent(s, 0.7, 1.4, 4.2, 3.6, C.amber);
    s.addText("申報流程", {
      x: 1.0, y: 1.55, w: 3.6, h: 0.4,
      fontSize: 18, fontFace: "Microsoft JhengHei", color: C.amber, bold: true, align: "left", valign: "middle",
      margin: 0,
    });
    s.addShape(pptx.ShapeType.rect, { x: 1.0, y: 1.95, w: 1.2, h: 0.03, fill: { color: C.amber } });

    const taxSteps = [
      ["01", "憑證整理", "進項/銷項發票分類整理"],
      ["02", "系統登錄", "鼎新ERP 傳票輸入"],
      ["03", "媒體申報", "匯出媒體檔申報"],
      ["04", "401 調節表", "核對申報金額與帳載差異"],
      ["05", "申報送件", "網路申報 / 媒體申報"],
      ["06", "繳納稅款", "核對繳款書與入帳"],
    ];
    taxSteps.forEach((ts, i) => {
      const ty = 2.15 + i * 0.43;
      s.addShape(pptx.ShapeType.roundRect, {
        x: 1.1, y: ty, w: 0.3, h: 0.3,
        fill: { color: C.amber },
        rectRadius: 0.04,
      });
      s.addText(ts[0], {
        x: 1.1, y: ty, w: 0.3, h: 0.3,
        fontSize: 10, fontFace: "Arial", color: C.white, bold: true, align: "center", valign: "middle",
        margin: 0,
      });
      s.addText(ts[1], {
        x: 1.55, y: ty, w: 1.5, h: 0.16,
        fontSize: 12, fontFace: "Microsoft JhengHei", color: C.text, bold: true, align: "left", valign: "middle",
        margin: 0,
      });
      s.addText(ts[2], {
        x: 1.55, y: ty + 0.16, w: 3.1, h: 0.15,
        fontSize: 10, fontFace: "Microsoft JhengHei", color: C.muted, align: "left", valign: "middle",
        margin: 0,
      });
      if (i < taxSteps.length - 1) {
        s.addShape(pptx.ShapeType.rect, { x: 1.24, y: ty + 0.32, w: 0.02, h: 0.1, fill: { color: C.border } });
      }
    });

    // Right: 要點
    addCardWithAccent(s, 5.3, 1.4, 4.2, 3.6, C.navy);
    s.addText("申報要點", {
      x: 5.6, y: 1.55, w: 3.6, h: 0.4,
      fontSize: 18, fontFace: "Microsoft JhengHei", color: C.navy, bold: true, align: "left", valign: "middle",
      margin: 0,
    });
    s.addShape(pptx.ShapeType.rect, { x: 5.6, y: 1.95, w: 1.2, h: 0.03, fill: { color: C.navy } });

    const taxPoints = [
      "401 調節表核對申報數與帳載數差異",
      "進項稅額扣抵（可扣抵/不可扣抵）",
      "零稅率/免稅銷售額申報",
      "固定資產進項稅額扣抵",
      "兼營營業人比例扣抵法",
      "申報期限：每單月 15 日前",
    ];
    taxPoints.forEach((tp, i) => {
      const ty = 2.15 + i * 0.38;
      s.addShape(pptx.ShapeType.ellipse, { x: 5.75, y: ty + 0.07, w: 0.08, h: 0.08, fill: { color: C.navy } });
      s.addText(tp, {
        x: 6.0, y: ty, w: 3.3, h: 0.32,
        fontSize: 11, fontFace: "Microsoft JhengHei", color: C.text, align: "left", valign: "middle",
        margin: 0,
      });
    });
  });

  // ════════════════════════════════════════════
  // SLIDE 9: 各類所得扣繳與營所稅
  // ════════════════════════════════════════════
  addContentSlide("各類所得扣繳 與 營所稅申報", "Tax Compliance & Reporting", (s) => {
    // Card 1: 各類所得扣繳
    addCardWithAccent(s, 0.7, 1.4, 4.2, 2.5, C.navy);
    s.addText("各類所得扣繳申報", {
      x: 1.0, y: 1.55, w: 3.6, h: 0.4,
      fontSize: 17, fontFace: "Microsoft JhengHei", color: C.navy, bold: true, align: "left", valign: "middle",
      margin: 0,
    });
    s.addShape(pptx.ShapeType.rect, { x: 1.0, y: 1.95, w: 1.2, h: 0.03, fill: { color: C.navy } });

    const withholdingItems = [
      "薪資所得扣繳（依薪資扣繳辦法）",
      "租金/執行業務所得扣繳",
      "股利/利息所得扣繳",
      "各類所得調節表編製",
      "申報時間：每年 1 月底前（上年度）",
      "熟悉各類所得扣繳率標準",
    ];
    withholdingItems.forEach((w, i) => {
      const wy = 2.15 + i * 0.28;
      s.addShape(pptx.ShapeType.ellipse, { x: 1.15, y: wy + 0.05, w: 0.07, h: 0.07, fill: { color: C.navy } });
      s.addText(w, {
        x: 1.4, y: wy, w: 3.3, h: 0.25,
        fontSize: 10, fontFace: "Microsoft JhengHei", color: C.text, align: "left", valign: "middle",
        margin: 0,
      });
    });

    // Card 2: 營所稅
    addCardWithAccent(s, 5.3, 1.4, 4.2, 2.5, C.teal);
    s.addText("營利事業所得稅申報", {
      x: 5.6, y: 1.55, w: 3.6, h: 0.4,
      fontSize: 17, fontFace: "Microsoft JhengHei", color: C.teal, bold: true, align: "left", valign: "middle",
      margin: 0,
    });
    s.addShape(pptx.ShapeType.rect, { x: 5.6, y: 1.95, w: 1.2, h: 0.03, fill: { color: C.teal } });

    const profitItems = [
      "年度結算申報（每年 5 月）",
      "財務報表調節為稅務報表",
      "所得稅費用估列與暫繳申報",
      "營利事業所得稅查核準則應用",
      "未分配盈餘申報",
      "配合會計師稅務簽證查帳",
    ];
    profitItems.forEach((p, i) => {
      const py = 2.15 + i * 0.28;
      s.addShape(pptx.ShapeType.ellipse, { x: 5.75, y: py + 0.05, w: 0.07, h: 0.07, fill: { color: C.teal } });
      s.addText(p, {
        x: 6.0, y: py, w: 3.3, h: 0.25,
        fontSize: 10, fontFace: "Microsoft JhengHei", color: C.text, align: "left", valign: "middle",
        margin: 0,
      });
    });

    // Bottom: 配合會計師查帳
    addCardWithAccent(s, 0.7, 4.1, 8.8, 1.3, C.amber);
    s.addText("配合會計師查帳作業", {
      x: 1.0, y: 4.2, w: 3, h: 0.35,
      fontSize: 15, fontFace: "Microsoft JhengHei", color: C.navy, bold: true, align: "left", valign: "middle",
      margin: 0,
    });

    const auditItems = [
      "備妥傳票憑證、科目餘額表、調節表等查核資料",
      "協助盤點作業（固定資產/庫存）",
      "提供關係人交易明細及相關合約",
      "回覆會計師查核問題，確保稅報順利出具",
    ];
    auditItems.forEach((a, i) => {
      const ay = 4.6 + i * 0.2;
      s.addText(`\u25B6 ${a}`, {
        x: 1.0, y: ay, w: 8.3, h: 0.2,
        fontSize: 9, fontFace: "Microsoft JhengHei", color: C.muted, align: "left", valign: "middle",
        margin: 0,
      });
    });
  });

  // ════════════════════════════════════════════
  // SLIDE 10: 成本分析
  // ════════════════════════════════════════════
  addContentSlide("成本分析", "成本結轉流程 與 毛利分析", (s) => {
    // Left: 成本結轉流程
    addCardWithAccent(s, 0.7, 1.4, 4.2, 3.6, C.teal);
    s.addText("鼎新ERP 成本結轉流程", {
      x: 1.0, y: 1.55, w: 3.6, h: 0.4,
      fontSize: 17, fontFace: "Microsoft JhengHei", color: C.teal, bold: true, align: "left", valign: "middle",
      margin: 0,
    });
    s.addShape(pptx.ShapeType.rect, { x: 1.0, y: 1.95, w: 1.2, h: 0.03, fill: { color: C.teal } });

    const costSteps = [
      ["進貨/領料", "原料採購入庫  生產領料出庫"],
      ["人工/費用", "直接人工投入  製造費用分攤"],
      ["成本計算", "批次成本計算  在製/製成品結轉"],
      ["毛利分析", "銷貨成本結轉  毛利率分析"],
    ];
    costSteps.forEach((cs, i) => {
      const cy = 2.15 + i * 0.6;
      s.addShape(pptx.ShapeType.roundRect, {
        x: 1.1, y: cy, w: 0.28, h: 0.28,
        fill: { color: C.teal },
        rectRadius: 0.04,
      });
      s.addText(String(i + 1), {
        x: 1.1, y: cy, w: 0.28, h: 0.28,
        fontSize: 11, fontFace: "Arial", color: C.white, bold: true, align: "center", valign: "middle",
        margin: 0,
      });
      s.addText(cs[0], {
        x: 1.55, y: cy, w: 1.8, h: 0.18,
        fontSize: 12, fontFace: "Microsoft JhengHei", color: C.text, bold: true, align: "left", valign: "middle",
        margin: 0,
      });
      s.addText(cs[1], {
        x: 1.55, y: cy + 0.18, w: 3.1, h: 0.35,
        fontSize: 10, fontFace: "Microsoft JhengHei", color: C.muted, align: "left", valign: "top",
        margin: 0,
      });
      if (i < costSteps.length - 1) {
        s.addShape(pptx.ShapeType.rect, { x: 1.23, y: cy + 0.3, w: 0.02, h: 0.28, fill: { color: C.border } });
      }
    });

    // Right: 分析項目
    addCardWithAccent(s, 5.3, 1.4, 4.2, 3.6, C.amber);
    s.addText("成本分析項目", {
      x: 5.6, y: 1.55, w: 3.6, h: 0.4,
      fontSize: 17, fontFace: "Microsoft JhengHei", color: C.amber, bold: true, align: "left", valign: "middle",
      margin: 0,
    });
    s.addShape(pptx.ShapeType.rect, { x: 5.6, y: 1.95, w: 1.2, h: 0.03, fill: { color: C.amber } });

    const analysisItems = [
      ["毛利率分析", "各產品/專案毛利率比較"],
      ["部門損益", "部門別損益計算與分析"],
      ["成本差異", "標準成本 vs 實際成本差異"],
      ["費用分攤", "製造費用合理分攤基礎"],
      ["損益兩平", "損益兩平點與安全邊際"],
      ["趨勢分析", "跨期比較與異常預警"],
    ];
    analysisItems.forEach((a, i) => {
      const ay = 2.15 + i * 0.42;
      s.addShape(pptx.ShapeType.roundRect, {
        x: 5.7, y: ay, w: 0.28, h: 0.28,
        fill: { color: C.amber },
        rectRadius: 0.04,
      });
      s.addText(String(i + 1), {
        x: 5.7, y: ay, w: 0.28, h: 0.28,
        fontSize: 10, fontFace: "Arial", color: C.white, bold: true, align: "center", valign: "middle",
        margin: 0,
      });
      s.addText(a[0], {
        x: 6.15, y: ay, w: 2, h: 0.18,
        fontSize: 12, fontFace: "Microsoft JhengHei", color: C.text, bold: true, align: "left", valign: "middle",
        margin: 0,
      });
      s.addText(a[1], {
        x: 6.15, y: ay + 0.18, w: 3.1, h: 0.18,
        fontSize: 10, fontFace: "Microsoft JhengHei", color: C.muted, align: "left", valign: "middle",
        margin: 0,
      });
    });
  });

  // ════════════════════════════════════════════
  // SLIDE 11: ERP導入經驗與效益
  // ════════════════════════════════════════════
  addContentSlide("ERP導入經驗 與 資訊應用效益", "實戰經驗分享", (s) => {
    // Top: Timeline cards
    const expCards = [
      {
        title: "ERP/POS 系統導入",
        desc: "參與餐飲業 ERP、POS、總帳系統導入專案，完成重新開帳、會計科目設定、部門設定等作業",
        tag: "系統導入",
        color: C.navy,
      },
      {
        title: "內控制度建置",
        desc: "協助創櫃公司建置內部控制制度（採購付款循環），完善傳票、憑證、用印作業流程",
        tag: "內控稽核",
        color: C.teal,
      },
      {
        title: "VBA 自動化優化",
        desc: "運用 AI 撰寫 VBA 自動化 Excel 流程，提升科目餘額表編製效率，降低人工整理成本",
        tag: "流程優化",
        color: C.amber,
      },
    ];

    expCards.forEach((ec, i) => {
      const ex = 0.7 + i * 3.15;
      addCardWithAccent(s, ex, 1.4, 2.95, 2.0, ec.color);

      // Tag badge
      s.addShape(pptx.ShapeType.roundRect, {
        x: ex + 0.2, y: 1.55, w: 1.0, h: 0.25,
        fill: { color: ec.color },
        rectRadius: 0.04,
      });
      s.addText(ec.tag, {
        x: ex + 0.2, y: 1.55, w: 1.0, h: 0.25,
        fontSize: 9, fontFace: "Microsoft JhengHei", color: C.white, bold: true, align: "center", valign: "middle",
        margin: 0,
      });

      s.addText(ec.title, {
        x: ex + 0.2, y: 1.9, w: 2.55, h: 0.3,
        fontSize: 14, fontFace: "Microsoft JhengHei", color: C.text, bold: true, align: "left", valign: "middle",
        margin: 0,
      });
      s.addText(ec.desc, {
        x: ex + 0.2, y: 2.25, w: 2.55, h: 0.95,
        fontSize: 10, fontFace: "Microsoft JhengHei", color: C.muted, align: "left", valign: "top",
        margin: 0,
      });
    });

    // Bottom: Key advantages
    addCardWithAccent(s, 0.7, 3.6, 8.8, 1.8, C.navy);
    s.addText("數位化會計核心優勢", {
      x: 1.0, y: 3.7, w: 4, h: 0.35,
      fontSize: 15, fontFace: "Microsoft JhengHei", color: C.navy, bold: true, align: "left", valign: "middle",
      margin: 0,
    });

    const advItems = [
      ["效率提升", "ERP 自動化結帳流程，縮短月結時間 50% 以上"],
      ["數據整合", "鼎新串接各模組數據，減少人工轉換錯誤"],
      ["決策支援", "即時財務報表與成本分析，提供管理層決策依據"],
      ["合規透明", "內控流程系統化，確保稅務申報與會計準則合規"],
    ];
    advItems.forEach((adv, i) => {
      const ax = 1.0 + (i % 2) * 4.2;
      const ay = 4.15 + Math.floor(i / 2) * 0.55;

      s.addShape(pptx.ShapeType.roundRect, {
        x: ax, y: ay, w: 0.22, h: 0.22,
        fill: { color: C.navy },
        rectRadius: 0.03,
      });
      s.addText(String(i + 1), {
        x: ax, y: ay, w: 0.22, h: 0.22,
        fontSize: 9, fontFace: "Arial", color: C.white, bold: true, align: "center", valign: "middle",
        margin: 0,
      });
      s.addText(adv[0], {
        x: ax + 0.35, y: ay - 0.02, w: 1.3, h: 0.25,
        fontSize: 11, fontFace: "Microsoft JhengHei", color: C.text, bold: true, align: "left", valign: "middle",
        margin: 0,
      });
      s.addText(adv[1], {
        x: ax + 1.7, y: ay - 0.02, w: 2.5, h: 0.25,
        fontSize: 10, fontFace: "Microsoft JhengHei", color: C.muted, align: "left", valign: "middle",
        margin: 0,
      });
    });
  });

  // ════════════════════════════════════════════
  // SLIDE 12: 結尾
  // ════════════════════════════════════════════
  const s12 = pptx.addSlide();
  s12.background = { color: C.darkNavy };

  // Decorative circles
  s12.addShape(pptx.ShapeType.ellipse, { x: -2, y: -1.5, w: 4.5, h: 4.5, fill: { color: C.navy, transparency: 40 } });
  s12.addShape(pptx.ShapeType.ellipse, { x: 8, y: 3.5, w: 4, h: 4, fill: { color: C.navy, transparency: 50 } });

  // Top and bottom accent lines
  s12.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.teal } });
  s12.addShape(pptx.ShapeType.rect, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.teal } });

  s12.addText("THANK YOU", {
    x: 0.8, y: 0.4, w: 8.4, h: 0.5,
    fontSize: 14, fontFace: "Arial", color: C.teal, bold: true, align: "left", valign: "middle",
    charSpacing: 6,
    margin: 0,
  });

  s12.addText("感謝您的聆聽", {
    x: 0.8, y: 1.5, w: 8.4, h: 1.0,
    fontSize: 42, fontFace: "Microsoft JhengHei", color: C.white, bold: true, align: "left", valign: "middle",
    margin: 0,
  });

  s12.addShape(pptx.ShapeType.rect, { x: 0.8, y: 2.65, w: 2.5, h: 0.04, fill: { color: C.amber } });

  s12.addText("期待有機會加入永吉電腦\n貢獻財務會計專業與 ERP 實務經驗", {
    x: 0.8, y: 3.0, w: 8.4, h: 0.9,
    fontSize: 18, fontFace: "Microsoft JhengHei", color: C.white, align: "left", valign: "middle",
    margin: 0,
    lineSpacingMultiple: 1.3,
  });

  // Contact info
  s12.addShape(pptx.ShapeType.roundRect, {
    x: 0.8, y: 4.2, w: 5.5, h: 1.1,
    fill: { color: C.navy, transparency: 30 },
    rectRadius: 0.08,
  });

  const contactLines = [
    { text: "許博宇  Paul Hsu", opts: { fontSize: 16, bold: true, color: C.white, breakLine: true } },
    { text: "paul830305@gmail.com  |  0975-470-790", opts: { fontSize: 12, color: C.teal, breakLine: true } },
    { text: "台北市中山區  |  會計資訊系畢業", opts: { fontSize: 11, color: C.lightMuted } },
  ];
  s12.addText(contactLines, {
    x: 1.1, y: 4.3, w: 5, h: 0.9,
    fontFace: "Microsoft JhengHei", align: "left", valign: "middle",
    margin: 0,
  });

  // ════════════════════════════════════════════
  // SAVE
  // ════════════════════════════════════════════
  const outputPath = "YongJi_Company.pptx";
  await pptx.writeFile({ fileName: outputPath });
  console.log(`\u2705 \u751f\u6210\u6210\u529f: ${outputPath}`);
}

generate().catch(console.error);
