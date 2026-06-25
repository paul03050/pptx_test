const pptxgen = require("pptxgenjs");

async function generate() {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "創思達科技";
  pptx.title = "創思達科技 - IPO準備、內控制度建立與ERP導入規劃";

  const C = {
    navy: "0D4A7A",
    teal: "028090",
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
    red: "E53E3E",
    purple: "6B46C1",
    green: "38A169",
  };

  function addContentSlide(title, subtitle, cb) {
    const s = pptx.addSlide();
    s.background = { color: C.lightBg };
    s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.35, h: 7.5, fill: { color: C.navy } });
    s.addShape(pptx.ShapeType.rect, { x: 0.35, y: 0, w: 12.98, h: 0.04, fill: { color: C.teal } });
    s.addText(title, {
      x: 0.7, y: 0.25, w: 11, h: 0.55,
      fontSize: 28, fontFace: "Microsoft JhengHei", color: C.navy, bold: true, align: "left", valign: "middle", margin: 0,
    });
    if (subtitle) {
      s.addText(subtitle, {
        x: 0.7, y: 0.75, w: 11, h: 0.35,
        fontSize: 14, fontFace: "Microsoft JhengHei", color: C.muted, align: "left", valign: "middle", margin: 0,
      });
    }
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

  function addCardAccent(slide, x, y, w, h, accentColor) {
    addCard(slide, x, y, w, h);
    slide.addShape(pptx.ShapeType.rect, { x, y, w: 0.06, h, fill: { color: accentColor || C.teal } });
  }

  function bulletItem(text, opts) {
    return { text, options: { bullet: true, breakLine: true, ...opts } };
  }

  // ════════════════════════════════════════════
  // SLIDE 1: 封面
  // ════════════════════════════════════════════
  const s1 = pptx.addSlide();
  s1.background = { color: C.darkNavy };

  s1.addShape(pptx.ShapeType.ellipse, { x: 11, y: -2, w: 5, h: 5, fill: { color: C.navy, transparency: 40 } });
  s1.addShape(pptx.ShapeType.ellipse, { x: -2, y: 4.5, w: 5.5, h: 5.5, fill: { color: C.navy, transparency: 50 } });
  s1.addShape(pptx.ShapeType.ellipse, { x: 9.5, y: 5, w: 3, h: 3, fill: { color: C.teal, transparency: 60 } });

  s1.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.33, h: 0.06, fill: { color: C.teal } });
  s1.addShape(pptx.ShapeType.rect, { x: 0, y: 7.44, w: 13.33, h: 0.06, fill: { color: C.teal } });

  s1.addText("CSD TECH", {
    x: 1.2, y: 1.8, w: 10.93, h: 0.5,
    fontSize: 14, fontFace: "Arial", color: C.teal, bold: true, align: "left", valign: "middle",
    charSpacing: 6, margin: 0,
  });

  s1.addText("創思達科技股份有限公司", {
    x: 1.2, y: 2.4, w: 10.93, h: 1.2,
    fontSize: 44, fontFace: "Microsoft JhengHei", color: C.white, bold: true, align: "left", valign: "middle", margin: 0,
  });

  s1.addText("IPO 準備 · 內控制度建立 · ERP 導入規劃", {
    x: 1.2, y: 3.6, w: 10.93, h: 0.7,
    fontSize: 24, fontFace: "Microsoft JhengHei", color: C.teal, align: "left", valign: "middle", margin: 0,
  });

  s1.addShape(pptx.ShapeType.rect, { x: 1.2, y: 4.5, w: 2.5, h: 0.04, fill: { color: C.amber } });

  s1.addText("財務會計部  |  專案規劃簡報", {
    x: 1.2, y: 4.8, w: 10.93, h: 0.4,
    fontSize: 16, fontFace: "Microsoft JhengHei", color: C.white, align: "left", valign: "middle", margin: 0,
  });
  s1.addText("2026 年度", {
    x: 1.2, y: 5.3, w: 10.93, h: 0.35,
    fontSize: 13, fontFace: "Microsoft JhengHei", color: C.lightMuted, align: "left", valign: "middle", margin: 0,
  });

  // ════════════════════════════════════════════
  // SLIDE 2: 目錄
  // ════════════════════════════════════════════
  addContentSlide("議 程 內 容", "AGENDA", (s) => {
    const items = [
      ["01", "公司現況與職務範疇", "帳務處理、稅務申報、存貨管理、ERP評估、專案控管"],
      ["02", "會計作業優化與憑證制度", "流程梳理、憑證標準化、內控優化"],
      ["03", "存貨管理與成本管控", "製造研發料件管理、成本結轉與分析"],
      ["04", "財務報表與分析體系", "管理報表架構、經營決策KPI"],
      ["05", "IPO 準備路徑圖", "上櫃/上市條件檢核、輔導團隊、時程規劃"],
      ["06", "內部控制制度建立", "COSO 架構、五大循環、控制活動設計"],
      ["07", "ERP 系統評估與導入", "需求分析、方案比較、導入步驟"],
      ["08", "專案流程規劃與控管", "WBS、里程碑、溝通與風險管理"],
    ];

    items.forEach((item, i) => {
      const row = Math.floor(i / 2);
      const col = i % 2;
      const x = 0.8 + col * 6.2;
      const y = 1.4 + row * 1.4;

      addCardAccent(s, x, y, 5.8, 1.15, C.navy);
      s.addText(item[0], {
        x: x + 0.25, y: y + 0.05, w: 0.7, h: 1.05,
        fontSize: 26, fontFace: "Arial", color: C.navy, bold: true, align: "center", valign: "middle", margin: 0,
      });
      s.addText(item[1], {
        x: x + 1.0, y: y + 0.1, w: 4.5, h: 0.45,
        fontSize: 18, fontFace: "Microsoft JhengHei", color: C.text, bold: true, align: "left", valign: "middle", margin: 0,
      });
      s.addText(item[2], {
        x: x + 1.0, y: y + 0.55, w: 4.5, h: 0.5,
        fontSize: 13, fontFace: "Microsoft JhengHei", color: C.muted, align: "left", valign: "top", margin: 0,
      });
    });
  });

  // ════════════════════════════════════════════
  // SLIDE 3: 公司現況與職務範疇
  // ════════════════════════════════════════════
  addContentSlide("公司現況與職務範疇", "創思達科技 · 五大核心職掌", (s) => {
    const duties = [
      { title: "帳務處理與稅務申報", items: ["公司帳務處理、傳票編製", "營業稅/各類所得/營所稅申報", "401 調節表編製", "會計憑證制度建立與優化"], color: C.navy },
      { title: "存貨與成本管控", items: ["製造及研發料件存貨管理", "成本管控流程建立", "存貨盤點制度設計", "料件分級與出入庫流程"], color: C.teal },
      { title: "財務報表與分析", items: ["定期編製財務報表", "經營分析與KPI指標", "預算與差異分析", "提供決策所需數據"], color: C.amber },
      { title: "ERP 評估與內控", items: ["ERP 系統評估導入", "內控制度編寫與落實", "流程系統化與自動化", "系統整合與資料治理"], color: C.purple },
    ];

    duties.forEach((d, i) => {
      const x = 0.5 + i * 3.2;
      addCardAccent(s, x, 1.4, 2.95, 5.2, d.color);
      s.addText(d.title, {
        x: x + 0.2, y: 1.6, w: 2.55, h: 0.5,
        fontSize: 17, fontFace: "Microsoft JhengHei", color: d.color, bold: true, align: "left", valign: "middle", margin: 0,
      });
      s.addShape(pptx.ShapeType.rect, { x: x + 0.2, y: 2.1, w: 1.5, h: 0.03, fill: { color: d.color } });

      d.items.forEach((it, j) => {
        const iy = 2.3 + j * 0.65;
        s.addShape(pptx.ShapeType.ellipse, { x: x + 0.35, y: iy + 0.1, w: 0.08, h: 0.08, fill: { color: d.color } });
        s.addText(it, {
          x: x + 0.55, y: iy, w: 2.2, h: 0.5,
          fontSize: 12, fontFace: "Microsoft JhengHei", color: C.text, align: "left", valign: "top", margin: 0,
        });
      });
    });

    s.addShape(pptx.ShapeType.roundRect, {
      x: 0.5, y: 6.8, w: 12.33, h: 0.5,
      fill: { color: C.navy }, rectRadius: 0.05,
    });
    s.addText("核心目標：建立制度化財務會計體系，為 IPO 上市櫃做好準備", {
      x: 0.7, y: 6.8, w: 11.93, h: 0.5,
      fontSize: 14, fontFace: "Microsoft JhengHei", color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
    });
  });

  // ════════════════════════════════════════════
  // SLIDE 4: 會計作業現況與流程優化
  // ════════════════════════════════════════════
  addContentSlide("會計作業現況與流程優化", "帳務處理 · 稅務申報 · 憑證制度", (s) => {
    // Left: 現況分析
    addCardAccent(s, 0.7, 1.4, 5.8, 5.0, C.navy);
    s.addText("現行作業分析", {
      x: 1.0, y: 1.55, w: 5.2, h: 0.4,
      fontSize: 20, fontFace: "Microsoft JhengHei", color: C.navy, bold: true, align: "left", valign: "middle", margin: 0,
    });
    s.addShape(pptx.ShapeType.rect, { x: 1.0, y: 1.95, w: 1.2, h: 0.03, fill: { color: C.navy } });

    const currentIssues = [
      ["憑證管理", "憑證分類與保存制度待建立，缺少標準化作業流程"],
      ["記帳流程", "部分帳務仍仰賴人工Excel作業，缺乏系統自動化"],
      ["稅務申報", "申報時程與調節表編製缺乏制度化控管"],
      ["溝通協作", "跨部門憑證傳遞與審核流程不明確"],
    ];
    currentIssues.forEach((issue, i) => {
      const iy = 2.2 + i * 0.9;
      s.addShape(pptx.ShapeType.roundRect, {
        x: 1.1, y: iy, w: 0.32, h: 0.32,
        fill: { color: C.navy }, rectRadius: 0.04,
      });
      s.addText(String(i + 1), {
        x: 1.1, y: iy, w: 0.32, h: 0.32,
        fontSize: 12, fontFace: "Arial", color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
      });
      s.addText(issue[0], {
        x: 1.6, y: iy - 0.02, w: 4.7, h: 0.22,
        fontSize: 14, fontFace: "Microsoft JhengHei", color: C.text, bold: true, align: "left", valign: "middle", margin: 0,
      });
      s.addText(issue[1], {
        x: 1.6, y: iy + 0.2, w: 4.7, h: 0.55,
        fontSize: 12, fontFace: "Microsoft JhengHei", color: C.muted, align: "left", valign: "top", margin: 0,
      });
      if (i < currentIssues.length - 1) {
        s.addShape(pptx.ShapeType.rect, { x: 1.25, y: iy + 0.35, w: 0.02, h: 0.52, fill: { color: C.border } });
      }
    });

    // Right: 優化方案
    addCardAccent(s, 6.8, 1.4, 5.8, 5.0, C.teal);
    s.addText("優化行動方案", {
      x: 7.1, y: 1.55, w: 5.2, h: 0.4,
      fontSize: 20, fontFace: "Microsoft JhengHei", color: C.teal, bold: true, align: "left", valign: "middle", margin: 0,
    });
    s.addShape(pptx.ShapeType.rect, { x: 7.1, y: 1.95, w: 1.2, h: 0.03, fill: { color: C.teal } });

    const solutions = [
      ["憑證標準化", "制定憑證分類編碼規則、審核層級、保存年限，建立會計憑證管理辦法"],
      ["流程自動化", "導入ERP系統取代人工Excel作業，實現傳票輸入至報表產出一鍵流程"],
      ["稅務制度化", "建立申報行事曆、調節表標準格式、複核機制，確保準時正確申報"],
      ["溝通透明化", "建立跨部門憑證審核流程圖，明確各節點負責人與時限要求"],
    ];
    solutions.forEach((sol, i) => {
      const sy = 2.2 + i * 0.9;
      s.addShape(pptx.ShapeType.roundRect, {
        x: 7.2, y: sy, w: 0.32, h: 0.32,
        fill: { color: C.teal }, rectRadius: 0.04,
      });
      s.addText(String(i + 1), {
        x: 7.2, y: sy, w: 0.32, h: 0.32,
        fontSize: 12, fontFace: "Arial", color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
      });
      s.addText(sol[0], {
        x: 7.7, y: sy - 0.02, w: 4.6, h: 0.22,
        fontSize: 14, fontFace: "Microsoft JhengHei", color: C.text, bold: true, align: "left", valign: "middle", margin: 0,
      });
      s.addText(sol[1], {
        x: 7.7, y: sy + 0.2, w: 4.6, h: 0.55,
        fontSize: 12, fontFace: "Microsoft JhengHei", color: C.muted, align: "left", valign: "top", margin: 0,
      });
      if (i < solutions.length - 1) {
        s.addShape(pptx.ShapeType.rect, { x: 7.35, y: sy + 0.35, w: 0.02, h: 0.52, fill: { color: C.border } });
      }
    });
  });

  // ════════════════════════════════════════════
  // SLIDE 5: 會計憑證制度建立
  // ════════════════════════════════════════════
  addContentSlide("會計憑證制度建立", "憑證分類 · 審核流程 · 保存管理", (s) => {
    // 三大支柱
    const pillars = [
      { title: "憑證分類與編碼", items: ["原始憑證：發票、收據、合約", "記帳憑證：傳票（現收/現支/轉帳）", "編碼規則：YYYYMM-XXX-部門", "電子憑證與紙本憑證對應管理"], color: C.navy },
      { title: "審核流程與權責", items: ["經辦人 → 部門主管 → 會計審核 → 核准", "金額分級授權：小額/一般/大額", "異常憑證退件與追蹤機制", "審核時限：標準 3 工作天完成"], color: C.teal },
      { title: "保存與調閱管理", items: ["紙本憑證：專人歸檔、按序存放", "電子備份：雲端 + 地端雙重備份", "保存年限：至少 10 年（稅法規定）", "調閱申請制度：記錄調閱人、用途、歸還日"], color: C.amber },
    ];

    pillars.forEach((p, i) => {
      const x = 0.6 + i * 4.15;
      addCardAccent(s, x, 1.4, 3.85, 3.5, p.color);
      s.addText(p.title, {
        x: x + 0.2, y: 1.55, w: 3.45, h: 0.4,
        fontSize: 18, fontFace: "Microsoft JhengHei", color: p.color, bold: true, align: "left", valign: "middle", margin: 0,
      });
      s.addShape(pptx.ShapeType.rect, { x: x + 0.2, y: 1.95, w: 1.2, h: 0.03, fill: { color: p.color } });
      p.items.forEach((it, j) => {
        const iy = 2.15 + j * 0.55;
        s.addShape(pptx.ShapeType.ellipse, { x: x + 0.35, y: iy + 0.08, w: 0.07, h: 0.07, fill: { color: p.color } });
        s.addText(it, {
          x: x + 0.55, y: iy, w: 3.1, h: 0.4,
          fontSize: 12, fontFace: "Microsoft JhengHei", color: C.text, align: "left", valign: "top", margin: 0,
        });
      });
    });

    // Bottom callout
    s.addShape(pptx.ShapeType.roundRect, {
      x: 0.6, y: 5.2, w: 12.13, h: 2.0,
      fill: { color: C.cardBg },
      rectRadius: 0.1, line: { color: C.border, width: 0.5 },
    });
    s.addShape(pptx.ShapeType.rect, { x: 0.6, y: 5.2, w: 12.13, h: 0.06, fill: { color: C.navy } });

    s.addText("會計憑證管理辦法 — 制定要點", {
      x: 1.0, y: 5.4, w: 11.3, h: 0.4,
      fontSize: 18, fontFace: "Microsoft JhengHei", color: C.navy, bold: true, align: "left", valign: "middle", margin: 0,
    });

    const policyItems = [
      "總則：目的（確保憑證真實/合法/完整）、適用範圍、名詞定義",
      "憑證種類與格式：明確定義各類憑證應附文件與填寫規範",
      "審核權限：金額分層授權表、簽核流程圖、代理人制度",
      "歸檔與保存：紙本/電子歸檔方式、保存年限、銷毀程序",
      "內部稽核：定期抽查憑證、缺失改善追蹤、年度內稽計畫",
    ];
    policyItems.forEach((it, i) => {
      s.addText(it, {
        x: 1.0, y: 5.85 + i * 0.25, w: 11.3, h: 0.25,
        fontSize: 12, fontFace: "Microsoft JhengHei", color: C.text, align: "left", valign: "middle", margin: 0,
      });
    });
  });

  // ════════════════════════════════════════════
  // SLIDE 6: 存貨管理制度
  // ════════════════════════════════════════════
  addContentSlide("存貨管理制度", "製造及研發料件 · 出入庫 · 盤點", (s) => {
    const sections = [
      { title: "料件分類與編碼", items: ["原料/物料/在製品/製成品/研發樣品", "ABC 分級管理（依金額與重要性）", "統一分類編碼規則（料號系統）", "BOM 表建立與維護"], color: C.navy },
      { title: "出入庫管理流程", items: ["進貨驗收：點收 → 檢驗 → 入庫", "領料出庫：工單領料/研發領料", "退貨/報廢流程：檢驗 → 核准 → 處理", "出入庫單據即時登錄ERP系統"], color: C.teal },
      { title: "盤點制度", items: ["年度全面盤點：每年至少一次", "循環盤點：每月重點品項抽查", "盤點差異分析與調整作業", "盤點績效指標：盤差率 < 1%"], color: C.amber },
    ];

    sections.forEach((sec, i) => {
      const x = 0.6 + i * 4.15;
      addCardAccent(s, x, 1.4, 3.85, 3.2, sec.color);
      s.addText(sec.title, {
        x: x + 0.2, y: 1.55, w: 3.45, h: 0.4,
        fontSize: 18, fontFace: "Microsoft JhengHei", color: sec.color, bold: true, align: "left", valign: "middle", margin: 0,
      });
      s.addShape(pptx.ShapeType.rect, { x: x + 0.2, y: 1.95, w: 1.2, h: 0.03, fill: { color: sec.color } });
      sec.items.forEach((it, j) => {
        const iy = 2.15 + j * 0.5;
        s.addShape(pptx.ShapeType.ellipse, { x: x + 0.35, y: iy + 0.07, w: 0.07, h: 0.07, fill: { color: sec.color } });
        s.addText(it, {
          x: x + 0.55, y: iy, w: 3.1, h: 0.4,
          fontSize: 12, fontFace: "Microsoft JhengHei", color: C.text, align: "left", valign: "top", margin: 0,
        });
      });
    });

    // Bottom: 研發料件特殊管理
    s.addShape(pptx.ShapeType.roundRect, {
      x: 0.6, y: 4.9, w: 12.13, h: 2.2,
      fill: { color: C.cardBg },
      rectRadius: 0.1, line: { color: C.border, width: 0.5 },
    });
    s.addShape(pptx.ShapeType.rect, { x: 0.6, y: 4.9, w: 0.06, h: 2.2, fill: { color: C.purple } });

    s.addText("研發料件管理要點", {
      x: 1.0, y: 5.0, w: 11.3, h: 0.35,
      fontSize: 17, fontFace: "Microsoft JhengHei", color: C.purple, bold: true, align: "left", valign: "middle", margin: 0,
    });

    const rdItems = [
      "研發樣品與量產料件分倉管理，避免混淆",
      "研發領料以專案為單位，建立專案領料單與成本歸屬",
      "研發報廢料件需經研發主管 + 會計雙重核准",
      "試製/試產階段料件視同在製品管理，納入盤點範圍",
    ];
    rdItems.forEach((it, i) => {
      s.addText(it, {
        x: 1.0, y: 5.4 + i * 0.35, w: 11.3, h: 0.3,
        fontSize: 12, fontFace: "Microsoft JhengHei", color: C.text, align: "left", valign: "middle", margin: 0,
      });
    });
  });

  // ════════════════════════════════════════════
  // SLIDE 7: 成本管控流程
  // ════════════════════════════════════════════
  addContentSlide("成本管控流程", "成本要素 · 結轉流程 · 差異分析", (s) => {
    // Top: 成本要素
    addCardAccent(s, 0.7, 1.4, 5.8, 2.2, C.navy);
    s.addText("成本要素架構", {
      x: 1.0, y: 1.55, w: 5.2, h: 0.4,
      fontSize: 18, fontFace: "Microsoft JhengHei", color: C.navy, bold: true, align: "left", valign: "middle", margin: 0,
    });
    s.addShape(pptx.ShapeType.rect, { x: 1.0, y: 1.95, w: 1.2, h: 0.03, fill: { color: C.navy } });

    const costElements = [
      { label: "直接原料", desc: "BOM 表用料 → 依實際領料入帳" },
      { label: "直接人工", desc: "生產線人員薪資 → 依工時分攤" },
      { label: "製造費用", desc: "廠房/水電/折舊/間接人員 → 合理分攤基礎" },
    ];
    costElements.forEach((ce, i) => {
      const cy = 2.15 + i * 0.4;
      s.addText(ce.label, { x: 1.0, y: cy, w: 1.5, h: 0.3, fontSize: 12, fontFace: "Microsoft JhengHei", color: C.text, bold: true, align: "left", valign: "middle", margin: 0 });
      s.addText(ce.desc, { x: 2.5, y: cy, w: 3.8, h: 0.3, fontSize: 11, fontFace: "Microsoft JhengHei", color: C.muted, align: "left", valign: "middle", margin: 0 });
    });

    // Top right: 結轉流程
    addCardAccent(s, 6.8, 1.4, 5.8, 2.2, C.teal);
    s.addText("成本結轉流程", {
      x: 7.1, y: 1.55, w: 5.2, h: 0.4,
      fontSize: 18, fontFace: "Microsoft JhengHei", color: C.teal, bold: true, align: "left", valign: "middle", margin: 0,
    });
    s.addShape(pptx.ShapeType.rect, { x: 7.1, y: 1.95, w: 1.2, h: 0.03, fill: { color: C.teal } });

    const flowSteps = ["進貨/領料", "人工/費用", "成本計算", "結轉製成品", "銷貨成本"];
    flowSteps.forEach((step, i) => {
      const fx = 7.1 + i * 1.1;
      s.addShape(pptx.ShapeType.roundRect, {
        x: fx, y: 2.2, w: 0.9, h: 0.8,
        fill: { color: C.teal }, rectRadius: 0.06,
      });
      s.addText(step, {
        x: fx, y: 2.2, w: 0.9, h: 0.8,
        fontSize: 10, fontFace: "Microsoft JhengHei", color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
      });
      if (i < flowSteps.length - 1) {
        s.addText("\u25B6", { x: fx + 0.9, y: 2.35, w: 0.2, h: 0.5, fontSize: 14, color: C.teal, align: "center", valign: "middle", margin: 0 });
      }
    });

    // Bottom: 差異分析
    addCardAccent(s, 0.7, 3.9, 11.9, 3.2, C.amber);
    s.addText("成本差異分析", {
      x: 1.0, y: 4.0, w: 5, h: 0.4,
      fontSize: 18, fontFace: "Microsoft JhengHei", color: C.amber, bold: true, align: "left", valign: "middle", margin: 0,
    });
    s.addShape(pptx.ShapeType.rect, { x: 1.0, y: 4.4, w: 1.2, h: 0.03, fill: { color: C.amber } });

    const diffs = [
      ["材料差異", "實際用量 vs 標準用量\n實際價格 vs 標準價格", "每週產出用料差異報表"],
      ["人工差異", "實際工時 vs 標準工時\n實際工資率 vs 標準工資率", "按月計算人工效率指標"],
      ["費用差異", "實際製造費用 vs 預計分攤\n產能利用率分析", "差異 > 10% 需提出說明"],
    ];
    diffs.forEach((d, i) => {
      const dx = 1.0 + i * 3.9;
      s.addShape(pptx.ShapeType.roundRect, {
        x: dx, y: 4.6, w: 3.6, h: 2.2,
        fill: { color: C.lightBg },
        rectRadius: 0.08, line: { color: C.border, width: 0.5 },
      });
      s.addShape(pptx.ShapeType.rect, { x: dx, y: 4.6, w: 3.6, h: 0.4, fill: { color: C.amber } });
      s.addText(d[0], {
        x: dx + 0.2, y: 4.6, w: 3.2, h: 0.4,
        fontSize: 14, fontFace: "Microsoft JhengHei", color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
      });
      s.addText(d[1], {
        x: dx + 0.2, y: 5.15, w: 3.2, h: 0.9,
        fontSize: 11, fontFace: "Microsoft JhengHei", color: C.text, align: "left", valign: "top", margin: 0,
      });
      s.addText(`\u25B6 ${d[2]}`, {
        x: dx + 0.2, y: 6.15, w: 3.2, h: 0.5,
        fontSize: 11, fontFace: "Microsoft JhengHei", color: C.teal, bold: true, align: "left", valign: "middle", margin: 0,
      });
    });
  });

  // ════════════════════════════════════════════
  // SLIDE 8: 財務報表與分析體系
  // ════════════════════════════════════════════
  addContentSlide("財務報表與分析體系", "定期報表 · KPI 指標 · 經營決策", (s) => {
    // Left: 報表架構
    addCardAccent(s, 0.7, 1.4, 5.8, 5.2, C.navy);
    s.addText("財務報告架構", {
      x: 1.0, y: 1.55, w: 5.2, h: 0.4,
      fontSize: 18, fontFace: "Microsoft JhengHei", color: C.navy, bold: true, align: "left", valign: "middle", margin: 0,
    });
    s.addShape(pptx.ShapeType.rect, { x: 1.0, y: 1.95, w: 1.2, h: 0.03, fill: { color: C.navy } });

    const reports = [
      { freq: "每日", items: ["現金日報表", "銀行存款餘額表"] },
      { freq: "每月", items: ["損益表", "資產負債表", "科目餘額表", "費用明細表", "應收/應付帳齡表"] },
      { freq: "每季", items: ["現金流量表", "部門損益表", "預算執行差異分析"] },
      { freq: "每年", items: ["年度財務報表（會計師查核）", "營所稅結算申報書"] },
    ];
    reports.forEach((r, i) => {
      const ry = 2.15 + i * 1.05;
      s.addShape(pptx.ShapeType.roundRect, {
        x: 1.1, y: ry, w: 0.55, h: 0.3,
        fill: { color: C.teal }, rectRadius: 0.04,
      });
      s.addText(r.freq, {
        x: 1.1, y: ry, w: 0.55, h: 0.3,
        fontSize: 10, fontFace: "Microsoft JhengHei", color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
      });
      r.items.forEach((it, j) => {
        s.addText(it, {
          x: 1.8, y: ry + j * 0.22, w: 4.5, h: 0.22,
          fontSize: 11, fontFace: "Microsoft JhengHei", color: C.text, align: "left", valign: "middle", margin: 0,
        });
      });
    });

    // Right: KPI 體系
    addCardAccent(s, 6.8, 1.4, 5.8, 5.2, C.teal);
    s.addText("經營分析 KPI 體系", {
      x: 7.1, y: 1.55, w: 5.2, h: 0.4,
      fontSize: 18, fontFace: "Microsoft JhengHei", color: C.teal, bold: true, align: "left", valign: "middle", margin: 0,
    });
    s.addShape(pptx.ShapeType.rect, { x: 7.1, y: 1.95, w: 1.2, h: 0.03, fill: { color: C.teal } });

    const kpis = [
      { category: "獲利能力", metrics: ["毛利率", "營業利益率", "淨利率", "每股盈餘 EPS"] },
      { category: "營運效率", metrics: ["存貨周轉天數", "應收帳款回收天數", "總資產周轉率"] },
      { category: "償債能力", metrics: ["流動比率", "速動比率", "利息保障倍數"] },
      { category: "成長指標", metrics: ["營收成長率 YoY/QoQ", "營業利益成長率"] },
    ];
    kpis.forEach((k, i) => {
      const ky = 2.15 + i * 0.95;
      s.addShape(pptx.ShapeType.roundRect, {
        x: 7.2, y: ky, w: 1.4, h: 0.3,
        fill: { color: C.navy }, rectRadius: 0.04,
      });
      s.addText(k.category, {
        x: 7.2, y: ky, w: 1.4, h: 0.3,
        fontSize: 10, fontFace: "Microsoft JhengHei", color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
      });
      k.metrics.forEach((m, j) => {
        s.addText(m, {
          x: 8.8, y: ky + j * 0.2, w: 3.6, h: 0.2,
          fontSize: 11, fontFace: "Microsoft JhengHei", color: C.text, align: "left", valign: "middle", margin: 0,
        });
      });
    });

    // Bottom: 決策支援
    s.addShape(pptx.ShapeType.roundRect, {
      x: 6.8, y: 5.8, w: 5.8, h: 0.7,
      fill: { color: C.amber }, rectRadius: 0.06,
    });
    s.addText("定時產出管理報表，提供經營團隊即時決策數據", {
      x: 7.0, y: 5.8, w: 5.4, h: 0.7,
      fontSize: 14, fontFace: "Microsoft JhengHei", color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
    });
  });

  // ════════════════════════════════════════════
  // SLIDE 9: IPO 準備路徑圖
  // ════════════════════════════════════════════
  addContentSlide("IPO 準備路徑圖", "上市櫃條件檢核 · 輔導團隊 · 時程規劃", (s) => {
    // Top: 上市櫃條件
    addCardAccent(s, 0.7, 1.4, 5.8, 2.5, C.navy);
    s.addText("上市櫃基本條件檢核", {
      x: 1.0, y: 1.55, w: 5.2, h: 0.4,
      fontSize: 17, fontFace: "Microsoft JhengHei", color: C.navy, bold: true, align: "left", valign: "middle", margin: 0,
    });
    s.addShape(pptx.ShapeType.rect, { x: 1.0, y: 1.95, w: 1.2, h: 0.03, fill: { color: C.navy } });

    const ipoChecks = [
      ["設立年限", "設立滿三個完整會計年度"],
      ["資本額/淨值", "興櫃：1億；上市：6億（淨值>1/3）"],
      ["獲利能力", "最近年度淨利 > 4%，或近二年均 > 3%"],
      ["股權分散", "持有股份 > 1,000 股之人數 > 300 人"],
    ];
    ipoChecks.forEach((ch, i) => {
      const cy = 2.15 + i * 0.4;
      s.addText(ch[0], { x: 1.0, y: cy, w: 1.5, h: 0.3, fontSize: 11, fontFace: "Microsoft JhengHei", color: C.text, bold: true, align: "left", valign: "middle", margin: 0 });
      s.addText(ch[1], { x: 2.6, y: cy, w: 3.7, h: 0.3, fontSize: 11, fontFace: "Microsoft JhengHei", color: C.muted, align: "left", valign: "middle", margin: 0 });
    });

    // Top right: 輔導團隊
    addCardAccent(s, 6.8, 1.4, 5.8, 2.5, C.teal);
    s.addText("IPO 輔導團隊組成", {
      x: 7.1, y: 1.55, w: 5.2, h: 0.4,
      fontSize: 17, fontFace: "Microsoft JhengHei", color: C.teal, bold: true, align: "left", valign: "middle", margin: 0,
    });
    s.addShape(pptx.ShapeType.rect, { x: 7.1, y: 1.95, w: 1.2, h: 0.03, fill: { color: C.teal } });

    const teams = [
      { role: "主辦承銷商", desc: "綜合券商，主導送件與詢價圈購" },
      { role: "會計師事務所", desc: "財稅報簽證、內控專審" },
      { role: "律師事務所", desc: "公司治理、法規遵循" },
      { role: "內部工作小組", desc: "財會、業務、研發、人資" },
    ];
    teams.forEach((tm, i) => {
      const ty = 2.2 + i * 0.4;
      s.addShape(pptx.ShapeType.ellipse, { x: 7.2, y: ty + 0.06, w: 0.08, h: 0.08, fill: { color: C.teal } });
      s.addText(tm.role, { x: 7.5, y: ty, w: 1.8, h: 0.22, fontSize: 11, fontFace: "Microsoft JhengHei", color: C.text, bold: true, align: "left", valign: "middle", margin: 0 });
      s.addText(tm.desc, { x: 9.3, y: ty, w: 3.1, h: 0.22, fontSize: 10, fontFace: "Microsoft JhengHei", color: C.muted, align: "left", valign: "middle", margin: 0 });
    });

    // Bottom: Timeline
    s.addShape(pptx.ShapeType.roundRect, {
      x: 0.7, y: 4.2, w: 11.9, h: 2.9,
      fill: { color: C.cardBg },
      rectRadius: 0.1, line: { color: C.border, width: 0.5 },
    });
    s.addText("IPO 準備時程規劃", {
      x: 1.0, y: 4.35, w: 11.3, h: 0.4,
      fontSize: 18, fontFace: "Microsoft JhengHei", color: C.navy, bold: true, align: "left", valign: "middle", margin: 0,
    });

    const phases = [
      { phase: "Phase 1\n體質調整", period: "Y1 Q1-Q2", tasks: "建立內控制度\n導入ERP系統\n財稅報制度建立" },
      { phase: "Phase 2\n輔導期", period: "Y1 Q3-Q4", tasks: "聘請輔導團隊\n內控專審\n財務整理與簽證" },
      { phase: "Phase 3\n興櫃", period: "Y2 H1", tasks: "送件興櫃\n股權分散作業\n法人說明會" },
      { phase: "Phase 4\n申請上市櫃", period: "Y2 H2~Y3", tasks: "申請上市櫃\n主管機關審查\n掛牌交易" },
    ];

    // Timeline line
    s.addShape(pptx.ShapeType.rect, { x: 1.0, y: 5.4, w: 11.3, h: 0.04, fill: { color: C.teal } });

    phases.forEach((ph, i) => {
      const px = 1.0 + i * 2.95;
      s.addShape(pptx.ShapeType.ellipse, { x: px + 1.2, y: 5.28, w: 0.28, h: 0.28, fill: { color: C.navy } });

      s.addShape(pptx.ShapeType.roundRect, {
        x: px, y: 4.85, w: 2.65, h: 0.35,
        fill: { color: C.navy }, rectRadius: 0.04,
      });
      s.addText(ph.phase, {
        x: px, y: 4.85, w: 2.65, h: 0.35,
        fontSize: 10, fontFace: "Microsoft JhengHei", color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
      });

      s.addText(ph.period, {
        x: px + 0.2, y: 5.6, w: 2.45, h: 0.25,
        fontSize: 10, fontFace: "Arial", color: C.teal, bold: true, align: "center", valign: "middle", margin: 0,
      });
      s.addText(ph.tasks, {
        x: px + 0.2, y: 5.85, w: 2.65, h: 1.0,
        fontSize: 10, fontFace: "Microsoft JhengHei", color: C.muted, align: "center", valign: "top", margin: 0,
      });
    });
  });

  // ════════════════════════════════════════════
  // SLIDE 10: 內控制度建立 — COSO 架構
  // ════════════════════════════════════════════
  addContentSlide("內部控制制度建立", "COSO 架構 · 五大循環 · 控制活動", (s) => {
    // Top left: COSO cube
    addCardAccent(s, 0.7, 1.4, 5.8, 2.5, C.navy);
    s.addText("COSO 內部控制整合架構", {
      x: 1.0, y: 1.55, w: 5.2, h: 0.4,
      fontSize: 17, fontFace: "Microsoft JhengHei", color: C.navy, bold: true, align: "left", valign: "middle", margin: 0,
    });
    s.addShape(pptx.ShapeType.rect, { x: 1.0, y: 1.95, w: 1.2, h: 0.03, fill: { color: C.navy } });

    const coso = [
      "控制環境 (Control Environment)：誠信與道德價值、董事會監督",
      "風險評估 (Risk Assessment)：目標設定、風險辨識與分析",
      "控制活動 (Control Activities)：授權、職能分工、覆核、稽核",
      "資訊與溝通 (Info & Communication)：內部報告、外部溝通管道",
      "監督 (Monitoring)：持續監控、個別評估、缺失通報",
    ];
    coso.forEach((c, i) => {
      s.addText(c, {
        x: 1.0, y: 2.15 + i * 0.32, w: 5.3, h: 0.3,
        fontSize: 11, fontFace: "Microsoft JhengHei", color: C.text, align: "left", valign: "middle", margin: 0,
      });
    });

    // Top right: 五大循環
    addCardAccent(s, 6.8, 1.4, 5.8, 2.5, C.teal);
    s.addText("內部控制五大循環", {
      x: 7.1, y: 1.55, w: 5.2, h: 0.4,
      fontSize: 17, fontFace: "Microsoft JhengHei", color: C.teal, bold: true, align: "left", valign: "middle", margin: 0,
    });
    s.addShape(pptx.ShapeType.rect, { x: 7.1, y: 1.95, w: 1.2, h: 0.03, fill: { color: C.teal } });

    const cycles = [
      { name: "1. 銷售及收款循環", desc: "訂單處理、授信、出貨、收款、客訴" },
      { name: "2. 採購及付款循環", desc: "請購、採購、驗收、付款、供應商管理" },
      { name: "3. 生產循環", desc: "生產規劃、領投料、在製品管理、製成品入庫" },
      { name: "4. 薪工循環", desc: "人事、差勤、薪資計算、發放、申報" },
      { name: "5. 固定資產循環", desc: "資產取得、折舊、維護、盤點、處分" },
    ];
    cycles.forEach((cy, i) => {
      const cy_y = 2.15 + i * 0.35;
      s.addText(cy.name, { x: 7.1, y: cy_y, w: 2.0, h: 0.3, fontSize: 11, fontFace: "Microsoft JhengHei", color: C.text, bold: true, align: "left", valign: "middle", margin: 0 });
      s.addText(cy.desc, { x: 9.2, y: cy_y, w: 3.2, h: 0.3, fontSize: 10, fontFace: "Microsoft JhengHei", color: C.muted, align: "left", valign: "middle", margin: 0 });
    });

    // Bottom: 內控制度建立步驟
    s.addShape(pptx.ShapeType.roundRect, {
      x: 0.7, y: 4.2, w: 11.9, h: 2.9,
      fill: { color: C.cardBg },
      rectRadius: 0.1, line: { color: C.border, width: 0.5 },
    });
    s.addText("內控制度建立步驟", {
      x: 1.0, y: 4.35, w: 11.3, h: 0.4,
      fontSize: 18, fontFace: "Microsoft JhengHei", color: C.navy, bold: true, align: "left", valign: "middle", margin: 0,
    });

    const steps = [
      { step: "01", title: "現況診斷", desc: "訪談各部門、盤點現有流程、辨識內控缺失" },
      { step: "02", title: "制度設計", desc: "依COSO架構及上市櫃審查準則，設計各循環控制活動" },
      { step: "03", title: "制度編寫", desc: "產出內部控制制度手冊、作業流程圖、表單設計" },
      { step: "04", title: "教育訓練", desc: "全員內控宣導、重點部門深度培訓、模擬演練" },
      { step: "05", title: "落實執行", desc: "正式公告實施、系統設定控制點、定期查核" },
    ];
    steps.forEach((st, i) => {
      const sx = 0.9 + i * 2.45;
      s.addShape(pptx.ShapeType.roundRect, {
        x: sx, y: 4.9, w: 2.2, h: 2.0,
        fill: { color: C.lightBg },
        rectRadius: 0.08, line: { color: C.border, width: 0.5 },
      });
      s.addShape(pptx.ShapeType.rect, { x: sx, y: 4.9, w: 2.2, h: 0.4, fill: { color: C.navy } });
      s.addText(st.step, {
        x: sx + 0.15, y: 4.9, w: 0.4, h: 0.4,
        fontSize: 14, fontFace: "Arial", color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
      });
      s.addText(st.title, {
        x: sx + 0.5, y: 4.9, w: 1.5, h: 0.4,
        fontSize: 14, fontFace: "Microsoft JhengHei", color: C.white, bold: true, align: "left", valign: "middle", margin: 0,
      });
      s.addText(st.desc, {
        x: sx + 0.2, y: 5.5, w: 1.8, h: 1.2,
        fontSize: 11, fontFace: "Microsoft JhengHei", color: C.text, align: "left", valign: "top", margin: 0,
      });
      if (i < steps.length - 1) {
        s.addText("\u25B6", { x: sx + 2.2, y: 5.6, w: 0.25, h: 0.5, fontSize: 16, color: C.teal, align: "center", valign: "middle", margin: 0 });
      }
    });
  });

  // ════════════════════════════════════════════
  // SLIDE 11: 內控制度 — 採購付款循環範例
  // ════════════════════════════════════════════
  addContentSlide("內控制度實例 — 採購付款循環", "流程圖 · 關鍵控制點 · 表單設計", (s) => {
    // Left: 流程圖
    addCardAccent(s, 0.7, 1.4, 5.8, 5.2, C.navy);
    s.addText("採購付款作業流程", {
      x: 1.0, y: 1.55, w: 5.2, h: 0.4,
      fontSize: 17, fontFace: "Microsoft JhengHei", color: C.navy, bold: true, align: "left", valign: "middle", margin: 0,
    });
    s.addShape(pptx.ShapeType.rect, { x: 1.0, y: 1.95, w: 1.2, h: 0.03, fill: { color: C.navy } });

    const procSteps = [
      "① 請購：需求部門提出請購單",
      "② 詢價：採購部詢價、比價、議價",
      "③ 採購：訂購單經核准後發出",
      "④ 驗收：倉管點收品檢，登錄系統",
      "⑤ 請款：供應商發票 + 驗收單送會計",
      "⑥ 審核：會計審核憑證、科目、金額",
      "⑦ 付款：依合約條件開票或匯款",
      "⑧ 入帳：傳票編製、科目餘額更新",
    ];
    procSteps.forEach((ps, i) => {
      const py = 2.15 + i * 0.48;
      s.addText(ps, {
        x: 1.0, y: py, w: 5.3, h: 0.4,
        fontSize: 12, fontFace: "Microsoft JhengHei", color: C.text, align: "left", valign: "middle", margin: 0,
      });
      if (i < procSteps.length - 1) {
        s.addShape(pptx.ShapeType.rect, { x: 1.8, y: py + 0.36, w: 0.3, h: 0.12, fill: { color: C.border } });
      }
    });

    // Right: 控制點與表單
    addCardAccent(s, 6.8, 1.4, 5.8, 2.5, C.teal);
    s.addText("關鍵控制點", {
      x: 7.1, y: 1.55, w: 5.2, h: 0.4,
      fontSize: 17, fontFace: "Microsoft JhengHei", color: C.teal, bold: true, align: "left", valign: "middle", margin: 0,
    });
    s.addShape(pptx.ShapeType.rect, { x: 7.1, y: 1.95, w: 1.2, h: 0.03, fill: { color: C.teal } });

    const controls = [
      "請購與採購職能分工（避免球員兼裁判）",
      "採購金額分級授權（經理/副總/總經理）",
      "三家以上比價記錄（大額採購）",
      "驗收單據必須與訂購單核對",
      "會計審核憑證三要件：真實/合法/完整",
      "付款期限控管，避免逾期或過早付款",
    ];
    controls.forEach((ct, i) => {
      const cty = 2.15 + i * 0.28;
      s.addText(`\u25B6 ${ct}`, {
        x: 7.1, y: cty, w: 5.3, h: 0.25,
        fontSize: 11, fontFace: "Microsoft JhengHei", color: C.text, align: "left", valign: "middle", margin: 0,
      });
    });

    // Right bottom: 表單
    addCardAccent(s, 6.8, 4.1, 5.8, 2.5, C.amber);
    s.addText("內控相關表單", {
      x: 7.1, y: 4.25, w: 5.2, h: 0.4,
      fontSize: 17, fontFace: "Microsoft JhengHei", color: C.amber, bold: true, align: "left", valign: "middle", margin: 0,
    });
    s.addShape(pptx.ShapeType.rect, { x: 7.1, y: 4.65, w: 1.2, h: 0.03, fill: { color: C.amber } });

    const forms = [
      "請購單（含需求說明、預算科目）",
      "採購比價記錄表",
      "訂購單 / 採購合約",
      "驗收單（品檢 + 數量確認）",
      "應付憑單（發票 + 驗收單 + 訂購單）",
      "付款申請單（含核准簽核）",
    ];
    forms.forEach((fm, i) => {
      const fmy = 4.85 + i * 0.28;
      s.addText(`\u25B6 ${fm}`, {
        x: 7.1, y: fmy, w: 5.3, h: 0.25,
        fontSize: 11, fontFace: "Microsoft JhengHei", color: C.text, align: "left", valign: "middle", margin: 0,
      });
    });
  });

  // ════════════════════════════════════════════
  // SLIDE 12: 內控制度 — 其他循環概述
  // ════════════════════════════════════════════
  addContentSlide("內控制度 — 其他循環概述", "銷售收款 · 薪工 · 生產 · 固定資產", (s) => {
    const loops = [
      {
        title: "銷售及收款循環",
        controls: ["授信額度審核（超額須特准）", "出貨與開立發票職能分工", "應收帳款帳齡分析與催收", "收款沖帳即時核對"],
        color: C.navy,
      },
      {
        title: "生產循環",
        controls: ["生產工單開立與核准", "領料單與BOM表核對", "在製品盤點與差異追蹤", "製成品入庫品檢記錄"],
        color: C.teal,
      },
      {
        title: "薪工循環",
        controls: ["人事異動與薪資計算分離", "出勤記錄與薪資核對", "年終/獎金計算需經核准", "勞健保/退休金申報正確性"],
        color: C.amber,
      },
      {
        title: "固定資產循環",
        controls: ["資產取得資本化/費用化判定", "資產標籤管理與盤點", "折舊計算與耐用年限檢討", "資產報廢/處分需經核准"],
        color: C.purple,
      },
    ];

    loops.forEach((lp, i) => {
      const x = 0.6 + i * 3.15;
      addCardAccent(s, x, 1.4, 2.9, 4.2, lp.color);
      s.addText(lp.title, {
        x: x + 0.15, y: 1.55, w: 2.6, h: 0.4,
        fontSize: 16, fontFace: "Microsoft JhengHei", color: lp.color, bold: true, align: "left", valign: "middle", margin: 0,
      });
      s.addShape(pptx.ShapeType.rect, { x: x + 0.15, y: 1.95, w: 1.2, h: 0.03, fill: { color: lp.color } });
      lp.controls.forEach((ctl, j) => {
        const cty = 2.15 + j * 0.65;
        s.addShape(pptx.ShapeType.ellipse, { x: x + 0.3, y: cty + 0.08, w: 0.07, h: 0.07, fill: { color: lp.color } });
        s.addText(ctl, {
          x: x + 0.5, y: cty, w: 2.2, h: 0.5,
          fontSize: 11, fontFace: "Microsoft JhengHei", color: C.text, align: "left", valign: "top", margin: 0,
        });
      });
    });

    // Bottom note
    s.addShape(pptx.ShapeType.roundRect, {
      x: 0.6, y: 5.9, w: 12.13, h: 1.2,
      fill: { color: C.cardBg },
      rectRadius: 0.1, line: { color: C.border, width: 0.5 },
    });
    s.addShape(pptx.ShapeType.rect, { x: 0.6, y: 5.9, w: 12.13, h: 0.06, fill: { color: C.navy } });
    s.addText("內控制度撰寫文件產出", {
      x: 1.0, y: 6.05, w: 11.3, h: 0.35,
      fontSize: 16, fontFace: "Microsoft JhengHei", color: C.navy, bold: true, align: "left", valign: "middle", margin: 0,
    });

    const docItems = [
      "內部控制制度手冊（含各循環控制活動）  |  作業流程圖（SOP + 流程圖）  |  控制自行評估表（CSA）",
      "內部稽核計畫與報告  |  職能分工表（授權矩陣）  |  各循環使用表單（已設計完成）",
    ];
    docItems.forEach((doc, i) => {
      s.addText(doc, {
        x: 1.0, y: 6.45 + i * 0.3, w: 11.3, h: 0.3,
        fontSize: 12, fontFace: "Microsoft JhengHei", color: C.text, align: "left", valign: "middle", margin: 0,
      });
    });
  });

  // ════════════════════════════════════════════
  // SLIDE 13: ERP 需求分析與評估維度
  // ════════════════════════════════════════════
  addContentSlide("ERP 需求分析與評估維度", "功能需求 · 技術評估 · 成本分析 · 廠商評比", (s) => {
    // Four dimensions
    const dims = [
      { title: "功能需求分析", items: [
        "財務會計：總帳/應收應付/票據/固定資產/成本",
        "進銷存管理：採購/銷售/庫存/物流",
        "製造管理：BOM/製令/工單/產能規劃",
        "研發管理：專案成本/研發費用歸屬",
        "HR 管理：人事/薪資/出勤/福利",
      ], color: C.navy },
      { title: "技術與整合", items: [
        "雲端/地端部署彈性",
        "API 開放性與第三方系統串接",
        "資料庫架構與效能",
        "行動裝置支援",
        "資料安全與備份機制",
      ], color: C.teal },
      { title: "成本評估", items: [
        "軟體授權費（買斷/年租）",
        "導入顧問費",
        "客製開發費用",
        "硬體/雲端基礎設施費用",
        "年度維護費（通常 15-20%）",
      ], color: C.amber },
      { title: "廠商評估", items: [
        "在地化服務能力與售後支援",
        "導入團隊經驗與參考客戶",
        "產品迭代更新頻率",
        "教育訓練與知識轉移",
        "市場佔有率與產業口碑",
      ], color: C.purple },
    ];

    dims.forEach((d, i) => {
      const x = 0.5 + i * 3.2;
      addCardAccent(s, x, 1.4, 2.95, 5.2, d.color);
      s.addText(d.title, {
        x: x + 0.15, y: 1.55, w: 2.65, h: 0.4,
        fontSize: 16, fontFace: "Microsoft JhengHei", color: d.color, bold: true, align: "left", valign: "middle", margin: 0,
      });
      s.addShape(pptx.ShapeType.rect, { x: x + 0.15, y: 1.95, w: 1.2, h: 0.03, fill: { color: d.color } });
      d.items.forEach((it, j) => {
        const iy = 2.15 + j * 0.65;
        s.addShape(pptx.ShapeType.ellipse, { x: x + 0.3, y: iy + 0.08, w: 0.07, h: 0.07, fill: { color: d.color } });
        s.addText(it, {
          x: x + 0.5, y: iy, w: 2.25, h: 0.5,
          fontSize: 11, fontFace: "Microsoft JhengHei", color: C.text, align: "left", valign: "top", margin: 0,
        });
      });
    });
  });

  // ════════════════════════════════════════════
  // SLIDE 14: ERP 導入步驟與組織
  // ════════════════════════════════════════════
  addContentSlide("ERP 導入步驟與組織", "導入團隊 · 階段時程 · 風險管理", (s) => {
    // Left: 導入組織
    addCardAccent(s, 0.7, 1.4, 3.8, 5.2, C.navy);
    s.addText("導入團隊組織", {
      x: 1.0, y: 1.55, w: 3.2, h: 0.4,
      fontSize: 17, fontFace: "Microsoft JhengHei", color: C.navy, bold: true, align: "left", valign: "middle", margin: 0,
    });
    s.addShape(pptx.ShapeType.rect, { x: 1.0, y: 1.95, w: 1.2, h: 0.03, fill: { color: C.navy } });

    const orgs = [
      { role: "專案發起人", person: "高階主管", desc: "賦予資源與決策" },
      { role: "專案經理", person: "財務/IT主管", desc: "進度管控與溝通" },
      { role: "關鍵使用者", person: "各部門種子", desc: "需求確認與測試" },
      { role: "IT 技術組", person: "資訊人員", desc: "系統維運與串接" },
      { role: "外部顧問", person: "ERP廠商", desc: "導入輔導與培訓" },
    ];
    orgs.forEach((og, i) => {
      const oy = 2.15 + i * 0.75;
      s.addShape(pptx.ShapeType.roundRect, {
        x: 1.0, y: oy, w: 3.3, h: 0.6,
        fill: { color: C.lightBg },
        rectRadius: 0.06, line: { color: C.border, width: 0.5 },
      });
      s.addText(og.role, { x: 1.15, y: oy + 0.02, w: 3.0, h: 0.22, fontSize: 11, fontFace: "Microsoft JhengHei", color: C.text, bold: true, align: "left", valign: "middle", margin: 0 });
      s.addText(`${og.person}  |  ${og.desc}`, { x: 1.15, y: oy + 0.25, w: 3.0, h: 0.3, fontSize: 9, fontFace: "Microsoft JhengHei", color: C.muted, align: "left", valign: "middle", margin: 0 });
    });

    // Middle: 導入階段
    addCardAccent(s, 4.8, 1.4, 3.8, 5.2, C.teal);
    s.addText("導入階段時程", {
      x: 5.1, y: 1.55, w: 3.2, h: 0.4,
      fontSize: 17, fontFace: "Microsoft JhengHei", color: C.teal, bold: true, align: "left", valign: "middle", margin: 0,
    });
    s.addShape(pptx.ShapeType.rect, { x: 5.1, y: 1.95, w: 1.2, h: 0.03, fill: { color: C.teal } });

    const erpPhases = [
      { phase: "Phase 1", period: "M1-M2", tasks: "需求訪談\n現況流程盤點\n系統選商決策" },
      { phase: "Phase 2", period: "M3-M5", tasks: "系統建置\n參數設定\n基礎資料匯入" },
      { phase: "Phase 3", period: "M6-M7", tasks: "單元測試\n整合測試\nUAT 驗收" },
      { phase: "Phase 4", period: "M8", tasks: "上線準備\n資料轉換\n教育訓練" },
      { phase: "Phase 5", period: "M9+", tasks: "正式上線\n平行測試\n持續優化" },
    ];
    erpPhases.forEach((ep, i) => {
      const epy = 2.15 + i * 0.85;
      s.addShape(pptx.ShapeType.roundRect, {
        x: 5.1, y: epy, w: 3.3, h: 0.7,
        fill: { color: C.lightBg },
        rectRadius: 0.06, line: { color: C.border, width: 0.5 },
      });
      s.addText(`${ep.phase}: ${ep.period}`, { x: 5.25, y: epy + 0.02, w: 3.0, h: 0.22, fontSize: 10, fontFace: "Arial", color: C.teal, bold: true, align: "left", valign: "middle", margin: 0 });
      s.addText(ep.tasks, { x: 5.25, y: epy + 0.25, w: 3.0, h: 0.4, fontSize: 10, fontFace: "Microsoft JhengHei", color: C.text, align: "left", valign: "top", margin: 0 });
    });

    // Right: 風險管理
    addCardAccent(s, 8.9, 1.4, 3.8, 5.2, C.amber);
    s.addText("風險管理", {
      x: 9.2, y: 1.55, w: 3.2, h: 0.4,
      fontSize: 17, fontFace: "Microsoft JhengHei", color: C.amber, bold: true, align: "left", valign: "middle", margin: 0,
    });
    s.addShape(pptx.ShapeType.rect, { x: 9.2, y: 1.95, w: 1.2, h: 0.03, fill: { color: C.amber } });

    const risks = [
      ["高", "需求變更頻繁", "明確範疇、變更管制"],
      ["高", "資料轉換錯誤", "多次演練、比對驗證"],
      ["中", "人員抗拒變革", "充分溝通、教育訓練"],
      ["中", "專案時程延誤", "週進度追蹤、buffer管理"],
      ["低", "廠商支援不足", "合約明訂SLA、備援廠商"],
    ];
    risks.forEach((rk, i) => {
      const rky = 2.15 + i * 0.78;
      // Risk level badge
      const levelColor = rk[0] === "高" ? C.red : rk[0] === "中" ? C.amber : C.green;
      s.addShape(pptx.ShapeType.roundRect, {
        x: 9.2, y: rky, w: 0.4, h: 0.28,
        fill: { color: levelColor }, rectRadius: 0.04,
      });
      s.addText(rk[0], { x: 9.2, y: rky, w: 0.4, h: 0.28, fontSize: 9, fontFace: "Arial", color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
      s.addText(rk[1], { x: 9.7, y: rky - 0.02, w: 2.8, h: 0.16, fontSize: 11, fontFace: "Microsoft JhengHei", color: C.text, bold: true, align: "left", valign: "middle", margin: 0 });
      s.addText(`\u25B6 ${rk[2]}`, { x: 9.7, y: rky + 0.15, w: 2.8, h: 0.16, fontSize: 10, fontFace: "Microsoft JhengHei", color: C.muted, align: "left", valign: "middle", margin: 0 });
      if (i < risks.length - 1) {
        s.addShape(pptx.ShapeType.rect, { x: 10.5, y: rky + 0.32, w: 0.3, h: 0.44, fill: { color: C.border } });
      }
    });
  });

  // ════════════════════════════════════════════
  // SLIDE 15: 候選 ERP 方案比較
  // ════════════════════════════════════════════
  addContentSlide("候選 ERP 方案比較", "鼎新 · SAP B1 · Oracle NetSuite · Odoo", (s) => {
    s.addShape(pptx.ShapeType.roundRect, {
      x: 0.5, y: 1.3, w: 12.33, h: 5.8,
      fill: { color: C.cardBg },
      rectRadius: 0.1, line: { color: C.border, width: 0.5 },
    });

    // Table header
    const headers = ["評估項目", "鼎新 ERP", "SAP Business One", "Oracle NetSuite", "Odoo（開源）"];
    const colW = [1.8, 2.4, 2.4, 2.4, 2.4];
    let hx = 0.8;
    headers.forEach((h, i) => {
      s.addShape(pptx.ShapeType.rect, { x: hx, y: 1.5, w: colW[i], h: 0.45, fill: { color: C.navy } });
      s.addText(h, {
        x: hx, y: 1.5, w: colW[i], h: 0.45,
        fontSize: 12, fontFace: "Microsoft JhengHei", color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
      });
      hx += colW[i];
    });

    // Table rows
    const rows = [
      ["適合規模", "中大型企業", "中小型企業", "中大型跨國企業", "中小型企業"],
      ["部署方式", "地端 / 雲端", "地端 / 雲端", "雲端原生", "地端 / 雲端"],
      ["在地化支援", "★★★★★", "★★★☆☆", "★★★☆☆", "★★★☆☆"],
      ["導入費用", "中高", "中", "高", "低（開源免費）"],
      ["導入週期", "4-8 個月", "3-6 個月", "4-8 個月", "2-4 個月"],
      ["製造管理", "完整支援", "基本支援", "基本支援", "模組擴展"],
      ["客製彈性", "中等", "中等", "中等", "高度彈性"],
      ["台灣市佔率", "最高", "中等", "較低", "新興"],
    ];

    rows.forEach((row, ri) => {
      const ry = 1.95 + ri * 0.58;
      const bgColor = ri % 2 === 0 ? C.lightBg : C.white;
      let rx = 0.8;
      row.forEach((cell, ci) => {
        s.addShape(pptx.ShapeType.rect, { x: rx, y: ry, w: colW[ci], h: 0.58, fill: { color: bgColor } });
        if (ci < 1) {
          s.addShape(pptx.ShapeType.rect, { x: rx, y: ry, w: colW[ci], h: 0.58, fill: { color: C.lightBg } });
        }
        s.addText(cell, {
          x: rx + 0.1, y: ry, w: colW[ci] - 0.2, h: 0.58,
          fontSize: 11, fontFace: "Microsoft JhengHei", color: ci === 0 ? C.navy : C.text, bold: ci === 0, align: "center", valign: "middle", margin: 0,
        });
        rx += colW[ci];
      });
    });

    // Recommendation bar
    s.addShape(pptx.ShapeType.roundRect, {
      x: 0.8, y: 6.7, w: 11.73, h: 0.55,
      fill: { color: C.teal }, rectRadius: 0.05,
    });
    s.addText("建議：依公司階段性需求，可先導入鼎新 ERP（台灣在地化最佳）或 Odoo（彈性高、成本低），後續再評估升級方案", {
      x: 1.0, y: 6.7, w: 11.33, h: 0.55,
      fontSize: 13, fontFace: "Microsoft JhengHei", color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
    });
  });

  // ════════════════════════════════════════════
  // SLIDE 16: 專案流程規劃與控管
  // ════════════════════════════════════════════
  addContentSlide("專案流程規劃與控管", "WBS · 里程碑 · 溝通計畫", (s) => {
    // Left: 專案WBS
    addCardAccent(s, 0.7, 1.4, 3.8, 5.2, C.navy);
    s.addText("專案工作分解（WBS）", {
      x: 1.0, y: 1.55, w: 3.2, h: 0.4,
      fontSize: 16, fontFace: "Microsoft JhengHei", color: C.navy, bold: true, align: "left", valign: "middle", margin: 0,
    });
    s.addShape(pptx.ShapeType.rect, { x: 1.0, y: 1.95, w: 1.2, h: 0.03, fill: { color: C.navy } });

    const wbsItems = [
      "1.0 會計制度建立（帳務/稅務/憑證）",
      "2.0 存貨管理與成本制度",
      "3.0 財務報表與分析體系",
      "4.0 內控制度建立（五循環）",
      "5.0 ERP 系統導入",
      "6.0 IPO 輔導與送件",
      "7.0 專案管理與品質保證",
    ];
    wbsItems.forEach((wbs, i) => {
      const wy = 2.15 + i * 0.52;
      s.addText(wbs, { x: 1.0, y: wy, w: 3.3, h: 0.45, fontSize: 11, fontFace: "Microsoft JhengHei", color: C.text, align: "left", valign: "middle", margin: 0 });
      if (i < wbsItems.length - 1) {
        s.addShape(pptx.ShapeType.rect, { x: 1.8, y: wy + 0.4, w: 0.2, h: 0.12, fill: { color: C.border } });
      }
    });

    // Middle: 里程碑
    addCardAccent(s, 4.8, 1.4, 3.8, 5.2, C.teal);
    s.addText("里程碑時程", {
      x: 5.1, y: 1.55, w: 3.2, h: 0.4,
      fontSize: 16, fontFace: "Microsoft JhengHei", color: C.teal, bold: true, align: "left", valign: "middle", margin: 0,
    });
    s.addShape(pptx.ShapeType.rect, { x: 5.1, y: 1.95, w: 1.2, h: 0.03, fill: { color: C.teal } });

    const milestones = [
      ["M1", "M1-M2", "內控制度完成編寫"],
      ["M2", "M3-M4", "會計制度與憑證辦法公告實施"],
      ["M3", "M5-M7", "ERP 導入完成上線"],
      ["M4", "M8-M10", "內控制度落實查核"],
      ["M5", "M11-M12", "IPO 輔導團隊進場"],
      ["M6", "Y2 H1", "送件興櫃"],
    ];
    milestones.forEach((ms, i) => {
      const msy = 2.15 + i * 0.6;
      s.addShape(pptx.ShapeType.roundRect, {
        x: 5.1, y: msy, w: 0.4, h: 0.25,
        fill: { color: C.teal }, rectRadius: 0.04,
      });
      s.addText(ms[0], { x: 5.1, y: msy, w: 0.4, h: 0.25, fontSize: 9, fontFace: "Arial", color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
      s.addText(ms[1], { x: 5.6, y: msy - 0.02, w: 1.0, h: 0.25, fontSize: 9, fontFace: "Arial", color: C.muted, align: "left", valign: "middle", margin: 0 });
      s.addText(ms[2], { x: 5.6, y: msy + 0.22, w: 2.8, h: 0.25, fontSize: 11, fontFace: "Microsoft JhengHei", color: C.text, align: "left", valign: "middle", margin: 0 });
    });

    // Right: 溝通與控管
    addCardAccent(s, 8.9, 1.4, 3.8, 5.2, C.amber);
    s.addText("溝通與控管機制", {
      x: 9.2, y: 1.55, w: 3.2, h: 0.4,
      fontSize: 16, fontFace: "Microsoft JhengHei", color: C.amber, bold: true, align: "left", valign: "middle", margin: 0,
    });
    s.addShape(pptx.ShapeType.rect, { x: 9.2, y: 1.95, w: 1.2, h: 0.03, fill: { color: C.amber } });

    const comms = [
      ["週報", "每週工作進度更新"],
      ["月會", "跨部門進度檢討會議"],
      ["季報告", "向經營層彙報成果"],
      ["變更管制", "範疇/時程變更需書面核准"],
      ["品質查核", "每階段 deliverables 檢核"],
      ["風險登錄", "風險項目追蹤與因應"],
    ];
    comms.forEach((cm, i) => {
      const cmy = 2.15 + i * 0.6;
      s.addShape(pptx.ShapeType.rect, { x: 9.2, y: cmy, w: 0.04, h: 0.45, fill: { color: C.amber } });
      s.addText(cm[0], { x: 9.4, y: cmy - 0.02, w: 1.5, h: 0.22, fontSize: 11, fontFace: "Microsoft JhengHei", color: C.text, bold: true, align: "left", valign: "middle", margin: 0 });
      s.addText(cm[1], { x: 9.4, y: cmy + 0.2, w: 3.1, h: 0.2, fontSize: 10, fontFace: "Microsoft JhengHei", color: C.muted, align: "left", valign: "middle", margin: 0 });
    });
  });

  // ════════════════════════════════════════════
  // SLIDE 17: 行動方案與下一步
  // ════════════════════════════════════════════
  const s17 = pptx.addSlide();
  s17.background = { color: C.darkNavy };

  s17.addShape(pptx.ShapeType.ellipse, { x: -2.5, y: -2, w: 5, h: 5, fill: { color: C.navy, transparency: 40 } });
  s17.addShape(pptx.ShapeType.ellipse, { x: 10, y: 4.5, w: 5, h: 5, fill: { color: C.navy, transparency: 50 } });
  s17.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.33, h: 0.06, fill: { color: C.teal } });
  s17.addShape(pptx.ShapeType.rect, { x: 0, y: 7.44, w: 13.33, h: 0.06, fill: { color: C.teal } });

  s17.addText("NEXT STEPS", {
    x: 1.2, y: 0.5, w: 10.93, h: 0.5,
    fontSize: 14, fontFace: "Arial", color: C.teal, bold: true, align: "left", valign: "middle",
    charSpacing: 6, margin: 0,
  });

  s17.addText("行動方案與下一步", {
    x: 1.2, y: 1.0, w: 10.93, h: 1.0,
    fontSize: 38, fontFace: "Microsoft JhengHei", color: C.white, bold: true, align: "left", valign: "middle", margin: 0,
  });

  s17.addShape(pptx.ShapeType.rect, { x: 1.2, y: 2.1, w: 2.5, h: 0.04, fill: { color: C.amber } });

  // Priority actions
  const actions = [
    { priority: "P0", label: "內控制度編寫", desc: "立即啟動五大循環內控手冊編寫，預計 2 個月完成", timeline: "Month 1-2" },
    { priority: "P0", label: "ERP 系統選商", desc: "成立評選小組，完成需求訪談與廠商評比", timeline: "Month 1-3" },
    { priority: "P1", label: "會計制度優化", desc: "憑證管理辦法、稅務行事曆、流程標準化", timeline: "Month 2-4" },
    { priority: "P1", label: "存貨與成本制度", desc: "料件編碼、出入庫流程、盤點制度建立", timeline: "Month 2-4" },
    { priority: "P2", label: "財務分析體系", desc: "管理報表自動化、KPI 儀表板建置", timeline: "Month 4-6" },
    { priority: "P2", label: "IPO 輔導遴選", desc: "評估承銷商與會計師事務所", timeline: "Month 6-8" },
  ];

  actions.forEach((act, i) => {
    const row = Math.floor(i / 2);
    const col = i % 2;
    const ax = 1.2 + col * 5.9;
    const ay = 2.5 + row * 1.5;

    s17.addShape(pptx.ShapeType.roundRect, {
      x: ax, y: ay, w: 5.5, h: 1.2,
      fill: { color: C.navy, transparency: 30 },
      rectRadius: 0.08,
      line: { color: C.teal, width: 0.5 },
    });

    // Priority badge
    const badgeColor = act.priority === "P0" ? C.red : act.priority === "P1" ? C.amber : C.teal;
    s17.addShape(pptx.ShapeType.roundRect, {
      x: ax + 0.2, y: ay + 0.15, w: 0.4, h: 0.28,
      fill: { color: badgeColor }, rectRadius: 0.04,
    });
    s17.addText(act.priority, {
      x: ax + 0.2, y: ay + 0.15, w: 0.4, h: 0.28,
      fontSize: 10, fontFace: "Arial", color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
    });

    s17.addText(act.label, {
      x: ax + 0.75, y: ay + 0.08, w: 3.5, h: 0.35,
      fontSize: 16, fontFace: "Microsoft JhengHei", color: C.white, bold: true, align: "left", valign: "middle", margin: 0,
    });
    s17.addText(act.desc, {
      x: ax + 0.75, y: ay + 0.45, w: 3.5, h: 0.4,
      fontSize: 11, fontFace: "Microsoft JhengHei", color: C.lightMuted, align: "left", valign: "middle", margin: 0,
    });

    // Timeline tag
    s17.addShape(pptx.ShapeType.roundRect, {
      x: ax + 4.0, y: ay + 0.85, w: 1.3, h: 0.25,
      fill: { color: C.teal, transparency: 40 }, rectRadius: 0.04,
    });
    s17.addText(act.timeline, {
      x: ax + 4.0, y: ay + 0.85, w: 1.3, h: 0.25,
      fontSize: 9, fontFace: "Arial", color: C.teal, bold: true, align: "center", valign: "middle", margin: 0,
    });
  });

  // Bottom quote
  s17.addText("建立制度化財務會計體系，為公司永續發展與 IPO 上市櫃奠定堅實基礎", {
    x: 1.2, y: 6.6, w: 10.93, h: 0.5,
    fontSize: 16, fontFace: "Microsoft JhengHei", color: C.amber, bold: true, align: "center", valign: "middle", margin: 0,
  });

  // ════════════════════════════════════════════
  // SAVE
  // ════════════════════════════════════════════
  const outputPath = "創思達科技_IPO_內控_ERP規劃簡報.pptx";
  await pptx.writeFile({ fileName: outputPath });
  console.log(`✅ 簡報已生成: ${outputPath}`);
}

generate().catch(console.error);
