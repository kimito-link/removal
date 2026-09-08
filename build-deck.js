// 提案書スライド生成 — Google評判 総合コンサルパック（レビューブースト）
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const {
  FaStar, FaRegStar, FaSearch, FaRobot, FaBalanceScale, FaChartLine,
  FaExclamationTriangle, FaTimesCircle, FaCheckCircle, FaArrowRight,
  FaGlobe, FaMobileAlt, FaCommentDots, FaArrowDown, FaUsers, FaClock,
  FaBullhorn, FaShieldAlt, FaHandshake, FaRegSmile, FaRegCreditCard
} = require("react-icons/fa");

// palette
const NAVY = "0F2C4D", NAVY2 = "1A4170", BLUE = "2563EB", GOLD = "F5A623",
  GOLDD = "D98E0B", GREEN = "16A34A", LINEG = "06C755", PURPLE = "7C3AED",
  TEAL = "0E7490", RED = "DC2626",
  INK = "1F2937", INKSOFT = "5B6776", LINE = "E2E8F0", BG = "F6F9FC",
  BGALT = "EEF4FB", WHITE = "FFFFFF";
const HEAD = "Georgia", BODY = "Calibri";

const iconCache = {};
async function icon(Comp, color, size = 256) {
  const key = (Comp.name || "i") + color + size;
  if (iconCache[key]) return iconCache[key];
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color, size: String(size) }));
  const png = await sharp(Buffer.from(svg)).png().toBuffer();
  const data = "image/png;base64," + png.toString("base64");
  iconCache[key] = data;
  return data;
}
const sh = () => ({ type: "outer", color: "0F2C4D", blur: 9, offset: 3, angle: 90, opacity: 0.16 });
const shSoft = () => ({ type: "outer", color: "000000", blur: 7, offset: 2, angle: 90, opacity: 0.10 });

(async () => {
  const p = new pptxgen();
  p.layout = "LAYOUT_WIDE";
  p.author = "レビューブースト";
  p.title = "Google評判 総合コンサルパック ご提案書";
  const W = 13.3, H = 7.5;

  let _pageNo = 1; // title is page 1 (no footer); content footers auto-increment from 2
  function footer(s) {
    _pageNo += 1;
    s.addText("レビューブースト ｜ Google評判 総合コンサルパック", {
      x: 0.6, y: H - 0.42, w: 9, h: 0.3, fontFace: BODY, fontSize: 9, color: "9AA7B5", align: "left",
    });
    s.addText(String(_pageNo).padStart(2, "0"), {
      x: W - 1.1, y: H - 0.42, w: 0.5, h: 0.3, fontFace: HEAD, fontSize: 10, color: "9AA7B5", align: "right",
    });
  }
  async function starRow(s, x, y, filled, total, size, onDark) {
    const ec = onDark ? "33507A" : "D7E0EA";
    for (let i = 0; i < total; i++) {
      s.addImage({ data: await icon(i < filled ? FaStar : FaRegStar, "#" + (i < filled ? GOLD : ec)), x: x + i * (size + 0.04), y, w: size, h: size });
    }
  }
  // generic section header (light slides)
  function head(s, title, sub) {
    s.addText(title, { x: 0.7, y: 0.5, w: 12, h: 0.7, fontFace: HEAD, fontSize: 27, bold: true, color: NAVY });
    if (sub) s.addText(sub, { x: 0.72, y: 1.18, w: 12, h: 0.4, fontFace: BODY, fontSize: 13.5, color: INKSOFT });
  }

  // ===================== SLIDE 1 — TITLE =====================
  {
    const s = p.addSlide();
    s.background = { color: NAVY };
    s.addShape(p.shapes.OVAL, { x: 9.6, y: -2.0, w: 6.5, h: 6.5, fill: { color: NAVY2, transparency: 35 }, line: { type: "none" } });
    s.addShape(p.shapes.OVAL, { x: -2.2, y: 4.2, w: 6.0, h: 6.0, fill: { color: BLUE, transparency: 70 }, line: { type: "none" } });

    s.addText("GOOGLE REPUTATION ｜ ALL-IN-ONE PACKAGE", {
      x: 0.9, y: 1.05, w: 11, h: 0.4, fontFace: BODY, fontSize: 13, color: "8FB3E6", charSpacing: 3, bold: true,
    });
    await starRow(s, 0.92, 1.65, 2, 5, 0.4, true);
    s.addImage({ data: await icon(FaArrowRight, "#7C90A8"), x: 3.4, y: 1.74, w: 0.34, h: 0.34 });
    await starRow(s, 3.95, 1.65, 4, 5, 0.4, true);

    s.addText([
      { text: "Googleの評判、まるごとお任せ。", options: { color: WHITE, breakLine: true } },
      { text: "「選ばれ続けるお店」を、", options: { color: WHITE, breakLine: true } },
      { text: "ひとつのパッケージで。", options: { color: GOLD } },
    ], { x: 0.85, y: 2.35, w: 11.8, h: 2.5, fontFace: HEAD, fontSize: 38, bold: true, lineSpacing: 50 });

    s.addText(
      "悪い口コミ・変なサジェスト・ネットの評判 ——\nAIによる口コミ対策サポート、提携弁護士の窓口、公式WEB・LINE・アプリまで。\n契約中ずっと使える月額パッケージで、Googleの評判を総合的に改善します。",
      { x: 0.9, y: 5.25, w: 11.4, h: 1.2, fontFace: BODY, fontSize: 15, color: "C7D6EA", lineSpacing: 23 }
    );
    s.addText("Google評判 総合コンサルパック　ご提案書", {
      x: 0.9, y: 6.7, w: 11, h: 0.4, fontFace: BODY, fontSize: 13, color: "8497AE",
    });
  }

  // ===================== SLIDE 2 — PROBLEM =====================
  {
    const s = p.addSlide();
    s.background = { color: BG };
    head(s, "こんな「ネットの評判」のお悩み、ありませんか？", "店舗・クリニック・サロン・士業 — Googleやネット検索で集客する事業者さま共通の課題です。");
    const probs = [
      [FaExclamationTriangle, "事実と違う★1を書かれた", "来店していない人や誹謗中傷に近い低評価で、店舗の印象が下がっている。", GOLDD],
      [FaChartLine, "低評価で新規客が減った", "★評価を見て来店をやめる人が増え、予約・問い合わせが落ちている。", GOLDD],
      [FaSearch, "変なサジェストが出る", "店名で検索すると「店名＋ネガティブな言葉」が候補に出て印象を損ねる。", NAVY],
      [FaGlobe, "検索上位にネガティブ情報", "悪い評判の記事やまとめが検索結果の上位に出て、ずっと気になっている。", RED],
      [FaRegStar, "良い口コミが集まらない", "満足客はいるのに、レビューを書いてもらう仕組みがなく★が増えない。", GOLDD],
      [FaClock, "対策する時間も知識もない", "日々の営業で手一杯。何から手をつければいいか分からず後回しに。", INKSOFT],
    ];
    const cw = 3.86, ch = 1.72, gx = 0.3, gy = 0.3, x0 = 0.7, y0 = 1.78;
    for (let i = 0; i < probs.length; i++) {
      const col = i % 3, row = Math.floor(i / 3);
      const x = x0 + col * (cw + gx), y = y0 + row * (ch + gy);
      s.addShape(p.shapes.ROUNDED_RECTANGLE, { x, y, w: cw, h: ch, fill: { color: WHITE }, line: { color: LINE, width: 1 }, rectRadius: 0.08, shadow: shSoft() });
      s.addShape(p.shapes.OVAL, { x: x + 0.28, y: y + 0.28, w: 0.62, h: 0.62, fill: { color: BGALT }, line: { type: "none" } });
      s.addImage({ data: await icon(probs[i][0], "#" + probs[i][3]), x: x + 0.42, y: y + 0.42, w: 0.34, h: 0.34 });
      s.addText(probs[i][1], { x: x + 1.05, y: y + 0.26, w: cw - 1.25, h: 0.66, fontFace: BODY, fontSize: 14.5, bold: true, color: NAVY, valign: "middle" });
      s.addText(probs[i][2], { x: x + 0.3, y: y + 0.96, w: cw - 0.55, h: 0.66, fontFace: BODY, fontSize: 11.5, color: INKSOFT, lineSpacing: 15 });
    }
    s.addText([
      { text: "これ全部、", options: { color: NAVY } },
      { text: "ひとつのパッケージ", options: { color: GOLDD } },
      { text: "でまとめてお任せできます。", options: { color: NAVY } },
    ], { x: 0.7, y: 6.35, w: 12, h: 0.5, fontFace: HEAD, fontSize: 18, bold: true, align: "center" });
    footer(s);
  }

  // ===================== SLIDE 3 — INSIGHT =====================
  {
    const s = p.addSlide();
    s.background = { color: NAVY };
    s.addShape(p.shapes.OVAL, { x: 10.2, y: 3.4, w: 7, h: 7, fill: { color: NAVY2, transparency: 45 }, line: { type: "none" } });

    s.addText("私たちの考え方", { x: 0.9, y: 0.95, w: 11, h: 0.5, fontFace: BODY, fontSize: 15, color: GOLD, charSpacing: 2, bold: true });
    s.addText([
      { text: "口コミを「消す」ことが、", options: { color: WHITE, breakLine: true } },
      { text: "ゴールではありません。", options: { color: GOLD } },
    ], { x: 0.85, y: 1.55, w: 11.6, h: 1.7, fontFace: HEAD, fontSize: 36, bold: true, lineSpacing: 48 });
    s.addText(
      "削除はあくまで手段。しかも消える保証はありません。私たちが目指すのは、★評価そのものを上げ、\nサジェストや検索結果も整えて「選ばれ続けるお店」になること。だから、消えても・消えなくても前に進めます。",
      { x: 0.9, y: 3.3, w: 11.3, h: 1.0, fontFace: BODY, fontSize: 15.5, color: "C7D6EA", lineSpacing: 23 }
    );
    const chip = async (x, ic, label, sub, accent) => {
      s.addShape(p.shapes.ROUNDED_RECTANGLE, { x, y: 4.65, w: 5.5, h: 1.9, fill: { color: "14315A" }, line: { color: accent, width: 1.5 }, rectRadius: 0.1 });
      s.addShape(p.shapes.OVAL, { x: x + 0.35, y: 5.05, w: 0.85, h: 0.85, fill: { color: accent }, line: { type: "none" } });
      s.addImage({ data: await icon(ic, "#FFFFFF"), x: x + 0.55, y: 5.25, w: 0.45, h: 0.45 });
      s.addText(label, { x: x + 1.45, y: 5.02, w: 3.85, h: 0.55, fontFace: HEAD, fontSize: 19, bold: true, color: WHITE, valign: "middle" });
      s.addText(sub, { x: x + 1.45, y: 5.62, w: 3.9, h: 0.8, fontFace: BODY, fontSize: 11.5, color: "AEC2DD", lineSpacing: 15 });
    };
    await chip(0.9, FaTimesCircle, "口コミを消す ＝ 手段", "違反の口コミはAIのサポートでご自身が申請。\n通らなければ行き止まりになる「手段」。", "5C7299");
    await chip(6.9, FaStar, "評判を上げる ＝ ゴール", "良い口コミを集め検索結果も整え、\n★が積み上がり続ける状態へ。", GOLD);
    footer(s);
  }

  // ===================== SLIDE 4 — ALL-IN-ONE PACKAGE =====================
  {
    const s = p.addSlide();
    s.background = { color: BG };
    head(s, "評判改善に必要なものが、ぜんぶ揃った総合パッケージ。", "「対策」も「集客の仕組み」も「守り」も、これひとつ。契約中はずっとご利用いただけます。");

    const items = [
      [FaRobot, BLUE, "AI 口コミ対策サポート", "違反の可能性や申請文の書き方をAIが助言。申請はご自身で。", false],
      [FaBalanceScale, GREEN, "提携弁護士の窓口", "法的対応が必要な事案は提携弁護士へおつなぎ。", false],
      [FaStar, GOLD, "口コミ獲得・評判改善", "レビュー導線づくりとMEOで★が貯まる状態に。", false],
      [FaGlobe, TEAL, "公式WEBサイト", "店舗の魅力を伝える公式サイトを制作・運用。", true],
      [FaCommentDots, LINEG, "公式LINE ＋ マーケ導線", "登録→来店→口コミ依頼までの導線を設計・運用。", true],
      [FaMobileAlt, PURPLE, "店舗アプリ（iOS / Android）", "再来店を促す自社アプリでリピーターを囲い込む。", true],
      [FaRegCreditCard, "475569", "NFC 口コミカード", "かざすだけで口コミツールが起動。投稿率アップ。", false],
      [FaSearch, NAVY, "サジェスト対策", "ネガティブな検索候補の改善に取り組む（上位プラン）。", false],
      [FaArrowDown, RED, "逆SEO 押し下げパック", "ネガティブな表示を相対的に下位へ（上位プラン）。", false],
    ];
    const cols = 3, cw = 3.86, ch = 1.55, gx = 0.26, gy = 0.22, x0 = 0.7, y0 = 1.72;
    for (let i = 0; i < items.length; i++) {
      const [ic, color, title, desc, free] = items[i];
      const col = i % cols, row = Math.floor(i / cols);
      const x = x0 + col * (cw + gx), y = y0 + row * (ch + gy);
      s.addShape(p.shapes.ROUNDED_RECTANGLE, { x, y, w: cw, h: ch, fill: { color: WHITE }, line: { color: LINE, width: 1 }, rectRadius: 0.08, shadow: shSoft() });
      s.addShape(p.shapes.ROUNDED_RECTANGLE, { x: x + 0.26, y: y + 0.26, w: 0.6, h: 0.6, fill: { color }, line: { type: "none" }, rectRadius: 0.12 });
      s.addImage({ data: await icon(ic, "#FFFFFF"), x: x + 0.41, y: y + 0.41, w: 0.3, h: 0.3 });
      const titleW = free ? cw - 2.15 : cw - 1.25;
      s.addText(title, { x: x + 1.0, y: y + 0.24, w: titleW, h: 0.64, fontFace: BODY, fontSize: 13.5, bold: true, color: NAVY, valign: "middle", lineSpacing: 15 });
      if (free) {
        s.addShape(p.shapes.ROUNDED_RECTANGLE, { x: x + cw - 1.12, y: y + 0.26, w: 0.86, h: 0.3, fill: { color: GOLD }, line: { type: "none" }, rectRadius: 0.15 });
        s.addText("契約中 無料", { x: x + cw - 1.12, y: y + 0.26, w: 0.86, h: 0.3, fontFace: BODY, fontSize: 8, bold: true, color: WHITE, align: "center", valign: "middle" });
      }
      s.addText(desc, { x: x + 0.26, y: y + 0.92, w: cw - 0.5, h: 0.5, fontFace: BODY, fontSize: 10.3, color: INKSOFT, lineSpacing: 13 });
    }
    footer(s);
  }

  // ===================== SLIDE 5 — LINE × REVIEW =====================
  {
    const s = p.addSlide();
    s.background = { color: BG };
    head(s, "強みは「LINE × 口コミ」の循環。", "ただLINEを作るだけではありません。来店からリピート、口コミ獲得までを“仕組み”で回します。");

    // green band
    const bx = 0.7, by = 1.95, bw = 11.9, bh = 4.55;
    s.addShape(p.shapes.ROUNDED_RECTANGLE, { x: bx, y: by, w: bw, h: bh, fill: { color: LINEG }, line: { type: "none" }, rectRadius: 0.1, shadow: sh() });
    s.addShape(p.shapes.OVAL, { x: bx + bw - 2.2, y: by - 1.0, w: 3.2, h: 3.2, fill: { color: "FFFFFF", transparency: 90 }, line: { type: "none" } });
    s.addImage({ data: await icon(FaCommentDots, "#FFFFFF"), x: bx + 0.5, y: by + 0.45, w: 0.5, h: 0.5 });
    s.addText("LINEを使った評判アップの循環", { x: bx + 1.15, y: by + 0.4, w: bw - 2, h: 0.6, fontFace: HEAD, fontSize: 22, bold: true, color: WHITE, valign: "middle" });
    s.addText("良い体験を、自然と「★」に変える導線を設計・運用します。", { x: bx + 0.55, y: by + 1.05, w: bw - 1, h: 0.4, fontFace: BODY, fontSize: 13, color: "EAFBF0" });

    const steps = [
      ["STEP 1", "LINE登録を促す", "来店時やWEB・アプリから公式LINEへ。お得情報でファン化。"],
      ["STEP 2", "来店・再来店を促進", "クーポンや通知で再来店を後押し。リピーターを増やす。"],
      ["STEP 3", "満足客に口コミ依頼", "満足したタイミングでLINEから口コミ投稿を自然にお願い。"],
      ["STEP 4", "不満は先に拾う", "不満はLINEで受け止め店内で解決。低評価を未然に防止。"],
    ];
    const sw = 2.72, sgx = 0.18, sx0 = bx + 0.5, sy = by + 1.65, sch = 2.4;
    for (let i = 0; i < steps.length; i++) {
      const x = sx0 + i * (sw + sgx);
      s.addShape(p.shapes.ROUNDED_RECTANGLE, { x, y: sy, w: sw, h: sch, fill: { color: "FFFFFF", transparency: 8 }, line: { color: "FFFFFF", width: 1 }, rectRadius: 0.08 });
      s.addText(steps[i][0], { x: x + 0.25, y: sy + 0.22, w: sw - 0.5, h: 0.3, fontFace: HEAD, fontSize: 11, bold: true, color: "DFF7E8" });
      s.addText(steps[i][1], { x: x + 0.25, y: sy + 0.58, w: sw - 0.5, h: 0.55, fontFace: BODY, fontSize: 14.5, bold: true, color: WHITE, lineSpacing: 17 });
      s.addText(steps[i][2], { x: x + 0.25, y: sy + 1.2, w: sw - 0.5, h: 1.0, fontFace: BODY, fontSize: 11, color: "EAFBF0", lineSpacing: 14 });
      if (i < steps.length - 1) {
        s.addImage({ data: await icon(FaArrowRight, "#FFFFFF"), x: x + sw + sgx / 2 - 0.13, y: sy + 0.95, w: 0.26, h: 0.26 });
      }
    }
    footer(s);
  }

  // ===================== SLIDE — AI CONCIERGE =====================
  {
    const s = p.addSlide();
    s.background = { color: "0E2A4A" };
    s.addShape(p.shapes.OVAL, { x: 9.4, y: -2.4, w: 7.5, h: 7.5, fill: { color: TEAL, transparency: 55 }, line: { type: "none" } });

    s.addText("AI CONCIERGE ｜ 全プラン共通", { x: 0.9, y: 0.85, w: 11, h: 0.4, fontFace: BODY, fontSize: 14, color: "5EEAD4", charSpacing: 2, bold: true });
    s.addText([
      { text: "困ったら、いつでもAIに相談。", options: { color: WHITE, breakLine: true } },
      { text: "専門知識ゼロでも、迷わない。", options: { color: "5EEAD4" } },
    ], { x: 0.85, y: 1.4, w: 7.2, h: 1.7, fontFace: HEAD, fontSize: 30, bold: true, lineSpacing: 42 });
    s.addText(
      "ご契約者さまには、24時間つかえる専用のAIサポートをご用意。口コミへの返信文、申請の進め方、次の打ち手まで、チャットで質問すればすぐに分かりやすく答えます。",
      { x: 0.9, y: 3.15, w: 6.9, h: 1.2, fontFace: BODY, fontSize: 14, color: "C7D6EA", lineSpacing: 21 }
    );
    const pts = [
      "全プラン共通・追加料金なしでご利用可能",
      "口コミ返信の文案づくり・申請のやり方も即サポート",
      "難しい操作や専門用語も、かみくだいて説明",
    ];
    let py = 4.5;
    for (const t of pts) {
      s.addShape(p.shapes.OVAL, { x: 0.92, y: py, w: 0.34, h: 0.34, fill: { color: "164E63" }, line: { color: "5EEAD4", width: 1 } });
      s.addImage({ data: await icon(FaCheckCircle, "#5EEAD4"), x: 0.99, y: py + 0.07, w: 0.2, h: 0.2 });
      s.addText(t, { x: 1.4, y: py - 0.07, w: 6.5, h: 0.5, fontFace: BODY, fontSize: 13.5, color: "EAF1FA", valign: "middle" });
      py += 0.62;
    }

    // chat mockup card (right)
    const cx = 8.35, cy = 1.35, cw2 = 4.25, ch2 = 5.3;
    s.addShape(p.shapes.ROUNDED_RECTANGLE, { x: cx, y: cy, w: cw2, h: ch2, fill: { color: "12345A" }, line: { color: "2C5E7A", width: 1 }, rectRadius: 0.1, shadow: sh() });
    // head
    s.addShape(p.shapes.OVAL, { x: cx + 0.35, y: cy + 0.35, w: 0.6, h: 0.6, fill: { color: TEAL }, line: { type: "none" } });
    s.addImage({ data: await icon(FaRobot, "#FFFFFF"), x: cx + 0.5, y: cy + 0.5, w: 0.3, h: 0.3 });
    s.addText("レビューブースト AIサポート", { x: cx + 1.05, y: cy + 0.34, w: cw2 - 1.3, h: 0.32, fontFace: BODY, fontSize: 12, bold: true, color: WHITE, valign: "middle" });
    s.addText("● オンライン｜24時間対応", { x: cx + 1.05, y: cy + 0.66, w: cw2 - 1.3, h: 0.26, fontFace: BODY, fontSize: 9, color: "5EEAD4", valign: "middle" });
    s.addShape(p.shapes.LINE, { x: cx + 0.3, y: cy + 1.12, w: cw2 - 0.6, h: 0, line: { color: "2C5E7A", width: 1 } });

    // bubbles
    const userBubble = (y, text, h) => {
      s.addShape(p.shapes.ROUNDED_RECTANGLE, { x: cx + 1.1, y, w: cw2 - 1.45, h, fill: { color: BLUE }, line: { type: "none" }, rectRadius: 0.12 });
      s.addText(text, { x: cx + 1.3, y: y + 0.08, w: cw2 - 1.8, h: h - 0.16, fontFace: BODY, fontSize: 10.5, color: WHITE, valign: "middle", lineSpacing: 13 });
    };
    const aiBubble = (y, runs, h) => {
      s.addShape(p.shapes.ROUNDED_RECTANGLE, { x: cx + 0.35, y, w: cw2 - 1.45, h, fill: { color: "EAF1FA" }, line: { type: "none" }, rectRadius: 0.12 });
      s.addText(runs, { x: cx + 0.55, y: y + 0.08, w: cw2 - 1.8, h: h - 0.16, fontFace: BODY, fontSize: 10.5, color: INK, valign: "middle", lineSpacing: 13 });
    };
    userBubble(cy + 1.32, "★1で「対応が遅い」と書かれました。どう返信すれば？", 0.7);
    aiBubble(cy + 2.16, [
      { text: "まず", options: {} },
      { text: "感謝とお詫び", options: { bold: true, color: TEAL } },
      { text: "を伝え、改善策を簡潔に。返信文の例を3パターン作りますね。", options: {} },
    ], 0.92);
    userBubble(cy + 3.22, "この口コミ、ガイドライン違反で消せそう？", 0.55);
    aiBubble(cy + 3.92, [
      { text: "第三者の投稿の可能性があり、", options: {} },
      { text: "違反として申請できる余地", options: { bold: true, color: TEAL } },
      { text: "が。手順をご案内します（申請はご自身で）。", options: {} },
    ], 0.95);

    footer(s);
  }

  // ===================== SLIDE 6 — PRICING (3 plans) =====================
  {
    const s = p.addSlide();
    s.background = { color: BG };
    head(s, "料金プラン", "守備範囲に合わせて3プラン。束ねるほど割安です。まずは無料診断からどうぞ。");

    const plans = [
      {
        name: "Light", for: "まずは評判改善と集客基盤から", price: "3", unit: "万円／月〜", featured: false,
        feats: [["c", "契約者専用 AIサポート（24時間）"], ["c", "口コミ獲得・評判改善コンサル"], ["c", "Googleプロフィール最適化（MEO）"], ["c", "公式WEB（契約中 無料）"], ["c", "公式LINE＋マーケ導線（契約中 無料）"]],
      },
      {
        name: "Standard", for: "口コミ対策と集客をまとめて強化", price: "6", unit: "万円／月〜", featured: true,
        feats: [["h", "Lightの内容すべて"], ["c", "AI 口コミ対策サポート"], ["c", "店舗アプリ iOS/Android（契約中 無料）"], ["c", "LINEマーケ導線の運用代行"], ["c", "提携弁護士の窓口ご案内"]],
      },
      {
        name: "Pro", for: "ネット評判リスクをまるごと防衛", price: "要", unit: "お見積り", featured: false,
        feats: [["h", "Standardの内容すべて"], ["c", "サジェスト対策"], ["c", "逆SEO 押し下げパック"], ["c", "提携弁護士の優先窓口"], ["c", "専任担当の手厚いサポート"]],
      },
    ];
    const cw = 3.95, gx = 0.33, x0 = 0.72, y0 = 2.0;
    for (let i = 0; i < plans.length; i++) {
      const pl = plans[i];
      const x = x0 + i * (cw + gx);
      const ch = pl.featured ? 4.7 : 4.4;
      const y = pl.featured ? y0 - 0.12 : y0;
      s.addShape(p.shapes.ROUNDED_RECTANGLE, { x, y, w: cw, h: ch, fill: { color: WHITE }, line: { color: pl.featured ? GOLD : LINE, width: pl.featured ? 2.5 : 1 }, rectRadius: 0.08, shadow: pl.featured ? sh() : shSoft() });
      if (pl.featured) {
        s.addShape(p.shapes.ROUNDED_RECTANGLE, { x: x + cw / 2 - 1.45, y: y - 0.22, w: 2.9, h: 0.46, fill: { color: GOLD }, line: { type: "none" }, rectRadius: 0.23, shadow: shSoft() });
        s.addText("★ いちばん選ばれています", { x: x + cw / 2 - 1.45, y: y - 0.22, w: 2.9, h: 0.46, fontFace: BODY, fontSize: 11, bold: true, color: WHITE, align: "center", valign: "middle" });
      }
      const pad = 0.4;
      s.addText(pl.name, { x: x + pad, y: y + 0.38, w: cw - pad * 2, h: 0.55, fontFace: HEAD, fontSize: 22, bold: true, color: NAVY });
      s.addText(pl.for, { x: x + pad, y: y + 0.95, w: cw - pad * 2, h: 0.45, fontFace: BODY, fontSize: 11, color: INKSOFT, lineSpacing: 14 });
      s.addText([
        { text: pl.price, options: { fontSize: 36, bold: true, color: NAVY, fontFace: HEAD } },
        { text: " " + pl.unit, options: { fontSize: 13, bold: true, color: INKSOFT, fontFace: BODY } },
      ], { x: x + pad, y: y + 1.45, w: cw - pad * 2, h: 0.6, align: "left", valign: "middle" });
      s.addText("（税別・契約期間中）", { x: x + pad, y: y + 2.04, w: cw - pad * 2, h: 0.25, fontFace: BODY, fontSize: 9.5, color: INKSOFT });
      s.addShape(p.shapes.LINE, { x: x + pad, y: y + 2.36, w: cw - pad * 2, h: 0, line: { color: LINE, width: 1 } });
      let fy = y + 2.55;
      for (const [type, txt] of pl.feats) {
        if (type === "h") {
          s.addImage({ data: await icon(FaArrowDown, "#" + BLUE), x: x + pad, y: fy + 0.02, w: 0.2, h: 0.2 });
          s.addText(txt, { x: x + pad + 0.32, y: fy - 0.06, w: cw - pad * 2 - 0.32, h: 0.4, fontFace: BODY, fontSize: 11.5, bold: true, color: NAVY });
        } else {
          s.addImage({ data: await icon(FaCheckCircle, "#" + GREEN), x: x + pad, y: fy + 0.02, w: 0.2, h: 0.2 });
          s.addText(txt, { x: x + pad + 0.32, y: fy - 0.06, w: cw - pad * 2 - 0.32, h: 0.42, fontFace: BODY, fontSize: 11, color: INK, lineSpacing: 13 });
        }
        fy += 0.4;
      }
    }
    s.addText("※ 金額は目安です。店舗数・口コミの状況により最適なプランをご提案します。サジェスト・逆SEO・口コミ削除はいずれも結果を保証するものではありません。", {
      x: 0.72, y: 6.66, w: 11.5, h: 0.32, fontFace: BODY, fontSize: 9.5, color: INKSOFT,
    });
    footer(s);
  }

  // ===================== SLIDE 7 — HOW IT WORKS =====================
  {
    const s = p.addSlide();
    s.background = { color: BG };
    head(s, "ご相談から改善までの流れ", "まずは無料診断から。継続的な改善で「選ばれ続けるお店」へ。");
    const steps = [
      [FaSearch, "無料の評判診断", "口コミ・★評価・サジェスト・検索結果を診断。改善点を無料でお伝えします。", BLUE],
      [FaHandshake, "プランのご提案", "お店の状況とご予算に合わせ最適なプランをご提案。ご納得後にスタート。", BLUE],
      [FaBullhorn, "基盤づくり＆対策開始", "WEB・LINE・アプリを整え、口コミ獲得の仕組みを導入。各窓口も利用開始。", BLUE],
      [FaChartLine, "毎月の改善サポート", "★評価や評判の推移をレポート。お役立ち情報の配信で評判を磨き続けます。", GOLDD],
    ];
    const cw = 2.95, gx = 0.28, x0 = 0.7, y0 = 2.05, ch = 4.0;
    for (let i = 0; i < steps.length; i++) {
      const x = x0 + i * (cw + gx);
      const accent = steps[i][3];
      s.addShape(p.shapes.ROUNDED_RECTANGLE, { x, y: y0, w: cw, h: ch, fill: { color: WHITE }, line: { color: LINE, width: 1 }, rectRadius: 0.07, shadow: shSoft() });
      s.addShape(p.shapes.RECTANGLE, { x, y: y0, w: cw, h: 0.13, fill: { color: accent }, line: { type: "none" } });
      s.addShape(p.shapes.OVAL, { x: x + cw / 2 - 0.42, y: y0 + 0.45, w: 0.84, h: 0.84, fill: { color: accent }, line: { type: "none" }, shadow: shSoft() });
      s.addText(String(i + 1), { x: x + cw / 2 - 0.42, y: y0 + 0.45, w: 0.84, h: 0.84, fontFace: HEAD, fontSize: 26, bold: true, color: WHITE, align: "center", valign: "middle" });
      s.addImage({ data: await icon(steps[i][0], "#" + accent), x: x + cw / 2 - 0.32, y: y0 + 1.55, w: 0.64, h: 0.64 });
      s.addText(steps[i][1], { x: x + 0.18, y: y0 + 2.35, w: cw - 0.36, h: 0.6, fontFace: BODY, fontSize: 14.5, bold: true, color: NAVY, align: "center", valign: "middle" });
      s.addText(steps[i][2], { x: x + 0.22, y: y0 + 2.95, w: cw - 0.44, h: 1.0, fontFace: BODY, fontSize: 11, color: INKSOFT, align: "center", lineSpacing: 14 });
      if (i < steps.length - 1) s.addImage({ data: await icon(FaArrowRight, "#C3D0DE"), x: x + cw + gx / 2 - 0.17, y: y0 + 1.65, w: 0.34, h: 0.34 });
    }
    footer(s);
  }

  // ===================== SLIDE 8 — COMPLIANCE / 安心の建て付け =====================
  {
    const s = p.addSlide();
    s.background = { color: BG };
    head(s, "安心してお任せいただける「役割分担」", "私たちは「削除の代行」はいたしません。役割を分けることで、安心・適法にご支援します。");

    const cards = [
      [FaRobot, BLUE, "AIは“やり方”を支援", "ガイドライン違反の可能性や申請フォームの書き方をAIが助言。実際の削除申請はお客様ご自身で行っていただきます。"],
      [FaBalanceScale, GREEN, "法的対応は提携弁護士", "権利侵害や悪質な誹謗中傷など、法律事務が必要な事案は提携弁護士の窓口へ。専門家が対応します。"],
      [FaStar, GOLD, "私たちは評判改善に専念", "口コミ獲得・MEO・WEB・LINE・アプリなど、評判を上げる施策に集中。長期的に「強いお店」をつくります。"],
    ];
    const cw = 3.95, gx = 0.33, x0 = 0.72, y = 2.05, ch = 3.0;
    for (let i = 0; i < cards.length; i++) {
      const [ic, color, title, desc] = cards[i];
      const x = x0 + i * (cw + gx);
      s.addShape(p.shapes.ROUNDED_RECTANGLE, { x, y, w: cw, h: ch, fill: { color: WHITE }, line: { color: LINE, width: 1 }, rectRadius: 0.08, shadow: shSoft() });
      s.addShape(p.shapes.OVAL, { x: x + 0.4, y: y + 0.4, w: 0.95, h: 0.95, fill: { color }, line: { type: "none" } });
      s.addImage({ data: await icon(ic, "#FFFFFF"), x: x + 0.66, y: y + 0.66, w: 0.43, h: 0.43 });
      s.addText(title, { x: x + 0.4, y: y + 1.55, w: cw - 0.8, h: 0.55, fontFace: HEAD, fontSize: 17, bold: true, color: NAVY });
      s.addText(desc, { x: x + 0.4, y: y + 2.1, w: cw - 0.8, h: 0.85, fontFace: BODY, fontSize: 11.5, color: INKSOFT, lineSpacing: 15 });
    }
    s.addShape(p.shapes.ROUNDED_RECTANGLE, { x: 0.72, y: 5.35, w: 11.85, h: 1.15, fill: { color: BGALT }, line: { color: "C9DBF1", width: 1, dashType: "dash" }, rectRadius: 0.1 });
    s.addImage({ data: await icon(FaShieldAlt, "#" + BLUE), x: 1.05, y: 5.68, w: 0.5, h: 0.5 });
    s.addText("当社は口コミの削除申請を代行するものではなく、お客様ご自身による申請をAI等で支援するサービスです。法的対応が必要な事案は提携弁護士をご案内します。口コミ削除・サジェスト改善・検索順位はGoogle等の判断や仕様によるもので、結果を保証するものではありません。", {
      x: 1.7, y: 5.5, w: 10.6, h: 0.9, fontFace: BODY, fontSize: 10.5, color: INKSOFT, valign: "middle", lineSpacing: 14,
    });
    footer(s);
  }

  // ===================== SLIDE 9 — CLOSING / CTA =====================
  {
    const s = p.addSlide();
    s.background = { color: NAVY };
    s.addShape(p.shapes.OVAL, { x: 8.8, y: 2.6, w: 8, h: 8, fill: { color: BLUE, transparency: 70 }, line: { type: "none" } });
    await starRow(s, 5.55, 1.2, 5, 5, 0.5, true);
    s.addText("まずは「無料の評判診断」から。", { x: 1, y: 2.2, w: 11.3, h: 1.0, fontFace: HEAD, fontSize: 34, bold: true, color: WHITE, align: "center" });
    s.addText("しつこい営業はいたしません。口コミ・サジェスト・検索結果の状況を、無料で診断いたします。\n「消す」ためではなく、「選ばれ続ける」ために。一緒に評判を育てましょう。", {
      x: 1.5, y: 3.3, w: 10.3, h: 1.0, fontFace: BODY, fontSize: 16, color: "C7D6EA", align: "center", lineSpacing: 24,
    });
    s.addShape(p.shapes.ROUNDED_RECTANGLE, { x: 4.55, y: 4.6, w: 4.2, h: 0.85, fill: { color: GOLD }, line: { type: "none" }, rectRadius: 0.42, shadow: sh() });
    s.addText("★ 無料で評判診断を申し込む", { x: 4.55, y: 4.6, w: 4.2, h: 0.85, fontFace: BODY, fontSize: 16, bold: true, color: WHITE, align: "center", valign: "middle" });
    s.addText([
      { text: "お問い合わせ  ", options: { color: "8FB3E6", bold: true } },
      { text: "レビューブースト", options: { color: WHITE, bold: true } },
    ], { x: 1, y: 5.85, w: 11.3, h: 0.4, fontFace: BODY, fontSize: 14, align: "center" });
    s.addText("※ 当社は口コミの削除申請を代行するものではなく、ご自身の申請をAI等で支援するサービスです。法的対応が必要な事案は提携弁護士をご案内します。各施策の結果を保証するものではありません。Google LLCとは関係ありません。", {
      x: 1.2, y: 6.75, w: 10.9, h: 0.6, fontFace: BODY, fontSize: 8.5, color: "5E7497", align: "center", lineSpacing: 12,
    });
  }

  await p.writeFile({ fileName: "Google評判総合コンサルパック_提案書.pptx" });
  console.log("✓ deck written");
})();
