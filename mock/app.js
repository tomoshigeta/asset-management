/* ============================================================
   共通ロジック — 集計・整形・描画ヘルパ
   ============================================================ */
(function () {
  "use strict";
  /* ---------- JSON の検査 ----------
     画面が壊れる前に、初心者でも直せる日本語で指摘する。
     形式の説明は data/README.md。 */
  const isNum = v => typeof v === "number" && isFinite(v);
  const isStr = v => typeof v === "string";
  const esc = s => String(s).replace(/[&<>"]/g, c =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

  function validateData(d) {
    const errs = [];
    const fmt = n => Math.round(n).toLocaleString("ja-JP");
    const need = (obj, key, type, where) => {
      const v = obj == null ? undefined : obj[key];
      const ok = type === "number" ? isNum(v) : type === "string" ? isStr(v) : Array.isArray(v);
      if (!ok) {
        const why = v === undefined ? "ありません"
          : type === "number" ? '数字ではありません（" で囲まず、カンマも入れずに書きます）'
          : type === "string" ? '文字ではありません（" で囲みます）'
          : "[ ] の形になっていません";
        errs.push(where + " に「" + key + "」が" + why);
      }
      return ok;
    };
    if (!d || typeof d !== "object" || Array.isArray(d)) return ["ファイル全体が { } で囲まれていません"];
    need(d.owner, "name", "string", "owner");
    need(d, "asOf", "string", "一番外側");
    need(d.rates, "JPY", "number", "rates");
    ["realEstate", "securities", "deposits", "otherLoans"].forEach(k => need(d, k, "array", "一番外側"));
    if (errs.length) return errs;
    const rateOK = cur => isStr(cur) && isNum(d.rates[cur]);
    const rateErr = (w, cur) => errs.push(w + ": currency「" + cur + "」の為替レートが rates にありません");

    d.realEstate.forEach((p, i) => {
      const w = "realEstate の " + (i + 1) + " 件目" + (p && isStr(p.name) ? "（" + p.name + "）" : "");
      need(p, "name", "string", w);
      need(p, "areaSqm", "number", w);
      need(p, "ageYears", "number", w);
      /* 時価は任意（一意に決まらないため必須にしない）。
         入っているときだけ、物件別の含み損益と LTV を出す。 */
      if (p && p.marketValue !== undefined && !isNum(p.marketValue)) {
        errs.push(w + " の marketValue（時価）が数字ではありません");
      }
      const okInv = need(p, "totalInvestment", "number", w);
      const okOwn = need(p, "ownFunds", "number", w);
      if (!p || !p.loan || typeof p.loan !== "object") { errs.push(w + " に「loan」がありません"); return; }
      const L = p.loan, wl = w + " の loan";
      need(L, "lender", "string", wl);
      const okPr = need(L, "principal", "number", wl);
      const okBa = need(L, "balance", "number", wl);
      const okRe = need(L, "cumulativeRepaid", "number", wl);
      need(L, "rate", "number", wl);
      if (L.termYears !== undefined && L.termMonths === undefined) {
        errs.push(w + ": loan.termYears は loan.termMonths（月数）に変わりました" +
                  "（形式2）。Excel の入力シートから作り直してください");
      } else {
        need(L, "termMonths", "number", wl);
      }
      need(L, "monthlyRepayment", "number", wl);
      if (okPr && okBa && okRe && Math.round(L.balance + L.cumulativeRepaid) !== Math.round(L.principal))
        errs.push(w + ": loan.principal（" + fmt(L.principal) + "）が balance + cumulativeRepaid（" +
          fmt(L.balance + L.cumulativeRepaid) + "）と一致しません");
      if (okInv && okOwn && okPr && Math.round(p.ownFunds + L.principal) !== Math.round(p.totalInvestment))
        errs.push(w + ": totalInvestment（" + fmt(p.totalInvestment) + "）が ownFunds + loan.principal（" +
          fmt(p.ownFunds + L.principal) + "）と一致しません");
      if (!p.monthly || typeof p.monthly !== "object") { errs.push(w + " に「monthly」がありません"); return; }
      ["rentIncome", "managementFee", "repairReserve", "propertyTaxMonthly"]
        .forEach(k => need(p.monthly, k, "number", w + " の monthly"));
    });

    const REGIONS = ["domestic", "overseas"], CLASSES = ["株", "債券", "投資信託", "その他"];
    d.securities.forEach((s, i) => {
      const w = "securities の " + (i + 1) + " 件目";
      if (!s || REGIONS.indexOf(s.region) < 0) errs.push(w + ': region は "domestic" か "overseas" です');
      if (!s || CLASSES.indexOf(s.assetClass) < 0) errs.push(w + ': assetClass は "株" "債券" "投資信託" "その他" のどれかです');
      need(s, "cost", "number", w);
      need(s, "marketValue", "number", w);
      if (!rateOK(s && s.currency)) rateErr(w, s && s.currency);
    });
    d.deposits.forEach((x, i) => {
      const w = "deposits の " + (i + 1) + " 件目" + (x && isStr(x.bank) ? "（" + x.bank + "）" : "");
      need(x, "bank", "string", w);
      need(x, "amount", "number", w);
      if (!rateOK(x && x.currency)) rateErr(w, x && x.currency);
    });
    /* 保険は無くてもよい（持っていない人がいるため）。あれば中身を検査する。 */
    if (d.insurance !== undefined) {
      if (!Array.isArray(d.insurance)) {
        errs.push("一番外側の「insurance」が [ ] の形になっていません");
      } else {
        d.insurance.forEach((x, i) => {
          const w = "insurance の " + (i + 1) + " 件目" +
                    (x && isStr(x.name) ? "（" + x.name + "）" : "");
          need(x, "name", "string", w);
          need(x, "premiumPaid", "number", w);
          need(x, "surrenderValue", "number", w);
          if (!rateOK(x && x.currency)) rateErr(w, x && x.currency);
          if (x && x.deathBenefit !== undefined && !isNum(x.deathBenefit)) {
            errs.push(w + " の deathBenefit が数字ではありません");
          }
        });
      }
    }

    d.otherLoans.forEach((l, i) => {
      const w = "otherLoans の " + (i + 1) + " 件目" + (l && isStr(l.name) ? "（" + l.name + "）" : "");
      need(l, "name", "string", w);
      need(l, "lender", "string", w);
      const okPr = need(l, "principal", "number", w);
      const okBa = need(l, "balance", "number", w);
      need(l, "rate", "number", w);
      need(l, "monthlyRepayment", "number", w);
      if (okPr && okBa && l.balance > l.principal)
        errs.push(w + ": balance（残債）が principal（当初借入額）より大きくなっています");
    });
    return errs;
  }

  /* JSON.parse が失敗したとき、初心者がやりがちな原因を推定して添える */
  function parseJSONText(text) {
    text = String(text).replace(/^\uFEFF/, "");
    try { return { data: JSON.parse(text) }; }
    catch (e) {
      const outside = text.replace(/"(?:[^"\\]|\\.)*"/g, '""');   // " " の中を除いた部分
      const hints = [];
      if (/[０-９．，：｛｝［］“”＂]/.test(text)) hints.push("全角の数字や記号（０１２ ， ． “ ”）が混ざっています。半角に直してください");
      if (/\d,\d{3}(?!\d)/.test(outside)) hints.push("数字の中のカンマ（1,000,000）は使えません。1000000 と書きます");
      if (/,\s*[\]}]/.test(outside)) hints.push("最後の項目のあとに余分な , があります。消してください");
      if (/\/\/|\/\*/.test(outside)) hints.push("JSON にはメモ（// や /* */）を書けません。消してください");
      if (/'/.test(outside)) hints.push("文字は ' ではなく \" で囲みます");
      hints.push("ブラウザの報告: " + e.message);
      return { errors: ["JSON として読めませんでした"].concat(hints) };
    }
  }

  /* ---------- 表示するデータの決定 ----------
     1. 「JSONファイルを読み込む」で取り込んだ JSON（ブラウザに控えが残る）
     2. なければ data.js のサンプル
     正はあくまで手元の JSON ファイル（仕様: 手書きの単一ファイルが正）。
     ブラウザ内の控えは「サンプルに戻す」で消える。 */
  const STORE_KEY  = "assetDashboard.json";
  const STORE_NAME = "assetDashboard.fileName";
  function storageGet(k)    { try { return localStorage.getItem(k); } catch (e) { return null; } }
  function storageSet(k, v) { try { localStorage.setItem(k, v); return true; } catch (e) { return false; } }
  function storageDel(k)    { try { localStorage.removeItem(k); } catch (e) {} }

  const SOURCE = { kind: "sample", label: "サンプル（data.js）" };
  let D = window.ASSET_DATA;
  (function () {
    const text = storageGet(STORE_KEY);
    if (!text) return;
    try {
      const parsed = JSON.parse(text);
      if (validateData(parsed).length) throw new Error("stored data is invalid");
      D = parsed;
      SOURCE.kind = "json";
      SOURCE.label = storageGet(STORE_NAME) || "JSON ファイル";
    } catch (e) {
      storageDel(STORE_KEY); storageDel(STORE_NAME);
    }
  })();

  /* ---------- JSON ファイルの読み込み ---------- */
  function showReport(fileName, errs) {
    let box = document.getElementById("load-report");
    if (!box) { box = document.createElement("div"); box.id = "load-report"; document.body.appendChild(box); }
    box.innerHTML =
      '<div class="load-report-head"><b>' + esc(fileName) + "</b> は読み込めませんでした" +
        '<button type="button" class="load-report-close" aria-label="閉じる">×</button></div>' +
      "<p>下の点を直して、もう一度「JSONファイルを読み込む」を押してください。" +
        "書き方は <code>data/README.md</code> にあります。</p>" +
      "<ol>" + errs.map(m => "<li>" + esc(m) + "</li>").join("") + "</ol>";
    box.querySelector(".load-report-close").onclick = () => box.remove();
  }
  function loadJSONFile(file) {
    const reader = new FileReader();
    reader.onload = () => {
      const r = parseJSONText(reader.result);
      const errs = r.errors || validateData(r.data);
      if (errs.length) { showReport(file.name, errs); return; }
      if (!storageSet(STORE_KEY, JSON.stringify(r.data))) {
        showReport(file.name, ["ブラウザに保存できませんでした（シークレットモードでは使えないことがあります）"]);
        return;
      }
      storageSet(STORE_NAME, file.name);
      location.reload();
    };
    reader.onerror = () => showReport(file.name, ["ファイルを開けませんでした"]);
    reader.readAsText(file, "utf-8");
  }
  document.addEventListener("change", e => {
    const input = e.target;
    if (!input || input.id !== "json-file") return;
    const f = input.files && input.files[0];
    if (f) loadJSONFile(f);
    input.value = "";
  });
  document.addEventListener("click", e => {
    if (e.target.closest && e.target.closest("#reset-sample")) {
      storageDel(STORE_KEY); storageDel(STORE_NAME);
      location.reload();
    }
  });

  /* ---------- 整形 ---------- */
  const yen = n => "¥ " + Math.round(n).toLocaleString("ja-JP");
  const plain = n => Math.round(n).toLocaleString("ja-JP");

  /* 損益の整形（日本式）: 益 = "+123" / 損 = "▲123"（▲は会計慣行のマイナス） */
  function pl(n) {
    const v = Math.round(n);
    if (v === 0) return { text: "±0", cls: "" };
    return v > 0
      ? { text: "+" + v.toLocaleString("ja-JP"), cls: "gain" }
      : { text: "▲" + Math.abs(v).toLocaleString("ja-JP"), cls: "loss" };
  }
  function plHTML(n, suffix) {
    const p = pl(n);
    return '<span class="' + p.cls + ' num">' + p.text + (suffix || "") + "</span>";
  }

  const pct = (a, b) => b === 0 ? "0.0%" : (a / b * 100).toFixed(1) + "%";

  /* 借入期間は月数で持ち、読むときだけ年と月に割る（323 → 26年11か月）。
     小数の年で持つと端数が何か月なのか復元できないため。 */
  function fmtTerm(months) {
    const m = Math.round(months);
    if (!isFinite(m) || m <= 0) return "\u2014";
    const y = Math.floor(m / 12), rest = m % 12;
    if (y && rest) return y + "年" + rest + "か月";
    if (y) return y + "年";
    return rest + "か月";
  }

  /* 損益率も日本式で表す（+x.x% / ▲x.x%）。文字色は損益色。 */
  function plPctHTML(n, base) {
    if (!base) return '<span class="num">\u2014</span>';
    const v = n / base * 100;
    const s = Math.abs(v).toFixed(1) + "%";
    if (Math.round(v * 10) === 0) return '<span class="num">\u00b10.0%</span>';
    return v > 0
      ? '<span class="gain num">+' + s + "</span>"
      : '<span class="loss num">\u25b2' + s + "</span>";
  }
  const rate = cur => D.rates[cur] != null ? D.rates[cur] : 1;

  /* ---------- 集計 ---------- */
  const sum = (arr, f) => arr.reduce((s, x) => s + f(x), 0);

  const RE = {
    totalInvestment:  sum(D.realEstate, p => p.totalInvestment),
    ownFunds:         sum(D.realEstate, p => p.ownFunds),
    principal:        sum(D.realEstate, p => p.loan.principal),
    balance:          sum(D.realEstate, p => p.loan.balance),
    cumulativeRepaid: sum(D.realEstate, p => p.loan.cumulativeRepaid)
  };
  RE.progress = RE.cumulativeRepaid / RE.principal;

  /* ---------- 不動産の時価 ----------
     時価は一意に決まらないので任意入力。入っている物件だけ含み損益を出す。
     総資産・純資産は「全物件に時価が入っているとき」だけ計算する。
     1件でも欠けたまま引き算すると、資産側だけが小さく出て純資産を過小評価するため
     （SPEC が当初 純資産を却下した理由そのもの）。 */
  const hasMV = p => isNum(p.marketValue);
  RE.valuedCount  = D.realEstate.filter(hasMV).length;
  RE.missingCount = D.realEstate.length - RE.valuedCount;
  RE.allValued    = RE.missingCount === 0;
  RE.marketValue  = sum(D.realEstate.filter(hasMV), p => p.marketValue);
  RE.pl  = RE.allValued ? RE.marketValue - RE.totalInvestment : null;
  RE.ltv = RE.allValued && RE.marketValue > 0 ? RE.balance / RE.marketValue : null;

  /* 物件ごとの含み損益と LTV（時価が入っている物件のみ） */
  function equity(p) {
    if (!hasMV(p)) return null;
    return {
      marketValue: p.marketValue,
      pl: p.marketValue - p.totalInvestment,
      ltv: p.marketValue > 0 ? p.loan.balance / p.marketValue : null,
      ownEquity: p.marketValue - p.loan.balance      /* 時価のうち自分の持ち分 */
    };
  }

  /* 月次収支: 家賃 − ローン返済 − 経費（管理費・修繕積立・固都税） */
  function cashflow(p) {
    const m = p.monthly;
    const expenses = m.managementFee + m.repairReserve + m.propertyTaxMonthly;
    return {
      rentIncome: m.rentIncome,
      repayment: p.loan.monthlyRepayment,
      managementFee: m.managementFee,
      repairReserve: m.repairReserve,
      propertyTax: m.propertyTaxMonthly,
      expenses,
      net: m.rentIncome - p.loan.monthlyRepayment - expenses
    };
  }
  const CF_TOTAL = D.realEstate.map(cashflow).reduce((a, c) => ({
    rentIncome: a.rentIncome + c.rentIncome,
    repayment:  a.repayment  + c.repayment,
    expenses:   a.expenses   + c.expenses,
    net:        a.net        + c.net
  }), { rentIncome: 0, repayment: 0, expenses: 0, net: 0 });

  /* 有価証券を円換算 */
  const SEC = D.securities.map(s => {
    const r = rate(s.currency);
    return Object.assign({}, s, {
      costJPY: s.cost * r,
      valueJPY: s.marketValue * r,
      plNative: s.marketValue - s.cost,
      plJPY: (s.marketValue - s.cost) * r
    });
  });

  /* 保険。評価額は解約返戻金、取得原価は払込保険料累計。
     死亡保険金額は「今使える金額」ではないので、資産の合計には一切足さない。 */
  const INS = (D.insurance || []).map(x => {
    const r = rate(x.currency);
    return Object.assign({}, x, {
      costJPY: x.premiumPaid * r,
      valueJPY: x.surrenderValue * r,
      plJPY: (x.surrenderValue - x.premiumPaid) * r
    });
  });

  const SEC_VALUE = sum(SEC, s => s.valueJPY);
  const SEC_COST  = sum(SEC, s => s.costJPY);
  const SEC_PL    = SEC_VALUE - SEC_COST;
  const DEPOSITS  = sum(D.deposits, d => d.amount * rate(d.currency));

  const INS_VALUE = sum(INS, x => x.valueJPY);
  const INS_COST  = sum(INS, x => x.costJPY);
  const INS_PL    = INS_VALUE - INS_COST;

  /* 金融資産 = 有価証券評価額 + 預貯金 + 保険の解約返戻金（不動産を含まない / Q24）。
     保険は当初 有価証券の「その他」に混ぜていたが、現物資産という定義に合わず
     国内/海外の軸も意味をなさないため独立させた。合計に入る点は変わらない。 */
  const FINANCIAL_ASSETS = SEC_VALUE + DEPOSITS + INS_VALUE;

  /* 金融資産全体の含み損益。合計と範囲をそろえるため保険も含める。
     保険の損益には保障の対価が混ざっている（画面で注記する）。 */
  const FIN_PL = SEC_PL + INS_PL;

  /* 総負債 = 不動産ローン残債 + その他ローン残債（Q24） */
  const OTHER_DEBT = sum(D.otherLoans, l => l.balance);
  const TOTAL_DEBT = RE.balance + OTHER_DEBT;

  /* 総資産・純資産は、全物件に時価が入っているときだけ出す。
     欠けているときは null にして、画面には「時価未入力 N件」と表示する。 */
  const TOTAL_ASSETS = RE.allValued ? FINANCIAL_ASSETS + RE.marketValue : null;
  const NET_WORTH    = RE.allValued ? TOTAL_ASSETS - TOTAL_DEBT : null;

  /* ---------- 描画ヘルパ ---------- */

  /** 積み上げ横棒 1本。segs = [{cls, value, label}] */
  function stackedBar(segs, scaleMax) {
    const total = sum(segs, s => s.value);
    const max = scaleMax || total;
    const html = segs.filter(s => s.value > 0).map(s =>
      '<div class="seg ' + s.cls + '" style="flex:' + s.value + ' 0 0"' +
      ' data-tip="' + s.label + '<br><b>' + yen(s.value) + '</b>（' + pct(s.value, total) + '）"></div>'
    ).join("");
    const restW = max - total;
    const rest = restW > 0 ? '<div class="seg seg-rest" style="flex:' + restW + ' 0 0"></div>' : "";
    return '<div class="bar-track">' + html + rest + "</div>";
  }

  /** 単色横棒（預貯金など、損益の概念がないもの） */
  function valueBar(value, max, tip) {
    const w = max === 0 ? 0 : value / max;
    return '<div class="bar-track">' +
      '<div class="seg seg-equity" style="flex:' + w + ' 0 0" data-tip="' + tip + '"></div>' +
      (1 - w > 0 ? '<div class="seg seg-rest" style="flex:' + (1 - w) + ' 0 0"></div>' : "") +
      "</div>";
  }

  /**
   * 含み損益バー。緑 = 元本のうち失われていない部分。
   *   含み益: [緑 取得原価][赤ハッチ 含み益]  → 棒の全長 = 評価額
   *   含み損: [緑 評価額  ][青ハッチ 含み損]  → 棒の全長 = 取得原価
   * どちらも棒の全長は max(取得原価, 評価額)。max を渡して行間で同じ縮尺にする。
   */
  function plBar(cost, value, max, label) {
    if (cost <= 0 && value <= 0) {
      return '<div class="bar-track"><div class="seg seg-rest" style="flex:1 0 0"></div></div>';
    }
    const gain = value - cost;
    const base = Math.min(cost, value);
    const segs = [{
      cls: "seg-equity", value: base,
      label: label + "<br>" + (gain >= 0 ? "取得原価" : "評価額（残っている元本）")
    }];
    if (gain > 0) segs.push({ cls: "seg-gain", value: gain,  label: label + "<br>含み益" });
    if (gain < 0) segs.push({ cls: "seg-loss", value: -gain, label: label + "<br>含み損" });
    return stackedBar(segs, max);
  }

  /* ---------- 円グラフ ----------
     構成比を一目で示すためだけに使う。金額は横の凡例で読む。
     SPEC で一度却下したが、分類を3つに固定したため
     「口座や借入先が増えると読めなくなる」という却下理由が当たらない。
     色は dataviz スキルの validate_palette.js で両モード・全ペアを検証済み。 */
  const PIE_ASSET = ["#bb6802", "#0089ca", "#319751"];   /* 不動産 / 有価証券 / 預貯金 */
  const PIE_DEBT  = ["#bb6802", "#8e6ac7", "#00989a"];   /* 不動産 / 多目的 / 自動車 */

  /**
   * slices = [{label, value}]。合計が0、または中身が1種類なら描かない
   * （1スライスの円は情報がなく、2スライスも棒や数字に劣るため）。
   * 戻り値は {svg, legend, drawn}。
   */
  function pieChart(slices, colors, opts) {
    const o = opts || {};
    const use = slices.filter(s => s.value > 0);
    const total = sum(use, s => s.value);
    if (use.length < 3 || total <= 0) return { drawn: false, rows: use, total: total };

    const size = o.size || 140, r = size / 2 - 2, cx = size / 2, cy = size / 2;
    const gap = 0.012;                       /* スライス間のすき間（ラジアン） */
    let angle = -Math.PI / 2;                /* 12時から時計回り */
    const paths = use.map((sl, i) => {
      const frac = sl.value / total;
      const a0 = angle + gap / 2, a1 = angle + frac * Math.PI * 2 - gap / 2;
      angle += frac * Math.PI * 2;
      const large = (a1 - a0) > Math.PI ? 1 : 0;
      const x0 = cx + r * Math.cos(a0), y0 = cy + r * Math.sin(a0);
      const x1 = cx + r * Math.cos(a1), y1 = cy + r * Math.sin(a1);
      const d = "M " + cx + " " + cy + " L " + x0.toFixed(2) + " " + y0.toFixed(2) +
                " A " + r + " " + r + " 0 " + large + " 1 " + x1.toFixed(2) + " " + y1.toFixed(2) + " Z";
      return '<path d="' + d + '" fill="' + colors[i % colors.length] + '"' +
             ' data-tip="' + sl.label + "<br><b>" + yen(sl.value) + "</b>（" +
             (frac * 100).toFixed(1) + '%）"></path>';
    }).join("");

    const svg = '<svg viewBox="0 0 ' + size + " " + size + '" width="' + size + '" height="' + size +
      '" role="img" aria-label="' + (o.title || "構成比") + '">' + paths + "</svg>";

    /* 凡例は必ず出す。色だけに意味を持たせないため、名前・金額・割合を並べる。 */
    const legend = '<table class="pie-legend">' + use.map((sl, i) =>
      '<tr><td><i class="swatch" style="background:' + colors[i % colors.length] + '"></i>' +
        sl.label + "</td>" +
      '<td class="n num">' + yen(sl.value) + "</td>" +
      '<td class="n num">' + (sl.value / total * 100).toFixed(1) + "%</td></tr>").join("") +
      '<tr class="total-row"><td>合計</td><td class="n num">' + yen(total) +
      '</td><td class="n num">100.0%</td></tr></table>';
    return { drawn: true, svg: svg, legend: legend, total: total, rows: use };
  }

  /* 資産の3分類。不動産の時価がそろっていれば不動産を立て、
     そうでなければ金融資産だけを3つに割る（常に3スライスに保つ）。 */
  function assetSlices() {
    if (RE.allValued && RE.marketValue > 0) {
      return { title: "総資産の内訳", slices: [
        { label: "不動産（時価）", value: RE.marketValue },
        { label: "有価証券", value: SEC_VALUE },
        { label: INS_VALUE > 0 ? "預貯金・保険" : "預貯金", value: DEPOSITS + INS_VALUE }
      ] };
    }
    return { title: "金融資産の内訳", note: "不動産は時価が未入力のため含みません", slices: [
      { label: "有価証券", value: SEC_VALUE },
      { label: "保険（解約返戻金）", value: INS_VALUE },
      { label: "預貯金", value: DEPOSITS }
    ] };
  }

  /* 負債の3分類。増えても3つのままにするため、車以外はすべて「多目的・その他」に寄せる。 */
  function debtSlices() {
    const car = sum(D.otherLoans.filter(l => l.type === "車"), l => l.balance);
    return { title: "総負債の内訳", slices: [
      { label: "不動産ローン", value: RE.balance },
      { label: "多目的・その他ローン", value: OTHER_DEBT - car },
      { label: "自動車ローン", value: car }
    ] };
  }

  /* ---------- サイドバー ---------- */
  function sidebar(current) {
    const pages = [
      ["index.html",      "① サマリー"],
      ["realestate.html", "② 不動産"],
      ["accounts.html",   "③ 資産・負債明細"],
      ["report.html",     "提出用（印刷・PDF）"]
    ];
    const nav = pages.map(([href, label]) =>
      '<a href="' + href + '"' + (href === current ? ' aria-current="page"' : "") + ">" + label + "</a>"
    ).join("");
    return '' +
      '<div class="owner">' +
        '<div class="owner-label">保有者</div>' +
        '<div class="owner-name">' + esc(D.owner.name) + "</div>" +
        '<div class="owner-asof num">' + esc(D.asOf) + " 時点</div>" +
      "</div>" +
      '<nav class="nav">' + nav + "</nav>" +
      '<div class="sidebar-foot">' +
        '<div class="data-source">' +
          '<div class="owner-label">表示中のデータ</div>' +
          '<div class="data-source-name">' + esc(SOURCE.label) + "</div>" +
          '<label class="btn-load">JSONファイルを読み込む' +
            '<input type="file" id="json-file" accept=".json,application/json"></label>' +
          (SOURCE.kind === "json"
            ? '<button type="button" class="btn-reset" id="reset-sample">サンプルに戻す</button>' : "") +
        "</div>" +
        '<button class="btn-future" disabled title="将来のアプリ化で実装します">' +
          "物件追加<small>準備中</small>" +
        "</button>" +
      "</div>";
  }

  /* ---------- 解説のたたみ込み ----------
     読み方の説明は毎月見る人には邪魔なのでたたむ。開閉はブラウザに覚えさせる。
     判断を誤らせないための短い囲み（.callout）はたたまない。
     印刷時は @media print で中身を必ず見せる（提出先は README を読めないため）。 */
  function foldNotes(label) {
    const key = "assetDashboard.notesOpen";
    let open = false;
    try { open = localStorage.getItem(key) === "1"; } catch (e) {}
    document.querySelectorAll("section > .sec-note").forEach(note => {
      const d = document.createElement("details");
      d.className = "note-fold";
      if (open) d.open = true;
      const sm = document.createElement("summary");
      sm.textContent = label || "この表の読み方";
      note.parentNode.insertBefore(d, note);
      d.appendChild(sm);
      d.appendChild(note);
      d.addEventListener("toggle", () => {
        if (printing) return;                 /* 印刷のための開閉は記憶しない */
        try { localStorage.setItem(key, d.open ? "1" : "0"); } catch (e) {}
        document.querySelectorAll(".note-fold").forEach(o => { o.open = d.open; });
      });
    });

    /* 印刷のあいだは必ず開く。閉じた <details> の中身は CSS では出せない
       （ブラウザが内部で隠すため）ので、開閉そのものを切り替える。
       提出先は README を読めないため、紙には説明が要る。 */
    let printing = false;
    const setOpen = on => {
      printing = true;
      document.querySelectorAll(".note-fold").forEach(d => {
        if (on) { d.dataset.wasOpen = d.open ? "1" : "0"; d.open = true; }
        else { d.open = d.dataset.wasOpen === "1"; }
      });
      printing = false;
    };
    addEventListener("beforeprint", () => setOpen(true));
    addEventListener("afterprint", () => setOpen(false));
    const mq = matchMedia("print");
    if (mq.addEventListener) mq.addEventListener("change", e => setOpen(e.matches));
    if (mq.matches) setOpen(true);
  }

  /* ---------- ツールチップ ---------- */
  function initTooltip() {
    const tip = document.createElement("div");
    tip.id = "tip";
    document.body.appendChild(tip);
    document.addEventListener("mouseover", e => {
      const t = e.target.closest("[data-tip]");
      if (!t) return;
      tip.innerHTML = t.dataset.tip;
      tip.classList.add("on");
    });
    document.addEventListener("mousemove", e => {
      if (!tip.classList.contains("on")) return;
      const pad = 14;
      let x = e.clientX + pad, y = e.clientY + pad;
      const r = tip.getBoundingClientRect();
      if (x + r.width > innerWidth - 8) x = e.clientX - r.width - pad;
      if (y + r.height > innerHeight - 8) y = e.clientY - r.height - pad;
      tip.style.left = x + "px";
      tip.style.top = y + "px";
    });
    document.addEventListener("mouseout", e => {
      if (e.target.closest("[data-tip]")) tip.classList.remove("on");
    });
  }

  window.APP = {
    D, SOURCE, validateData, parseJSONText, RE, SEC, SEC_VALUE, SEC_COST, SEC_PL, DEPOSITS,
    INS, INS_VALUE, INS_COST, INS_PL, FIN_PL,
    FINANCIAL_ASSETS, OTHER_DEBT, TOTAL_DEBT, CF_TOTAL,
    TOTAL_ASSETS, NET_WORTH, equity,
    yen, plain, pl, plHTML, plPctHTML, pct, sum, cashflow, fmtTerm,
    stackedBar, valueBar, plBar, sidebar, initTooltip, foldNotes,
    pieChart, assetSlices, debtSlices, PIE_ASSET, PIE_DEBT
  };
})();
