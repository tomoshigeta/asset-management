/* ============================================================
   資産データ — このファイル1枚だけを手で書き換えます（仕様 Q8）
   ------------------------------------------------------------
   不変条件（守らないと画面が壊れます）:
     - 期間を持つ金額（収支・返済）はすべて「月額」    … Q25
     - 固定資産税は「年額 ÷ 12」の平均月額を入れる
     - securities のみ元通貨で持ち、rates で円換算       … Q21/Q26
     - totalInvestment = ownFunds + loan.principal
     - loan.principal  = loan.balance + loan.cumulativeRepaid
   ------------------------------------------------------------
   .json ではなく .js にしてあるのは、ブラウザで index.html を
   ダブルクリックしただけで動かすためです（file:// では fetch が
   CORS で失敗するため）。ローカルサーバで配信する段階になったら
   中身の { ... } をそのまま data.json に移し、fetch に変えられます。
   ============================================================ */

window.ASSET_DATA = {
  owner: { name: "山田 太郎" },
  asOf: "2026-09-01",

  /* 為替レート（手書き。取得日時点で固定する） */
  rates: { JPY: 1, USD: 155.2 },

  /* ---------- 不動産 ---------- */
  realEstate: [
    {
      name: "物件A",
      address: "東京都港区白金台",
      areaSqm: 45.5, floor: "3 / 10 階", ageYears: 15,
      totalInvestment: 50000000,
      ownFunds: 10000000,
      /* 時価は任意。手で入れた推定値で、根拠と評価日も残せる（どちらも省略可） */
      marketValue: 58000000, valuedAt: "2026-08-01", valuationBasis: "不動産会社の査定書",
      loan: {
        lender: "A銀行", principal: 40000000,
        balance: 35500000, cumulativeRepaid: 4500000,
        rate: 1.8, termMonths: 420, monthlyRepayment: 128000
      },
      monthly: {
        rentIncome: 180000, managementFee: 12000,
        repairReserve: 8000, propertyTaxMonthly: 9000
      }
    },
    {
      name: "物件B",
      address: "東京都渋谷区恵比寿",
      areaSqm: 60.2, floor: "8 / 15 階", ageYears: 5,
      totalInvestment: 75000000,
      ownFunds: 15000000,
      marketValue: 71000000, valuedAt: "2026-08-01", valuationBasis: "取引事例から推定",
      loan: {
        lender: "B信用金庫", principal: 60000000,
        balance: 58000000, cumulativeRepaid: 2000000,
        rate: 0.8, termMonths: 360, monthlyRepayment: 190000
      },
      monthly: {
        rentIncome: 240000, managementFee: 15000,
        repairReserve: 12000, propertyTaxMonthly: 12000
      }
    },
    {
      name: "物件C",
      address: "神奈川県横浜市西区",
      areaSqm: 85.0, floor: "2 / 5 階", ageYears: 22,
      totalInvestment: 25000000,
      ownFunds: 5000000,
      marketValue: 27500000, valuedAt: "2026-07-15", valuationBasis: "路線価から換算",
      loan: {
        lender: "C銀行", principal: 20000000,
        balance: 11500000, cumulativeRepaid: 8500000,
        rate: 1.2, termMonths: 323, monthlyRepayment: 105000
      },
      monthly: {
        rentIncome: 95000, managementFee: 8000,
        repairReserve: 6000, propertyTaxMonthly: 5000
      }
    }
  ],

  /* ---------- 有価証券（資産クラス別・評価額ベース） ----------
     assetClass は 株 / 債券 / 投資信託 / その他 の4種類のみ。
     「その他」は暗号資産・金・銀などの現物資産をまとめる器です。
     国内・海外とも4行を常に表示するため、残高ゼロでも行は消しません。 */
  securities: [
    { region: "domestic", assetClass: "株",       currency: "JPY", cost: 2000000, marketValue: 5000000 },
    { region: "domestic", assetClass: "債券",     currency: "JPY", cost: 2700000, marketValue: 2500000 },
    { region: "domestic", assetClass: "投資信託", currency: "JPY", cost: 1300000, marketValue: 1500000 },
    { region: "domestic", assetClass: "その他",   currency: "JPY", cost:  800000, marketValue: 1200000,
      note: "金地金" },
    { region: "overseas", assetClass: "株",       currency: "USD", cost:   20000, marketValue:   32000 },
    { region: "overseas", assetClass: "債券",     currency: "USD", cost:   18000, marketValue:   16000 },
    { region: "overseas", assetClass: "投資信託", currency: "USD", cost:    8000, marketValue:    9500 },
    { region: "overseas", assetClass: "その他",   currency: "USD", cost:    5000, marketValue:   11000,
      note: "暗号資産" }
  ],

  /* ---------- 保険 ----------
     評価額は解約返戻金（今解約したら戻る額）、取得原価は払込保険料累計。
     deathBenefit（死亡保険金額）は参考表示のみで、資産の合計には足しません。
     valuedAt は評価日。年1回の通知の日付になるため asOf とズレることがあります。 */
  insurance: [
    { name: "変額終身保険", insurer: "○○生命", currency: "JPY",
      premiumPaid: 3000000, surrenderValue: 2600000,
      deathBenefit: 10000000, valuedAt: "2026-03-31" },
    { name: "外貨建終身保険", insurer: "△△生命", currency: "USD",
      premiumPaid: 30000, surrenderValue: 34000,
      deathBenefit: 100000, valuedAt: "2026-03-31", note: "通貨選択型" }
  ],

  /* ---------- 預貯金 ---------- */
  deposits: [
    { bank: "A銀行", currency: "JPY", amount:  2000000 },
    { bank: "B銀行", currency: "JPY", amount:  3000000 },
    { bank: "C銀行", currency: "JPY", amount: 13000000 }
  ],

  /* ---------- 不動産以外の借入（ローン単位で保持） ---------- */
  otherLoans: [
    { name: "自動車ローン", type: "車",     lender: "D銀行",
      principal: 3000000, balance: 1800000, rate: 2.5, monthlyRepayment: 45000 },
    { name: "多目的ローン", type: "多目的", lender: "E銀行",
      principal: 5000000, balance: 3200000, rate: 3.2, monthlyRepayment: 62000 }
  ]
};
