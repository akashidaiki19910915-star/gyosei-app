(function () {
  const DEBUG_MODE = false;
  const debugLog = (...args) => {
    if (DEBUG_MODE) console.log(...args);
  };
  const ESTIMATE_CALCULATION_MUTATION_COLUMNS = [
    "user_id", "client_id", "project_name", "work_type", "application_type", "corporate_type", "governor_type",
    "general_specific", "industry_count", "officer_count", "office_count", "document_level", "urgent",
    "keikan_level", "sengi_level", "zaisan_level", "keikan_reason", "sengi_reason", "zaisan_reason", "document_level_reason",
    "expense_amount", "discount_amount", "memo", "base_fee", "addon_fee", "taxable_subtotal", "tax", "total",
    "addon_breakdown", "reflected_estimate_id", "reflected_at",
  ];
  const pickObjectKeys = (source, keys) => keys.reduce((acc, key) => {
    if (Object.prototype.hasOwnProperty.call(source, key)) acc[key] = source[key];
    return acc;
  }, {});
  const buildSaveErrorMessage = (target, error) => {
    const detail = error?.message ? ` ${error.message}` : "";
    return `${target}の保存に失敗しました。${detail} DB未定義カラムが含まれている可能性があります。`;
  };
  const NOTICE = "この金額は入力内容に基づく概算です。正式な報酬額は必要資料・申請先・業務範囲確認後に確定します。";
  const PRINT_NOTICE_EXTRA = "上記報酬には、申請書類作成、要件確認、必要書類案内、申請先確認、提出準備に関する業務を含みます。";
  const CONSTRUCTION_BURDEN_NOTICE = "この区分は許可取得の可否を示すものではなく、要件確認・証明資料収集の作業負担を見積金額へ反映するためのものです。";
  const CONSTRUCTION_BURDEN_WORK_TYPES = new Set(["建設業許可", "業種追加"]);
  const CONSTRUCTION_BURDEN_FIELDS = ["keikan", "sengi", "zaisan"];
  const CONSTRUCTION_BURDEN_LEVEL_ORDER = { "低": 0, "中": 1, "高": 2 };
  const CONSTRUCTION_BURDEN_AUTO_REASON_PREFIX = "【ヒアリング自動判定】";
  const WORK_DIFFICULTY_AUTO_REASON_PREFIX = "【業務別ヒアリング自動判定】";
  const CONSTRUCTION_BURDEN_CONFIG = {
    keikan: {
      label: "経管要件の確認負担",
      reasonLabel: "経管要件の判定理由",
      reasonName: "keikanReason",
      addonAmounts: { "建設業許可": 30000 },
      criteria: {
        "低": "同一会社で5年以上の役員・個人事業主経験があり、決算書、登記簿、工事実績、常勤資料がそろっている。",
        "中": "複数会社・個人事業の経験合算、過去資料の取寄せ、一部資料の追加確認が必要だが、通常の5年経験ルートで確認できる。",
        "高": "執行役員への権限委譲、6年以上の補佐経験、役員と補佐体制を組み合わせる特殊ルート、事前相談、証明元の廃業・非協力等がある。",
      },
    },
    sengi: {
      label: "営業所技術者要件の確認負担",
      reasonLabel: "営業所技術者要件の判定理由",
      reasonName: "sengiReason",
      addonAmounts: { "建設業許可": 30000, "業種追加": 30000 },
      criteria: {
        "低": "申請業種に直接対応する国家資格があり、資格・常勤資料がそろっている。",
        "中": "指定学科卒業後3年・5年の実務経験、第一次検定合格後3年・5年の実務経験などで、経験資料の確認が必要。",
        "高": "資格なしで10年以上の実務経験、複数勤務先・複数業種の経験合算、期間重複、資料不足、特定建設業の指導監督的経験等がある。",
      },
    },
    zaisan: {
      label: "財産要件の確認負担",
      reasonLabel: "財産要件の判定理由",
      reasonName: "zaisanReason",
      addonAmounts: { "建設業許可": 20000 },
      criteria: {
        "低": "一般建設業で、直前決算の自己資本が500万円以上、または5年目更新で要件が明確。",
        "中": "一般建設業で500万円以上の預金残高証明を使用する、新設法人・個人、基準付近で資料確認が必要。",
        "高": "特定建設業、または欠損・流動比率・資本金・自己資本について詳細確認や基準抵触の可能性がある。",
      },
    },
  };
  const WORK_DIFFICULTY_CONFIG = {
    "建設業許可": {
      label: "申請資料の確認負担",
      addonAmounts: { "中": 20000, "高": 40000 },
      criteria: {
        "低": "必要年分の決算書、登記簿、工事関係資料、常勤資料などが整理され、申請内容との照合が容易。",
        "中": "証明書の取寄せ、一部年度・工事資料の追加収集、記載内容の照合や軽微な補足説明が必要。",
        "高": "複数年度・複数法人にまたがる資料不足、内容不一致、実績の再整理、行政庁への事前確認が必要。",
      },
    },
    "業種追加": {
      label: "追加業種資料の確認負担",
      addonAmounts: { "中": 15000, "高": 30000 },
      criteria: {
        "低": "既存許可資料と追加業種の資格・工事実績資料がそろい、対象業種との対応関係が明確。",
        "中": "追加業種の工事資料取寄せ、業種区分の照合、一部期間・請負内容の補足確認が必要。",
        "高": "複数業種・複数勤務先の実績整理、業種判定の事前相談、資料不足や期間重複の詳細確認が必要。",
      },
    },
    "決算変更届": {
      label: "決算資料の整理負担",
      addonAmounts: { "中": 10000, "高": 25000 },
      criteria: {
        "低": "対象年度の決算書、税務申告書、工事経歴書、直前三年の工事施工金額資料がそろっている。",
        "中": "工事台帳・請求書との照合、工事区分の整理、一部資料の追加収集や税理士確認が必要。",
        "高": "複数年度の未提出、工事資料の大幅不足、決算数値との不一致、工事実績の再構成が必要。",
      },
    },
    "各種変更届": {
      label: "変更事項の整理負担",
      addonAmounts: { "中": 10000, "高": 25000 },
      criteria: {
        "低": "単一の変更で、変更日・登記事項・裏付け資料が明確かつ提出期限内。",
        "中": "複数変更、過去日に遡る変更、役員・営業所等の関連資料の追加確認が必要。",
        "高": "未届変更が累積し、変更順序・効力発生日が不明確、他の許可事項との整合や事前相談が必要。",
      },
    },
    "宅建業免許": {
      label: "宅建業免許資料の確認負担",
      addonAmounts: { "中": 15000, "高": 30000 },
      criteria: {
        "低": "営業所、専任宅建士、役員、財務関係の必要資料がそろい、使用権限も明確。",
        "中": "複数営業所・複数役員、使用承諾、専任性や常勤性について追加資料の確認が必要。",
        "高": "営業所使用権限・専任性の資料不足、複雑な役員関係、複数行政庁との調整や事前相談が必要。",
      },
    },
    "株式会社設立": {
      label: "会社設計の確認負担",
      addonAmounts: { "中": 10000, "高": 25000 },
      criteria: {
        "低": "発起人・役員が少数で、現金出資、標準的な機関設計・定款内容で確認事項が明確。",
        "中": "発起人・役員が複数、事業目的の調整、任期・株式譲渡制限・資本構成の個別確認が必要。",
        "高": "現物出資、種類株式、複雑な機関設計、外部投資家、専門家間の調整を伴う。",
      },
    },
    "合同会社設立": {
      label: "会社設計の確認負担",
      addonAmounts: { "中": 10000, "高": 20000 },
      criteria: {
        "低": "社員が少数で、現金出資、標準的な定款・利益配分で確認事項が明確。",
        "中": "社員が複数で、業務執行、代表、利益配分、退社時の取扱いについて個別調整が必要。",
        "高": "法人社員、現物出資、外国関係者、複雑な利益配分・意思決定設計や専門家調整を伴う。",
      },
    },
    "創業融資": {
      label: "事業計画作成の確認負担",
      addonAmounts: { "中": 20000, "高": 40000 },
      criteria: {
        "低": "事業内容、必要資金、見積書、自己資金、売上・経費見込みの根拠が整理されている。",
        "中": "市場・競合、売上根拠、資金使途、返済計画について追加ヒアリングと資料補強が必要。",
        "高": "複数事業・複数調達、資金計画の大幅な組直し、既存借入や収支根拠の詳細整理が必要。",
      },
    },
    "車庫証明": {
      label: "図面・使用権限資料の確認負担",
      addonAmounts: { "中": 3000, "高": 8000 },
      criteria: {
        "低": "自己所有地または明確な契約駐車場で、所在図・配置図に必要な寸法情報がそろっている。",
        "中": "使用承諾証明書の取得、現地寸法確認、所在図・配置図の追加作成が必要。",
        "高": "保管場所の境界・使用権限・本拠との位置関係が不明確、複数台・複数区画や警察署確認が必要。",
      },
    },
    "会社設立": {
      label: "定款・機関設計の確認負担",
      addonAmounts: { "中": 10000, "高": 25000 },
      criteria: {
        "低": "会社形態、出資者、役員、事業目的、資本金が確定し、標準的な定款設計で対応できる。",
        "中": "複数関係者、事業目的・役員任期・持分や株式構成などの個別調整が必要。",
        "高": "現物出資、外国関係者、複雑な機関・資本設計、種類株式等の専門家調整を伴う。",
      },
    },
    "産業廃棄物収集運搬業許可": {
      label: "許可要件資料の確認負担",
      addonAmounts: { "中": 20000, "高": 40000 },
      criteria: {
        "低": "単一都道府県・少数車両で、講習修了証、車検証、駐車場、財務資料がそろっている。",
        "中": "複数車両・複数品目、賃貸車両・駐車場、追加の使用権限資料や財務説明が必要。",
        "高": "複数都道府県、特殊車両・品目、積替え保管等の個別確認、財務資料不足や行政庁との事前相談が必要。",
      },
    },
    "在留資格関連": {
      label: "疎明資料作成の確認負担",
      addonAmounts: { "中": 15000, "高": 30000 },
      criteria: {
        "低": "本人情報、契約・雇用内容、所属機関資料、在留状況が明確で必要資料がそろっている。",
        "中": "職務内容・学歴経歴・身分関係の追加説明、海外書類の取寄せや翻訳が必要。",
        "高": "複雑な在留歴・活動変更、説明の不一致、複数国書類、所属機関側資料の大幅補強が必要。",
      },
    },
    "古物商許可": {
      label: "古物商許可資料の確認負担",
      addonAmounts: { "中": 10000, "高": 25000 },
      criteria: {
        "低": "単一営業所・少数役員で、住民票、略歴書、誓約書、営業所使用権限がそろっている。",
        "中": "複数営業所・複数役員、URL利用、使用承諾などの追加資料確認が必要。",
        "高": "営業所使用権限や管理者関係が不明確、過去変更の未整理、外国関係書類や警察署への事前確認が必要。",
      },
    },
  };
  const BASE = {
    "建設業許可|法人|新規|知事": 150000,
    "建設業許可|個人|新規|知事": 120000,
    "建設業許可|法人|更新|知事": 75000,
    "建設業許可|個人|更新|知事": 65000,
    "業種追加": 75000,
    "決算変更届": 40000,
    "各種変更届": 30000,
    "宅建業免許|新規": 120000,
    "宅建業免許|更新": 80000,
    "株式会社設立": 80000,
    "合同会社設立": 60000,
    "創業融資|事業計画書作成": 80000,
    "創業融資|面談対策込み": 120000,
    "車庫証明": 10000,
    "会社設立|株式会社": 90000,
    "会社設立|合同会社": 70000,
    "産業廃棄物収集運搬業許可|新規": 140000,
    "在留資格関連|認定": 120000,
    "在留資格関連|変更": 100000,
    "在留資格関連|更新": 70000,
    "古物商許可|法人": 70000,
    "古物商許可|個人": 55000,
  };

  const OPTION_SETS = {
    applicationDefault: ["新規", "更新", "変更", "事業計画書作成", "面談対策込み"],
    applicationByWorkType: {
      "建設業許可": ["新規", "更新"],
      "業種追加": ["追加"],
      "決算変更届": ["決算変更届"],
      "各種変更届": ["変更届"],
      "宅建業免許": ["新規", "更新"],
      "株式会社設立": ["設立"],
      "合同会社設立": ["設立"],
      "創業融資": ["事業計画書作成", "面談対策込み"],
      "車庫証明": ["新規", "変更", "代替"],
      "会社設立": ["株式会社", "合同会社"],
      "産業廃棄物収集運搬業許可": ["新規"],
      "在留資格関連": ["認定", "変更", "更新"],
      "古物商許可": ["新規"],
    },
  };

  const FIELD_CONFIG = {
    common: ["clientId", "projectName", "workType", "applicationType", "expense", "discount", "memo"],
    "建設業許可": ["corporateType", "governorType", "generalSpecific", "industryCount", "officerCount", "officeCount", "documentLevel", "urgent", "keikan", "sengi", "zaisan", "visit", "agent"],
    "業種追加": ["corporateType", "governorType", "industryCount", "officeCount", "documentLevel", "urgent", "sengi", "visit", "agent"],
    "決算変更届": ["industryCount", "documentLevel", "urgent", "visit", "agent"],
    "各種変更届": ["corporateType", "governorType", "officerCount", "officeCount", "documentLevel", "urgent", "visit", "agent"],
    "宅建業免許": ["corporateType", "governorType", "officeCount", "officerCount", "qualifiedStaffCheck", "guaranteeAssociation", "documentLevel", "urgent", "visit", "agent"],
    "株式会社設立": ["officerCount", "documentLevel", "urgent", "visit", "agent"],
    "合同会社設立": ["officerCount", "documentLevel", "urgent", "visit", "agent"],
    "創業融資": ["documentLevel", "urgent", "keikan", "sengi", "zaisan", "visit", "agent"],
    "車庫証明": ["industryCount", "selfCertification", "usageConsent", "baseLocationCheck", "documentLevel", "urgent", "visit", "agent"],
    "会社設立": ["officerCount", "documentLevel", "urgent", "visit", "agent"],
    "産業廃棄物収集運搬業許可": ["corporateType", "industryCount", "officeCount", "courseCompletion", "documentLevel", "urgent", "visit", "agent"],
    "在留資格関連": ["industryCount", "renewalDeadline", "documentLevel", "urgent", "visit", "agent"],
    "古物商許可": ["corporateType", "officeCount", "industryCount", "officerCount", "urlNotification", "documentLevel", "urgent", "visit", "agent"],
  };

  const FIELD_LABELS = {
    default: {
      corporateType: "法人/個人", governorType: "知事/大臣", generalSpecific: "一般/特定", industryCount: "業種数", officerCount: "役員数", officeCount: "営業所数", documentLevel: "書類不足レベル", urgent: "急ぎ対応", keikan: "経管確認難易度", sengi: "専技確認難易度", zaisan: "財産要件確認難易度", visit: "訪問対応", agent: "代理取得", urlNotification: "URL届出", selfCertification: "自認書", usageConsent: "使用承諾証明書", baseLocationCheck: "保管場所の本拠確認", qualifiedStaffCheck: "専任者/資格者確認", guaranteeAssociation: "保証協会/営業保証金", courseCompletion: "講習修了証", renewalDeadline: "更新期限",
    },
    "建設業許可": { keikan: "経管要件の確認負担", sengi: "営業所技術者要件の確認負担", zaisan: "財産要件の確認負担" },
    "業種追加": { industryCount: "追加業種数", sengi: "営業所技術者要件の確認負担" },
    "決算変更届": { industryCount: "対象年数", documentLevel: "資料整理難易度", agent: "税理士連携" },
    "各種変更届": { officerCount: "変更対象役員数", officeCount: "対象営業所数", documentLevel: "資料整理難易度" },
    "宅建業免許": { officerCount: "専任宅建士数", officeCount: "営業所数", qualifiedStaffCheck: "専任者/資格者確認", guaranteeAssociation: "保証協会/営業保証金", documentLevel: "資料整理難易度", agent: "保証協会・供託関連確認" },
    "株式会社設立": { officerCount: "発起人・役員数", documentLevel: "設計・確認事項の多さ", visit: "許認可同時相談", agent: "定款・登記連携調整" },
    "合同会社設立": { officerCount: "社員数", documentLevel: "設計・確認事項の多さ", visit: "許認可同時相談", agent: "定款・登記連携調整" },
    "創業融資": { documentLevel: "事業計画作成難易度", keikan: "資金繰り表作成", sengi: "面談対策", zaisan: "補助金・許認可併用確認", visit: "面談同席", agent: "追加資料作成" },
    "車庫証明": { industryCount: "台数", selfCertification: "自認書", usageConsent: "使用承諾証明書", baseLocationCheck: "保管場所の本拠確認", documentLevel: "図面・資料作成難易度", visit: "現地調査", agent: "使用承諾等の取得支援" },
    "会社設立": { officerCount: "関係者数", documentLevel: "定款・機関設計難易度", visit: "面談対応", agent: "公証人・司法書士連携調整" },
    "産業廃棄物収集運搬業許可": { industryCount: "運搬車両数", officeCount: "収集運搬先都道府県数", courseCompletion: "講習修了証", documentLevel: "許可要件整理難易度", visit: "現地確認", agent: "講習会・証明書取得支援" },
    "在留資格関連": { industryCount: "対象人数", renewalDeadline: "更新期限", documentLevel: "疎明資料作成難易度", visit: "本人面談対応", agent: "受入機関調整・追加資料対応" },
    "古物商許可": { officeCount: "営業所数", industryCount: "管理者数", officerCount: "役員数", urlNotification: "URL届出", documentLevel: "書類不足レベル", visit: "訪問対応", agent: "代理取得" },
  };

  function n(v) { const x = Number(v || 0); return Number.isFinite(x) ? x : 0; }
  function yen(v) { return `${Math.floor(v).toLocaleString("ja-JP")} 円`; }
  function h(s) { return String(s ?? "").replace(/[&<>\"']/g, ch => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#39;" }[ch])); }
  function fmtDate(v) { if (!v) return "-"; const d = new Date(v); return Number.isNaN(d.getTime()) ? String(v) : d.toLocaleDateString("ja-JP"); }
  function setOptions(select, options, currentValue) { select.innerHTML = (options || []).map(value => `<option value="${h(value)}">${h(value)}</option>`).join(""); if (currentValue && Array.from(select.options).some(o => o.value === currentValue)) select.value = currentValue; }
  function parseAddonBreakdown(v) { try { const parsed = typeof v === "string" ? JSON.parse(v || "[]") : v; return Array.isArray(parsed) ? parsed : []; } catch { return []; } }
  function getClientName(app, calc) { const clients = app?.getClients?.() || []; const client = clients.find((c) => String(c.id) === String(calc.client_id || "")); return client?.name || client?.companyName || client?.contactName || calc.client_name || calc.customer_name || "-"; }
  function getConstructionBurdenAddonAmount(workType, fieldName, level) {
    if (level !== "高") return 0;
    return n(CONSTRUCTION_BURDEN_CONFIG[fieldName]?.addonAmounts?.[workType]);
  }
  function getWorkDifficultyAddonAmount(workType, level) {
    return n(WORK_DIFFICULTY_CONFIG[workType]?.addonAmounts?.[level]);
  }
  function renderConstructionBurdenField(fieldName) {
    const config = CONSTRUCTION_BURDEN_CONFIG[fieldName];
    const criteria = Object.entries(config.criteria).map(([level, text]) => `<p><strong>${h(level)}：</strong>${h(text)}</p>`).join("");
    const hearingOptions = Object.entries(config.criteria).map(([level, text]) => `<label class="estimate-burden-hearing-option"><input type="checkbox" value="${h(level)}" data-burden-hearing-option="${h(fieldName)}"><span><strong>${h(level)}：</strong>${h(text)}</span></label>`).join("");
    return `<div class="estimate-burden-field" data-field="${h(fieldName)}"><label><span data-field-label>${h(config.label)}</span><select name="${h(fieldName)}"><option>低</option><option>高</option></select></label><p class="estimate-burden-addition" data-burden-addition="${h(fieldName)}" hidden></p><details class="estimate-burden-hearing" data-burden-hearing="${h(fieldName)}" hidden><summary>ヒアリング回答から自動判定</summary><div class="estimate-burden-hearing-body"><p class="meta">顧客から確認できた内容をすべて選択してください。選択した回答のうち、最も確認負担が高い区分を暫定判定します。</p><div class="estimate-burden-hearing-options">${hearingOptions}</div><div class="estimate-burden-hearing-result-row"><p class="estimate-burden-hearing-result" data-burden-hearing-result="${h(fieldName)}">自動判定：未判定</p><button type="button" class="secondary-btn" data-burden-hearing-clear="${h(fieldName)}">回答をクリア</button></div><p class="meta">自動判定は見積作業負担の目安です。許可可否や法的判断を示すものではなく、判定後も区分を手動で変更できます。</p></div></details><details class="estimate-burden-criteria" data-burden-criteria="${h(fieldName)}" hidden><summary>判定基準を見る</summary><div>${criteria}</div></details><label class="estimate-burden-reason" data-burden-reason="${h(fieldName)}" hidden><span>${h(config.reasonLabel)}</span><textarea name="${h(config.reasonName)}" rows="3"></textarea></label></div>`;
  }
  function renderWorkDifficultyField() {
    return `<div class="estimate-burden-field estimate-work-difficulty-field" data-field="documentLevel"><label><span data-field-label>書類不足レベル</span><select name="documentLevel"><option>低</option><option>中</option><option>高</option></select></label><p class="estimate-burden-addition" data-work-difficulty-addition></p><details class="estimate-burden-hearing" data-work-difficulty-hearing><summary>ヒアリング回答から自動判定</summary><div class="estimate-burden-hearing-body"><p class="meta">顧客から確認できた内容をすべて選択してください。選択した回答のうち、最も作業負担が高い区分を暫定判定します。</p><div class="estimate-burden-hearing-options" data-work-difficulty-options></div><div class="estimate-burden-hearing-result-row"><p class="estimate-burden-hearing-result" data-work-difficulty-result>自動判定：未判定</p><button type="button" class="secondary-btn" data-work-difficulty-clear>回答をクリア</button></div><p class="meta">自動判定は見積作業負担の目安です。許可可否、融資可否、在留可否その他の法的・行政上の結論を示すものではなく、判定後も区分を手動で変更できます。</p></div></details><label class="estimate-burden-reason" data-work-difficulty-reason hidden><span data-work-difficulty-reason-label>難易度の判定理由</span><textarea name="documentLevelReason" rows="3"></textarea></label></div>`;
  }
  async function waitForGyoseiApp(maxMs = 10000) { const start = Date.now(); while (Date.now() - start < maxMs) { if (window.GyoseiApp) return window.GyoseiApp; await new Promise(resolve => setTimeout(resolve, 100)); } return window.GyoseiApp || null; }
  async function waitForCurrentUser(app, maxMs = 10000) { const start = Date.now(); while (Date.now() - start < maxMs) { const u = app?.getCurrentUser?.(); if (u?.id) return u; await new Promise(resolve => setTimeout(resolve, 100)); } return app?.getCurrentUser?.() || null; }

  function calculateBase(f) { const wt = f.workType.value; if (wt === "建設業許可") return BASE[`建設業許可|${f.corporateType.value}|${f.applicationType.value}|${f.governorType.value}`] || 0; if (wt === "宅建業免許") return BASE[`宅建業免許|${f.applicationType.value}`] || 0; if (wt === "創業融資") return BASE[`創業融資|${f.applicationType.value}`] || BASE["創業融資|事業計画書作成"] || 0; if (wt === "会社設立") return BASE[`会社設立|${f.applicationType.value}`] || BASE["会社設立|株式会社"] || 0; if (wt === "産業廃棄物収集運搬業許可") return BASE[`産業廃棄物収集運搬業許可|${f.applicationType.value}`] || BASE["産業廃棄物収集運搬業許可|新規"] || 0; if (wt === "在留資格関連") return BASE[`在留資格関連|${f.applicationType.value}`] || BASE["在留資格関連|認定"] || 0; if (wt === "古物商許可") return BASE[`古物商許可|${f.corporateType.value}`] || BASE["古物商許可|法人"] || 0; return BASE[wt] || 0; }
  function calculateAddons(f) {
    const wt = f.workType.value; const addons = []; const add = (name, amount) => { if (amount > 0) addons.push({ name, amount }); }; const doc = f.documentLevel.value; const docMidHigh = (mid, high) => doc === "中" ? mid : (doc === "高" ? high : 0);
    switch (wt) {
      case "建設業許可": add("業種加算", Math.max(0, n(f.industryCount.value) - 1) * 10000); add("役員加算", Math.max(0, n(f.officerCount.value) - 2) * 5000); add("営業所加算", Math.max(0, n(f.officeCount.value) - 1) * 30000); add("急ぎ対応", n(f.urgent.value) ? 30000 : 0); add("書類不足", docMidHigh(20000, 40000)); add("経管要件の確認負担", getConstructionBurdenAddonAmount(wt, "keikan", f.keikan.value)); add("営業所技術者要件の確認負担", getConstructionBurdenAddonAmount(wt, "sengi", f.sengi.value)); add("財産要件の確認負担", getConstructionBurdenAddonAmount(wt, "zaisan", f.zaisan.value)); add("訪問対応", n(f.visit.value) ? 10000 : 0); add("代理取得", n(f.agent.value) ? 10000 : 0); break;
      case "業種追加": add("追加業種加算", Math.max(0, n(f.industryCount.value) - 1) * 30000); add("営業所加算", Math.max(0, n(f.officeCount.value) - 1) * 20000); add("急ぎ対応", n(f.urgent.value) ? 25000 : 0); add("書類不足", docMidHigh(15000, 30000)); add("営業所技術者要件の確認負担", getConstructionBurdenAddonAmount(wt, "sengi", f.sengi.value)); add("訪問対応", n(f.visit.value) ? 10000 : 0); add("代理取得", n(f.agent.value) ? 10000 : 0); break;
      case "決算変更届": add("対象年数加算", Math.max(0, n(f.industryCount.value) - 1) * 30000); add("資料整理", docMidHigh(10000, 25000)); add("急ぎ対応", n(f.urgent.value) ? 15000 : 0); add("訪問対応", n(f.visit.value) ? 10000 : 0); add("税理士連携", n(f.agent.value) ? 10000 : 0); break;
      case "各種変更届": add("変更役員加算", Math.max(0, n(f.officerCount.value) - 1) * 5000); add("営業所加算", Math.max(0, n(f.officeCount.value) - 1) * 15000); add("資料整理", docMidHigh(10000, 25000)); add("急ぎ対応", n(f.urgent.value) ? 15000 : 0); add("訪問対応", n(f.visit.value) ? 10000 : 0); add("代理取得", n(f.agent.value) ? 10000 : 0); break;
      case "宅建業免許": add("営業所加算", Math.max(0, n(f.officeCount.value) - 1) * 40000); add("専任宅建士確認", Math.max(0, n(f.officerCount.value) - 1) * 10000); add("資料整理", docMidHigh(15000, 30000)); add("急ぎ対応", n(f.urgent.value) ? 30000 : 0); add("訪問対応", n(f.visit.value) ? 10000 : 0); add("保証協会・供託関連確認", n(f.agent.value) ? 20000 : 0); break;
      case "株式会社設立": add("発起人・役員加算", Math.max(0, n(f.officerCount.value) - 2) * 5000); add("設計・確認事項加算", docMidHigh(10000, 25000)); add("急ぎ対応", n(f.urgent.value) ? 20000 : 0); add("許認可同時相談", n(f.visit.value) ? 20000 : 0); add("定款・登記連携調整", n(f.agent.value) ? 15000 : 0); break;
      case "合同会社設立": add("社員数加算", Math.max(0, n(f.officerCount.value) - 1) * 5000); add("設計・確認事項加算", docMidHigh(10000, 20000)); add("急ぎ対応", n(f.urgent.value) ? 15000 : 0); add("許認可同時相談", n(f.visit.value) ? 20000 : 0); add("定款・登記連携調整", n(f.agent.value) ? 10000 : 0); break;
      case "創業融資": add("事業計画難易度", docMidHigh(20000, 40000)); add("急ぎ対応", n(f.urgent.value) ? 20000 : 0); add("資金繰り表作成", f.keikan.value === "高" ? 20000 : 0); add("面談対策強化", f.sengi.value === "高" ? 20000 : 0); add("補助金・許認可併用確認", f.zaisan.value === "高" ? 20000 : 0); add("面談同席", n(f.visit.value) ? 30000 : 0); add("追加資料作成", n(f.agent.value) ? 15000 : 0); break;
      case "車庫証明": add("台数加算", Math.max(0, n(f.industryCount.value) - 1) * 5000); add("図面・資料作成", docMidHigh(3000, 8000)); add("急ぎ対応", n(f.urgent.value) ? 5000 : 0); add("現地調査", n(f.visit.value) ? 5000 : 0); add("使用承諾等の取得支援", n(f.agent.value) ? 5000 : 0); break;
      case "会社設立": add("関係者加算", Math.max(0, n(f.officerCount.value) - 2) * 5000); add("定款・機関設計", docMidHigh(10000, 25000)); add("急ぎ対応", n(f.urgent.value) ? 20000 : 0); add("面談対応", n(f.visit.value) ? 10000 : 0); add("公証人・司法書士連携調整", n(f.agent.value) ? 15000 : 0); break;
      case "産業廃棄物収集運搬業許可": add("車両台数加算", Math.max(0, n(f.industryCount.value) - 1) * 7000); add("都道府県加算", Math.max(0, n(f.officeCount.value) - 1) * 12000); add("要件整理", docMidHigh(20000, 40000)); add("急ぎ対応", n(f.urgent.value) ? 30000 : 0); add("現地確認", n(f.visit.value) ? 15000 : 0); add("講習会・証明書取得支援", n(f.agent.value) ? 15000 : 0); break;
      case "在留資格関連": add("対象人数加算", Math.max(0, n(f.industryCount.value) - 1) * 15000); add("疎明資料作成", docMidHigh(15000, 30000)); add("急ぎ対応", n(f.urgent.value) ? 20000 : 0); add("本人面談対応", n(f.visit.value) ? 10000 : 0); add("受入機関調整・追加資料対応", n(f.agent.value) ? 15000 : 0); break;
      case "古物商許可": add("営業所加算", Math.max(0, n(f.officeCount.value) - 1) * 10000); add("役員加算", Math.max(0, n(f.officerCount.value) - 2) * 5000); add("URL届出", n(f.urlNotification.value) ? 10000 : 0); add("使用権限整理", n(f.visit.value) ? 10000 : 0); add("書類不足", docMidHigh(10000, 25000)); add("急ぎ対応", n(f.urgent.value) ? 15000 : 0); add("訪問対応", n(f.visit.value) ? 10000 : 0); add("代理取得", n(f.agent.value) ? 10000 : 0); break;
      default: add("急ぎ対応", n(f.urgent.value) ? 10000 : 0); add("資料整理", docMidHigh(10000, 20000)); add("訪問対応", n(f.visit.value) ? 10000 : 0); add("代理取得", n(f.agent.value) ? 10000 : 0);
    }
    return addons;
  }

  function openPrintWindow(app, calc) { const w = window.open("", "_blank", "width=960,height=1200"); if (!w) { alert("印刷用ウィンドウを開けませんでした。ポップアップブロックをご確認ください。"); return; } const createdAt = fmtDate(calc.created_at || new Date()); const gyoseiReward = n(calc.base_fee) + n(calc.addon_fee) - n(calc.discount_amount); const html = `<!doctype html><html lang="ja"><head><meta charset="utf-8"><title>概算見積書</title><style>@page{size:A4 portrait;margin:12mm}body{margin:0;padding:24px;font-family:"Yu Gothic","Hiragino Kaku Gothic ProN",Meiryo,sans-serif;color:#1f2937;background:#f8fafc}.sheet{max-width:840px;margin:0 auto;background:#fff;border:1px solid #e5e7eb;border-radius:10px;padding:24px}h1{margin:0 0 12px;font-size:28px;letter-spacing:.08em}.print-toolbar{display:flex;justify-content:flex-end;margin-bottom:12px}.print-btn{border:none;background:#2563eb;color:#fff;padding:10px 16px;border-radius:6px;font-size:14px;cursor:pointer}.meta{margin-bottom:16px;color:#4b5563}table{width:100%;border-collapse:collapse;margin:10px 0 18px}th,td{border:1px solid #d1d5db;padding:8px 10px;vertical-align:top;text-align:left}th{width:28%;background:#f3f4f6;font-weight:700}.money{text-align:right;font-variant-numeric:tabular-nums}.notice{margin-top:18px;padding:12px;border:1px solid #fbbf24;background:#fffbeb;border-radius:8px}@media print{body{background:#fff;padding:0}.sheet{border:none;border-radius:0;padding:0;max-width:none}.print-toolbar{display:none}}</style></head><body><main class="sheet"><div class="print-toolbar"><button type="button" class="print-btn" id="print-trigger">印刷する</button></div><h1>概算見積書</h1><p class="meta">作成日: ${h(createdAt)}</p><table><tbody><tr><th>顧客名</th><td>${h(getClientName(app, calc))}</td></tr><tr><th>案件名</th><td>${h(calc.project_name || "-")}</td></tr><tr><th>業務種別</th><td>${h(calc.work_type || "-")}</td></tr><tr><th>申請区分</th><td>${h(calc.application_type || "-")}</td></tr><tr><th>行政書士報酬</th><td class="money">${h(yen(gyoseiReward))}</td></tr><tr><th>実費・法定費用</th><td class="money">${h(yen(calc.expense_amount || 0))}</td></tr><tr><th>消費税</th><td class="money">${h(yen(calc.tax || 0))}</td></tr><tr><th>合計金額</th><td class="money"><strong>${h(yen(calc.total || 0))}</strong></td></tr><tr><th>注意書き</th><td>${h(`${NOTICE} ${PRINT_NOTICE_EXTRA}`)}</td></tr></tbody></table></main><script>document.getElementById('print-trigger').addEventListener('click',function(){window.print();});</script></body></html>`; w.document.open(); w.document.write(html); w.document.close(); w.focus(); }

  async function init() {
    const root = document.getElementById("estimate-calculator-root"); if (!root) return;
    root.innerHTML = `<form id="estimate-calc-form" class="form"><div class="grid cols-2"><label data-field="clientId">顧客<select name="clientId"></select></label><label data-field="projectName">案件名<input name="projectName" /></label><label data-field="workType">業務種別<select name="workType"><option>建設業許可</option><option>業種追加</option><option>決算変更届</option><option>各種変更届</option><option>宅建業免許</option><option>株式会社設立</option><option>合同会社設立</option><option>創業融資</option><option>車庫証明</option><option>会社設立</option><option>産業廃棄物収集運搬業許可</option><option>在留資格関連</option><option>古物商許可</option></select></label><label data-field="applicationType">申請区分<select name="applicationType"></select></label><label data-field="corporateType">法人/個人<select name="corporateType"><option>法人</option><option>個人</option></select></label><label data-field="governorType">知事/大臣<select name="governorType"><option>知事</option><option>大臣</option></select></label><label data-field="generalSpecific">一般/特定<select name="generalSpecific"><option>一般</option><option>特定</option></select></label><label data-field="industryCount">業種数<input name="industryCount" type="number" value="1" min="1" /></label><label data-field="officerCount">役員数<input name="officerCount" type="number" value="2" min="0" /></label><label data-field="officeCount">営業所数<input name="officeCount" type="number" value="1" min="1" /></label><label data-field="documentLevel">書類不足レベル<select name="documentLevel"><option>低</option><option>中</option><option>高</option></select></label><label data-field="urgent">急ぎ対応<select name="urgent"><option value="0">なし</option><option value="1">あり</option></select></label><label data-field="keikan">経管確認難易度<select name="keikan"><option>低</option><option>高</option></select></label><label data-field="sengi">専技確認難易度<select name="sengi"><option>低</option><option>高</option></select></label><label data-field="zaisan">財産要件確認難易度<select name="zaisan"><option>低</option><option>高</option></select></label><label data-field="visit">訪問対応<select name="visit"><option value="0">なし</option><option value="1">あり</option></select></label><label data-field="agent">代理取得<select name="agent"><option value="0">なし</option><option value="1">あり</option></select></label><label data-field="urlNotification">URL届出<select name="urlNotification"><option value="0">なし</option><option value="1">あり</option></select></label><label data-field="selfCertification">自認書<select name="selfCertification"><option value="0">不要</option><option value="1">要確認</option></select></label><label data-field="usageConsent">使用承諾証明書<select name="usageConsent"><option value="0">不要</option><option value="1">要確認</option></select></label><label data-field="baseLocationCheck">保管場所の本拠確認<select name="baseLocationCheck"><option value="0">不要</option><option value="1">要確認</option></select></label><label data-field="qualifiedStaffCheck">専任者/資格者確認<select name="qualifiedStaffCheck"><option value="0">不要</option><option value="1">要確認</option></select></label><label data-field="guaranteeAssociation">保証協会/営業保証金<select name="guaranteeAssociation"><option value="0">不要</option><option value="1">要確認</option></select></label><label data-field="courseCompletion">講習修了証<select name="courseCompletion"><option value="0">未確認</option><option value="1">確認済</option></select></label><label data-field="renewalDeadline">更新期限<select name="renewalDeadline"><option value="0">通常</option><option value="1">期限迫る</option></select></label><label data-field="expense">実費<input name="expense" type="number" value="0" min="0" /></label><label data-field="discount">値引き<input name="discount" type="number" value="0" min="0" /></label></div><label data-field="memo">メモ<textarea name="memo"></textarea></label><div class="row-actions"><button type="button" id="calc-run" class="secondary-btn">再計算</button><button type="button" id="calc-save">保存</button><button type="button" id="calc-apply">見積へ反映</button></div><div id="calc-result" class="panel"></div></form><section class="panel" id="calc-saved-list-wrap"><h3>保存済み概算見積</h3><div id="calc-saved-list"></div></section>`;
    root.querySelector('[data-field="documentLevel"]').outerHTML = renderWorkDifficultyField();
    root.querySelector('[data-field="keikan"]').insertAdjacentHTML("beforebegin", `<p id="construction-burden-notice" class="estimate-burden-notice" hidden>${h(CONSTRUCTION_BURDEN_NOTICE)}</p>`);
    CONSTRUCTION_BURDEN_FIELDS.forEach((fieldName) => {
      root.querySelector(`[data-field="${fieldName}"]`).outerHTML = renderConstructionBurdenField(fieldName);
    });
    const app = await waitForGyoseiApp(); if (!app) return;
    const form = root.querySelector('#estimate-calc-form'); const cs = form.elements.clientId; const savedList = root.querySelector('#calc-saved-list');
    (app?.getClients?.() || []).forEach(c => { const o = document.createElement('option'); o.value = c.id; o.textContent = c.name || c.companyName || c.contactName || '未設定'; cs.appendChild(o); });
    const visibleFieldsFor = (workType) => new Set([...(FIELD_CONFIG.common || []), ...(FIELD_CONFIG[workType] || [])]);
    const labelFor = (workType, fieldName) => FIELD_LABELS[workType]?.[fieldName] || FIELD_LABELS.default[fieldName] || null;
    const workDifficultyStateByType = new Map();
    const getWorkDifficultyInputs = () => Array.from(root.querySelectorAll("[data-work-difficulty-option]"));
    const removeWorkDifficultyAutoReasonLine = () => {
      const reason = form.elements.documentLevelReason;
      reason.value = String(reason.value || "")
        .split(/\r?\n/)
        .filter((line) => !line.startsWith(WORK_DIFFICULTY_AUTO_REASON_PREFIX))
        .join("\n")
        .trim();
    };
    const syncWorkDifficultyManualOverride = () => {
      const result = root.querySelector("[data-work-difficulty-result]");
      const autoLevel = result?.dataset.autoLevel;
      const baseText = result?.dataset.baseText;
      if (!result || !autoLevel || !baseText) return;
      const currentLevel = form.elements.documentLevel.value || "低";
      result.textContent = currentLevel === autoLevel
        ? baseText
        : `${baseText}／現在の選択：${currentLevel}（手動変更）`;
    };
    const updateWorkDifficultySelectionUi = (workType) => {
      const config = WORK_DIFFICULTY_CONFIG[workType];
      if (!config) return;
      const level = form.elements.documentLevel.value || "低";
      const amount = getWorkDifficultyAddonAmount(workType, level);
      root.querySelector("[data-work-difficulty-addition]").textContent = `現在の選択：${level}　加算額${Math.floor(amount).toLocaleString("ja-JP")}円`;
      root.querySelector("[data-work-difficulty-reason]").hidden = level === "低";
      root.querySelector("[data-work-difficulty-reason-label]").textContent = `${config.label}の判定理由`;
      syncWorkDifficultyManualOverride();
    };
    const applyWorkDifficultyDecision = ({ applyLevel = true, updateReason = true } = {}) => {
      const workType = form.elements.workType.value;
      const config = WORK_DIFFICULTY_CONFIG[workType];
      if (!config) return null;
      const selected = getWorkDifficultyInputs()
        .filter((input) => input.checked)
        .map((input) => ({ level: input.value, text: config.criteria[input.value] }))
        .filter((entry) => entry.text);
      const result = root.querySelector("[data-work-difficulty-result]");
      if (!selected.length) {
        delete result.dataset.autoLevel;
        delete result.dataset.baseText;
        result.textContent = "自動判定：未判定（現在の区分は変更していません）";
        if (updateReason) removeWorkDifficultyAutoReasonLine();
        updateWorkDifficultySelectionUi(workType);
        return null;
      }
      const level = selected.reduce((highest, entry) => (
        CONSTRUCTION_BURDEN_LEVEL_ORDER[entry.level] > CONSTRUCTION_BURDEN_LEVEL_ORDER[highest]
          ? entry.level
          : highest
      ), "低");
      if (applyLevel) form.elements.documentLevel.value = level;
      if (updateReason) {
        const reason = form.elements.documentLevelReason;
        const manualReason = String(reason.value || "")
          .split(/\r?\n/)
          .filter((line) => !line.startsWith(WORK_DIFFICULTY_AUTO_REASON_PREFIX))
          .join("\n")
          .trim();
        const autoReason = `${WORK_DIFFICULTY_AUTO_REASON_PREFIX}区分：${level}／回答：${selected.map((entry) => `${entry.level}：${entry.text}`).join("／")}`;
        reason.value = [autoReason, manualReason].filter(Boolean).join("\n");
      }
      const highestCount = selected.filter((entry) => entry.level === level).length;
      const baseText = `自動判定：${level}（${level}の該当内容${highestCount}件／選択${selected.length}件）`;
      result.dataset.autoLevel = level;
      result.dataset.baseText = baseText;
      updateWorkDifficultySelectionUi(workType);
      return level;
    };
    const restoreWorkDifficultyAnswers = () => {
      const workType = form.elements.workType.value;
      const config = WORK_DIFFICULTY_CONFIG[workType];
      if (!config) return;
      const reason = String(form.elements.documentLevelReason.value || "");
      const autoReasonLine = reason.split(/\r?\n/).find((line) => line.startsWith(WORK_DIFFICULTY_AUTO_REASON_PREFIX)) || "";
      getWorkDifficultyInputs().forEach((input) => {
        input.checked = !!autoReasonLine && autoReasonLine.includes(`${input.value}：${config.criteria[input.value]}`);
      });
      applyWorkDifficultyDecision({ applyLevel: false, updateReason: false });
    };
    const renderWorkDifficultyOptions = (workType, selectedLevels = []) => {
      const config = WORK_DIFFICULTY_CONFIG[workType];
      const options = root.querySelector("[data-work-difficulty-options]");
      options.innerHTML = Object.entries(config.criteria).map(([level, text]) => `<label class="estimate-burden-hearing-option"><input type="checkbox" value="${h(level)}" data-work-difficulty-option ${selectedLevels.includes(level) ? "checked" : ""}><span><strong>${h(level)}：</strong>${h(text)}</span></label>`).join("");
      getWorkDifficultyInputs().forEach((input) => input.addEventListener("change", () => {
        applyWorkDifficultyDecision();
        run();
      }));
    };
    const activateWorkDifficultyUi = (workType) => {
      const config = WORK_DIFFICULTY_CONFIG[workType];
      const panel = root.querySelector(".estimate-work-difficulty-field");
      if (!config || !panel) return;
      const previousWorkType = panel.dataset.workType || "";
      if (previousWorkType !== workType) {
        if (previousWorkType && WORK_DIFFICULTY_CONFIG[previousWorkType]) {
          workDifficultyStateByType.set(previousWorkType, {
            level: form.elements.documentLevel.value || "低",
            reason: form.elements.documentLevelReason.value || "",
            selectedLevels: getWorkDifficultyInputs().filter((input) => input.checked).map((input) => input.value),
          });
        }
        const savedState = workDifficultyStateByType.get(workType);
        panel.dataset.workType = workType;
        form.elements.documentLevel.value = savedState?.level || "低";
        form.elements.documentLevelReason.value = savedState?.reason || "";
        renderWorkDifficultyOptions(workType, savedState?.selectedLevels || []);
        if (savedState?.selectedLevels?.length) applyWorkDifficultyDecision({ applyLevel: false, updateReason: false });
        else applyWorkDifficultyDecision({ applyLevel: false, updateReason: false });
      }
      panel.querySelector("[data-field-label]").textContent = config.label;
      updateWorkDifficultySelectionUi(workType);
    };
    const updateConstructionBurdenUi = (workType, visibleFields) => {
      const isConstructionBurdenWork = CONSTRUCTION_BURDEN_WORK_TYPES.has(workType);
      root.querySelector("#construction-burden-notice").hidden = !isConstructionBurdenWork;
      CONSTRUCTION_BURDEN_FIELDS.forEach((fieldName) => {
        const isApplicable = isConstructionBurdenWork
          && visibleFields.has(fieldName)
          && Object.prototype.hasOwnProperty.call(CONSTRUCTION_BURDEN_CONFIG[fieldName].addonAmounts, workType);
        const level = form.elements[fieldName].value || "低";
        const amount = getConstructionBurdenAddonAmount(workType, fieldName, level);
        const addition = root.querySelector(`[data-burden-addition="${fieldName}"]`);
        addition.hidden = !isApplicable;
        addition.textContent = isApplicable ? `現在の選択：${level}　加算額${Math.floor(amount).toLocaleString("ja-JP")}円${level === "中" ? "（現行設定）" : ""}` : "";
        root.querySelector(`[data-burden-hearing="${fieldName}"]`).hidden = !isApplicable;
        root.querySelector(`[data-burden-criteria="${fieldName}"]`).hidden = !isApplicable;
        root.querySelector(`[data-burden-reason="${fieldName}"]`).hidden = !(isApplicable && level !== "低");
        syncHearingManualOverride(fieldName);
      });
    };
    const getBurdenHearingInputs = (fieldName) => Array.from(root.querySelectorAll(`[data-burden-hearing-option="${fieldName}"]`));
    const removeAutoReasonLine = (fieldName) => {
      const reason = form.elements[CONSTRUCTION_BURDEN_CONFIG[fieldName].reasonName];
      reason.value = String(reason.value || "")
        .split(/\r?\n/)
        .filter((line) => !line.startsWith(CONSTRUCTION_BURDEN_AUTO_REASON_PREFIX))
        .join("\n")
        .trim();
    };
    const syncHearingManualOverride = (fieldName) => {
      const result = root.querySelector(`[data-burden-hearing-result="${fieldName}"]`);
      const autoLevel = result?.dataset.autoLevel;
      const baseText = result?.dataset.baseText;
      if (!result || !autoLevel || !baseText) return;
      const currentLevel = form.elements[fieldName].value || "低";
      result.textContent = currentLevel === autoLevel
        ? baseText
        : `${baseText}／現在の選択：${currentLevel}（手動変更）`;
    };
    const applyBurdenHearingDecision = (fieldName, { applyLevel = true, updateReason = true } = {}) => {
      const config = CONSTRUCTION_BURDEN_CONFIG[fieldName];
      const selected = getBurdenHearingInputs(fieldName)
        .filter((input) => input.checked)
        .map((input) => ({ level: input.value, text: config.criteria[input.value] }))
        .filter((entry) => entry.text);
      const result = root.querySelector(`[data-burden-hearing-result="${fieldName}"]`);
      if (!selected.length) {
        delete result.dataset.autoLevel;
        delete result.dataset.baseText;
        result.textContent = "自動判定：未判定（現在の区分は変更していません）";
        if (updateReason) removeAutoReasonLine(fieldName);
        return null;
      }
      const level = selected.reduce((highest, entry) => (
        CONSTRUCTION_BURDEN_LEVEL_ORDER[entry.level] > CONSTRUCTION_BURDEN_LEVEL_ORDER[highest]
          ? entry.level
          : highest
      ), "低");
      if (applyLevel) form.elements[fieldName].value = level;
      if (updateReason) {
        const reason = form.elements[config.reasonName];
        const manualReason = String(reason.value || "")
          .split(/\r?\n/)
          .filter((line) => !line.startsWith(CONSTRUCTION_BURDEN_AUTO_REASON_PREFIX))
          .join("\n")
          .trim();
        const autoReason = `${CONSTRUCTION_BURDEN_AUTO_REASON_PREFIX}区分：${level}／回答：${selected.map((entry) => `${entry.level}：${entry.text}`).join("／")}`;
        reason.value = [autoReason, manualReason].filter(Boolean).join("\n");
      }
      const highestCount = selected.filter((entry) => entry.level === level).length;
      const baseText = `自動判定：${level}（${level}の該当内容${highestCount}件／選択${selected.length}件）`;
      result.dataset.autoLevel = level;
      result.dataset.baseText = baseText;
      syncHearingManualOverride(fieldName);
      return level;
    };
    const restoreBurdenHearingAnswers = (fieldName) => {
      const config = CONSTRUCTION_BURDEN_CONFIG[fieldName];
      const reason = String(form.elements[config.reasonName].value || "");
      const autoReasonLine = reason.split(/\r?\n/).find((line) => line.startsWith(CONSTRUCTION_BURDEN_AUTO_REASON_PREFIX)) || "";
      getBurdenHearingInputs(fieldName).forEach((input) => {
        input.checked = !!autoReasonLine && autoReasonLine.includes(`${input.value}：${config.criteria[input.value]}`);
      });
      applyBurdenHearingDecision(fieldName, { applyLevel: false, updateReason: false });
    };
    const applyWorkTypeUi = (keepApplicationValue = false) => {
      const wt = form.elements.workType.value;
      const visible = visibleFieldsFor(wt);
      root.querySelectorAll('[data-field]').forEach((field) => {
        const name = field.getAttribute('data-field');
        field.style.display = visible.has(name) ? '' : 'none';
        const labelText = labelFor(wt, name);
        const labelElement = field.querySelector?.('[data-field-label]');
        if (labelText && labelElement) labelElement.textContent = labelText;
        else if (labelText) field.firstChild.textContent = labelText;
      });
      activateWorkDifficultyUi(wt);
      const burdenLevels = CONSTRUCTION_BURDEN_WORK_TYPES.has(wt) ? ["低", "中", "高"] : ["低", "高"];
      CONSTRUCTION_BURDEN_FIELDS.forEach((fieldName) => setOptions(form.elements[fieldName], burdenLevels, form.elements[fieldName].value));
      const currentApplication = keepApplicationValue ? form.elements.applicationType.value : '';
      setOptions(form.elements.applicationType, OPTION_SETS.applicationByWorkType[wt] || OPTION_SETS.applicationDefault, currentApplication);
      updateConstructionBurdenUi(wt, visible);
    };
    const run = () => { applyWorkTypeUi(true); const f = form.elements; const base = calculateBase(f); const addons = calculateAddons(f); const addon = addons.reduce((s, x) => s + x.amount, 0), discount = n(f.discount.value), expense = n(f.expense.value), taxable = base + addon - discount; const tax = Math.floor(taxable * (app?.getTaxRate?.() ?? 0.1)), total = taxable + tax + expense; form.dataset.result = JSON.stringify({ base, addons, addon, discount, expense, tax, taxable, total }); root.querySelector('#calc-result').innerHTML = `<p>基本報酬: ${yen(base)}</p><p>加算明細: ${(addons.map(a => `${a.name} ${yen(a.amount)}`).join(' / ') || 'なし')}</p><p>値引き: ${yen(discount)}</p><p>実費: ${yen(expense)}</p><p>消費税: ${yen(tax)}</p><p><strong>合計: ${yen(total)}</strong></p><p class='meta'>${NOTICE}</p>`; };
    const fillForm = (calc) => {
      const f = form.elements;
      const dynamic = calc?.permit_dynamic_answers && typeof calc.permit_dynamic_answers === 'object' ? calc.permit_dynamic_answers : {};
      const as01 = (v) => ['1','あり','要確認','確認済','期限迫る','要','true'].includes(String(v ?? '').trim()) ? '1' : '0';
      const setLevel = (select, value) => {
        const normalized = ["低", "中", "高"].includes(value) ? value : "低";
        select.value = Array.from(select.options).some((option) => option.value === normalized) ? normalized : "低";
      };
      f.clientId.value = calc.client_id || "";
      f.projectName.value = calc.project_name || "";
      f.workType.value = calc.work_type || "建設業許可";
      applyWorkTypeUi(true);
      if (Array.from(f.applicationType.options).some(o => o.value === calc.application_type)) f.applicationType.value = calc.application_type;
      f.corporateType.value = calc.corporate_type || "法人";
      f.governorType.value = calc.governor_type || "知事";
      f.generalSpecific.value = calc.general_specific || "一般";
      f.industryCount.value = n(calc.industry_count) || 1;
      f.officerCount.value = n(calc.officer_count) || 2;
      f.officeCount.value = n(calc.office_count) || 1;
      f.documentLevel.value = calc.document_level || "低";
      f.documentLevelReason.value = calc.document_level_reason || "";
      restoreWorkDifficultyAnswers();
      f.urgent.value = calc.urgent ? "1" : "0";
      setLevel(f.keikan, calc.keikan_level);
      setLevel(f.sengi, calc.sengi_level);
      setLevel(f.zaisan, calc.zaisan_level);
      f.keikanReason.value = calc.keikan_reason || "";
      f.sengiReason.value = calc.sengi_reason || "";
      f.zaisanReason.value = calc.zaisan_reason || "";
      CONSTRUCTION_BURDEN_FIELDS.forEach(restoreBurdenHearingAnswers);
      f.visit.value = calc.visit_required ? "1" : "0";
      f.agent.value = calc.agent_required ? "1" : "0";
      f.urlNotification.value = as01(calc.url_notification || dynamic.urlNotification || dynamic.permitHasWebsite);
      f.selfCertification.value = as01(dynamic.selfCertification || (dynamic.permitStorageCategory === '自己所有' ? '1' : '0'));
      f.usageConsent.value = as01(dynamic.usageConsent || (dynamic.permitStorageCategory === '他人所有' && dynamic.permitParkingProofType === '使用承諾証明書' ? '1' : '0'));
      f.baseLocationCheck.value = as01(dynamic.baseLocationCheck || dynamic.permitBaseLocation);
      f.qualifiedStaffCheck.value = as01(dynamic.qualifiedStaffCheck || dynamic.permitQualifiedCount);
      f.guaranteeAssociation.value = as01(dynamic.guaranteeAssociation);
      f.courseCompletion.value = as01(dynamic.courseCompletion || dynamic.permitCourseCertificate);
      f.renewalDeadline.value = as01(dynamic.renewalDeadline);
      f.expense.value = n(calc.expense_amount) || 0;
      f.discount.value = n(calc.discount_amount) || 0;
      f.memo.value = calc.memo || "";
      run();
    };

    const applyBridgeData = () => {
      const bridge = app?.getEstimateCalculatorBridgeData?.();
      if (!bridge) return;
      fillForm(bridge);
      app?.clearEstimateCalculatorBridgeData?.();
      app?.showMessage?.('許認可ヒアリングの入力値を見積自動算出フォームへ反映しました。');
    };
    const getReflectionInfo = (calc) => ({ reflectedAt: calc?.reflected_at || null, reflectedEstimateId: calc?.reflected_estimate_id || null, isReflected: !!(calc?.reflected_at || calc?.reflected_estimate_id) });
    const reflectToEstimate = async (calc) => { const sb = app?.getSupabaseClient?.(); const u = app?.getCurrentUser?.(); if (!sb || !u) return; const reflectionInfo = getReflectionInfo(calc); if (reflectionInfo.isReflected) { const reflectedDate = reflectionInfo.reflectedAt ? fmtDate(reflectionInfo.reflectedAt) : '日時未記録'; if (!confirm(`この概算見積はすでに見積へ反映済みです。（反映日: ${reflectedDate}）\n再反映しますか？`)) return false; } const clients = app?.getClients?.() || []; const selected = clients.find(c => String(c.id) === String(calc.client_id || '')); const customerName = selected?.name || selected?.companyName || selected?.contactName || calc.project_name || '自動算出'; const customerNote = '本見積は、現時点で確認できる情報に基づくものです。必要資料、申請先、業務範囲により金額が変動する場合があります。'; const est = { user_id: u.id, client_id: calc.client_id || null, customer_name: customerName, estimate_title: calc.project_name || '見積自動算出', estimate_date: new Date().toISOString().slice(0, 10), status: '未回答', memo: customerNote, subtotal: n(calc.taxable_subtotal) + n(calc.expense_amount), tax: n(calc.tax), total: n(calc.total), estimate_number: 'M-AUTO-' + Date.now(), estimate_source: 'auto_calculation' }; const e = await sb.from('estimates').insert(est).select('*').single(); if (e.error) { app.showMessage('見積登録に失敗しました。' + e.error.message, true); return false; } const gyoseiReward = n(calc.base_fee) + n(calc.addon_fee); const discountAmount = -n(calc.discount_amount); const items = [{ item_name: '行政書士報酬', quantity: 1, unit_price: gyoseiReward, amount: gyoseiReward, sort_order: 1 }, { item_name: '実費・法定費用', quantity: 1, unit_price: n(calc.expense_amount), amount: n(calc.expense_amount), sort_order: 2 }, ...(discountAmount < 0 ? [{ item_name: '値引き', quantity: 1, unit_price: discountAmount, amount: discountAmount, sort_order: 3 }] : [])].map(i => ({ ...i, user_id: u.id, estimate_id: e.data.id })); const ir = await sb.from('estimate_items').insert(items).select('*'); if (ir.error) { app.showMessage('見積明細登録に失敗しました。' + ir.error.message, true); return false; } app?.addEstimateToState?.(e.data); app?.addEstimateItemsToState?.(ir.data || []); app?.renderEstimatesNow?.(); if (calc?.id) { const reflectedAt = new Date().toISOString(); const reflected = await sb.from('estimate_calculations').update({ reflected_estimate_id: e.data.id, reflected_at: reflectedAt }).eq('id', calc.id).eq('user_id', u.id); if (reflected.error) { app.showMessage('反映済み情報の保存に失敗しました。' + reflected.error.message, true); return false; } } if (app?.refreshEstimateListData) { await app.refreshEstimateListData(); } else { await app.reloadAllData(); } await loadSavedCalculations(); renderSavedList(); app.showMessage('見積へ反映しました。'); return true; };
    let localEstimateCalculations = [];
    let selectedSavedCalculationIds = [];
    const loadSavedCalculations = async () => { const sb = app?.getSupabaseClient?.(); if (!sb) { localEstimateCalculations = []; return; } const u = await waitForCurrentUser(app); if (!u?.id) { localEstimateCalculations = []; return; } try { const { data, error } = await sb.from('estimate_calculations').select('*').eq('user_id', u.id).order('created_at', { ascending: false }); if (error) throw error; localEstimateCalculations = Array.isArray(data) ? data : []; debugLog("LOAD ESTIMATE CALCULATIONS RESULT", { count: localEstimateCalculations.length }); } catch (error) { console.error("LOAD ESTIMATE CALCULATIONS ERROR", error); localEstimateCalculations = []; app?.showMessage?.('保存済み概算見積の取得に失敗しました。' + error.message, true); } };
    const getCalculationById = (id) => { const appCalcs = app?.getEstimateCalculations?.() || []; const source = appCalcs.length ? appCalcs : localEstimateCalculations; return source.find(x => String(x.id) === String(id)); };
    const toggleSavedCalculationSelection = (id) => { const normalized = String(id); if (!normalized) return; if (selectedSavedCalculationIds.includes(normalized)) selectedSavedCalculationIds = selectedSavedCalculationIds.filter((entry) => entry !== normalized); else selectedSavedCalculationIds = [...selectedSavedCalculationIds, normalized]; };
    const renderSavedList = () => { const appCalcs = app?.getEstimateCalculations?.() || []; const sourceCalcs = appCalcs.length ? appCalcs : localEstimateCalculations; const calcs = sourceCalcs.slice().sort((a, b) => new Date(b.created_at || 0) - new Date(a.created_at || 0)); const idSet = new Set(calcs.map((entry) => String(entry.id))); selectedSavedCalculationIds = selectedSavedCalculationIds.filter((id) => idSet.has(id)); if (!calcs.length) { savedList.innerHTML = '<p class="meta">保存済みデータはありません。</p>'; return; } const selectedCount = selectedSavedCalculationIds.length; savedList.innerHTML = `<div class="row-actions"><button type="button" class="secondary-btn" data-calc-action="select_all">全選択</button><button type="button" class="secondary-btn" data-calc-action="clear_all">全解除</button><button type="button" class="danger-btn" data-calc-action="bulk_delete">選択を一括削除</button><span class="meta">選択件数: ${h(String(selectedCount))}件 ${selectedCount ? '' : '（対象を選択してください）'}</span></div><p class="meta">保存済み概算見積は新しい順で表示しています。</p><div class="saved-calc-list">${calcs.map(c => `<article class="panel saved-calc-card"><div class="saved-calc-header"><label><input type="checkbox" data-calc-action="toggle_select" data-id="${h(c.id)}" ${selectedSavedCalculationIds.includes(String(c.id)) ? 'checked' : ''}/> 選択</label> <strong>${h(c.project_name || "-")}</strong> / <span>${h(c.work_type || "-")}</span> / <strong>${h(yen(c.total || 0))}</strong> ${getReflectionInfo(c).isReflected ? `<span class="status-badge">反映済み</span>` : ''}</div><div class="saved-calc-meta">作成日: ${h(fmtDate(c.created_at))} / 顧客名: ${h(c.client_name || c.customer_name || "-")} / 申請区分: ${h(c.application_type || "-")}</div><div class="saved-calc-money">基本報酬: ${h(yen(c.base_fee || 0))} / 加算: ${h(yen(c.addon_fee || 0))} / 実費: ${h(yen(c.expense_amount || 0))} / 消費税: ${h(yen(c.tax || 0))}</div><div class="saved-calc-note">メモ: ${h(c.memo || "-")}</div><div class="row-actions saved-calc-actions"><button type="button" data-calc-action="detail" data-id="${h(c.id)}" class="secondary-btn">詳細表示</button><button type="button" data-calc-action="reload" data-id="${h(c.id)}" class="secondary-btn">フォームに再読込</button><button type="button" data-calc-action="reflect" data-id="${h(c.id)}" class="secondary-btn">見積へ反映</button><button type="button" data-calc-action="print" data-id="${h(c.id)}" class="secondary-btn">概算書出力</button><button type="button" data-calc-action="delete" data-id="${h(c.id)}" class="danger-btn">削除</button></div></article>`).join('')}</div>`; };
    form.elements.workType.addEventListener('change', () => { applyWorkTypeUi(false); run(); });
    form.elements.documentLevel.addEventListener("change", () => {
      updateWorkDifficultySelectionUi(form.elements.workType.value);
      run();
    });
    root.querySelector("[data-work-difficulty-clear]").addEventListener("click", () => {
      getWorkDifficultyInputs().forEach((input) => { input.checked = false; });
      applyWorkDifficultyDecision();
      run();
    });
    CONSTRUCTION_BURDEN_FIELDS.forEach((fieldName) => {
      form.elements[fieldName].addEventListener('change', () => {
        syncHearingManualOverride(fieldName);
        run();
      });
      getBurdenHearingInputs(fieldName).forEach((input) => input.addEventListener('change', () => {
        applyBurdenHearingDecision(fieldName);
        run();
      }));
      root.querySelector(`[data-burden-hearing-clear="${fieldName}"]`).addEventListener('click', () => {
        getBurdenHearingInputs(fieldName).forEach((input) => { input.checked = false; });
        applyBurdenHearingDecision(fieldName);
        run();
      });
    });
    root.querySelector('#calc-run').addEventListener('click', run);
    root.querySelector('#calc-save').addEventListener('click', async () => {
      const sb = app?.getSupabaseClient?.();
      const u = app?.getCurrentUser?.();
      if (!sb || !u) return;
      run();
      const r = JSON.parse(form.dataset.result || '{}');
      const f = form.elements;
      const memoSuffix = n(f.urlNotification.value) > 0 ? 'URL届出: あり' : '';
      const mergedMemo = [f.memo.value || '', memoSuffix].filter(Boolean).join('\n');
      const workType = f.workType.value;
      const reasonOrNull = (field) => String(field?.value || "").trim() || null;
      const rawPayload = {
        user_id: u.id,
        client_id: f.clientId.value || null,
        project_name: f.projectName.value || null,
        work_type: workType,
        application_type: f.applicationType.value,
        corporate_type: f.corporateType.value,
        governor_type: f.governorType.value,
        general_specific: f.generalSpecific.value,
        industry_count: n(f.industryCount.value),
        officer_count: n(f.officerCount.value),
        office_count: n(f.officeCount.value),
        document_level: f.documentLevel.value,
        document_level_reason: reasonOrNull(f.documentLevelReason),
        urgent: n(f.urgent.value) > 0,
        keikan_level: f.keikan.value,
        sengi_level: f.sengi.value,
        zaisan_level: f.zaisan.value,
        keikan_reason: workType === "建設業許可" ? reasonOrNull(f.keikanReason) : null,
        sengi_reason: CONSTRUCTION_BURDEN_WORK_TYPES.has(workType) ? reasonOrNull(f.sengiReason) : null,
        zaisan_reason: workType === "建設業許可" ? reasonOrNull(f.zaisanReason) : null,
        expense_amount: r.expense || 0,
        discount_amount: r.discount || 0,
        memo: mergedMemo || null,
        base_fee: r.base || 0,
        addon_fee: r.addon || 0,
        taxable_subtotal: r.taxable || 0,
        tax: r.tax || 0,
        total: r.total || 0,
        addon_breakdown: JSON.stringify(r.addons || []),
      };
      const payload = pickObjectKeys(rawPayload, ESTIMATE_CALCULATION_MUTATION_COLUMNS);
      const res = await sb.from('estimate_calculations').insert(payload);
      if (res.error) {
        app.showMessage(buildSaveErrorMessage('見積自動算出', res.error), true);
        return;
      }
      await app.reloadAllData();
      await loadSavedCalculations();
      renderSavedList();
      app.showMessage('見積自動算出を保存しました。');
    });
    root.querySelector('#calc-apply').addEventListener('click', async () => { run(); const r = JSON.parse(form.dataset.result || '{}'); const f = form.elements; await reflectToEstimate({ client_id: f.clientId.value || null, project_name: f.projectName.value || null, memo: f.memo.value || null, base_fee: r.base || 0, addon_fee: r.addon || 0, discount_amount: r.discount || 0, expense_amount: r.expense || 0, taxable_subtotal: r.taxable || 0, tax: r.tax || 0, total: r.total || 0 }); });
    savedList.addEventListener('click', async (event) => { const actionElement = event.target.closest('[data-calc-action]'); if (!actionElement) return; const action = actionElement.dataset.calcAction; const targetId = actionElement.dataset.id; if (action === 'toggle_select') { toggleSavedCalculationSelection(targetId); renderSavedList(); return; } if (action === 'select_all') { const appCalcs = app?.getEstimateCalculations?.() || []; const sourceCalcs = appCalcs.length ? appCalcs : localEstimateCalculations; selectedSavedCalculationIds = sourceCalcs.map((entry) => String(entry.id)); renderSavedList(); return; } if (action === 'clear_all') { selectedSavedCalculationIds = []; renderSavedList(); return; } if (action === 'bulk_delete') { if (!selectedSavedCalculationIds.length) { app.showMessage('対象を選択してください。', true); return; } if (!confirm(`選択した${selectedSavedCalculationIds.length}件を削除しますか？`)) return; const sb = app?.getSupabaseClient?.(); if (!sb) return; const del = await sb.from('estimate_calculations').delete().in('id', selectedSavedCalculationIds); if (del.error) { app.showMessage('選択データの削除に失敗しました。' + del.error.message, true); return; } selectedSavedCalculationIds = []; await loadSavedCalculations(); renderSavedList(); if (app?.reloadAllData) await app.reloadAllData(); app.showMessage('選択した保存済み算出データを削除しました。'); return; } const btn = event.target.closest('button[data-calc-action]'); if (!btn) return; const calc = getCalculationById(btn.dataset.id); if (!calc) return; if (action === 'detail') { const addon = parseAddonBreakdown(calc.addon_breakdown).map(a => `${a.name}: ${yen(a.amount)}`).join(' / ') || 'なし'; alert(`作成日: ${fmtDate(calc.created_at)}\n顧客: ${calc.client_name || '-'}\n案件: ${calc.project_name || '-'}\n業務種別: ${calc.work_type || '-'}\n申請区分: ${calc.application_type || '-'}\n基本報酬: ${yen(calc.base_fee || 0)}\n加算: ${yen(calc.addon_fee || 0)}\n値引き: ${yen(calc.discount_amount || 0)}\n実費: ${yen(calc.expense_amount || 0)}\n消費税: ${yen(calc.tax || 0)}\n合計: ${yen(calc.total || 0)}\n加算明細: ${addon}\nメモ: ${calc.memo || '-'}`); } else if (action === 'reload') { fillForm(calc); app.showMessage('保存済みデータをフォームへ再読込しました。'); } else if (action === 'reflect') { await reflectToEstimate(calc); } else if (action === 'print') { openPrintWindow(app, calc); } else if (action === 'delete') { if (!confirm('この保存済み算出データを削除しますか？')) return; const sb = app?.getSupabaseClient?.(); if (!sb) return; const del = await sb.from('estimate_calculations').delete().eq('id', calc.id); if (del.error) { app.showMessage('保存済み算出データの削除に失敗しました。' + del.error.message, true); return; } await loadSavedCalculations(); renderSavedList(); if (app?.reloadAllData) await app.reloadAllData(); app.showMessage('保存済み算出データを削除しました。'); } });
    applyWorkTypeUi(false); run(); applyBridgeData(); window.addEventListener('estimate-calculator-bridge-updated', applyBridgeData); if (app?.reloadAllData) await app.reloadAllData(); await loadSavedCalculations(); renderSavedList();
  }
  document.addEventListener('DOMContentLoaded', () => { init(); });
})();
