/**
 * 型號解讀＋替代品對照（只有資料和計算，不碰畫面；可以直接用 node 測）
 * 只放查證過的規則，出處：
 * - 乾燥過濾器型號：Danfoss DML／DCL 規格書（AI192386435941，2021.03）
 * - 各牌乾燥過濾器對照：Sanhua cross-reference（2021-04）
 * - Danfoss 膨脹閥：T2＝內平衡、TE2＝外平衡（Danfoss T2／TE2 規格書）；X＝R22、N＝R134a、S＝R404A／R507
 * - KP 壓力開關範圍：Danfoss KP 規格書
 * - 排×支×鏡面、銅管分數：上課錄音（錄音05、09、10）
 * - HFC 管制：環境部「氫氟碳化物管理辦法」（2025-02-25 發布）
 */

/** 比對用：全形轉半形、小寫、去掉空白和標點（「DML 083 S」「dml-083s」都變成 dml083s） */
export const canon = (s) =>
  String(s ?? '')
    .normalize('NFKC')
    .toLowerCase()
    .replace(/[\s\-_/.,，、・·()（）[\]【】"'“”‘’]/g, '')
export const modelKey = (m) => canon(m)
/** 規格拆成詞再排序：「三通 3分」和「3分 三通」視為一樣 */
const specKey = (s) =>
  String(s ?? '')
    .normalize('NFKC')
    .split(/[\s,，、・·/]+/)
    .map(canon)
    .filter(Boolean)
    .sort()
    .join('+')
/** 寬鬆的品項鍵：品名｜規格（詞排序）｜品牌｜型號 */
export const looseKey = (f) => [canon(f.label), specKey(f.spec), canon(f.brand), modelKey(f.model)].join('|')

// ───────────── 銅管分數 ─────────────
/** 1 分＝1/8 吋；ACR 銅管外徑（錄音05） */
const PIPE = {
  2: ['1/4″', '2分', 6.35],
  3: ['3/8″', '3分', 9.52],
  4: ['1/2″', '4分', 12.7],
  5: ['5/8″', '5分', 15.88],
  6: ['3/4″', '6分', 19.05],
  7: ['7/8″', '7分', 22.22],
  9: ['1-1/8″', '1吋1分', 28.58],
  11: ['1-3/8″', '1吋3分', 34.92],
  13: ['1-5/8″', '1吋5分', 41.28],
}
export const pipeName = (eighths) => PIPE[eighths]?.[1] || ''

// ───────────── 乾燥過濾器 ─────────────
/** Sanhua 對照表有列的尺寸（前兩碼＝立方吋、最後一碼＝接頭幾個 1/8 吋） */
const FD = {
  flare: ['032', '052', '082', '162', '053', '083', '163', '303', '164', '304', '414', '165', '305', '415'],
  solder: ['032', '052', '082', '033', '053', '083', '163', '303', '084', '164', '304', '414', '165', '305', '415', '417', '757', '759'],
  biFlare: ['083', '163', '084', '164', '304', '165'],
  biSolder: ['083', '163', '303', '084', '164', '304', '165', '305', '307'],
}
const BRAND_OF = { fd: 'Sanhua', fdbi: 'Sanhua', dml: 'Danfoss', dcl: 'Danfoss', dmb: 'Danfoss', dcb: 'Danfoss', ek: 'Emerson', ekz: 'Emerson', adk: 'Emerson', bfk: 'Emerson', bfkz: 'Emerson', c: 'Sporlan', hpc: 'Sporlan' }
const BI = new Set(['fdbi', 'dmb', 'dcb', 'bfk', 'bfkz', 'hpc'])
const DRIER_RE = /\b(FDBI|FD|DML|DCL|DMB|DCB|EKZ|EK|ADK|BFKZ|BFK)\s*-?\s*(\d{3})\s*-?\s*(S)?(?![0-9a-z])|\b(C|HPC)-(\d{3})(-S)?(?![0-9a-z])/i

/** 統一寫法：Danfoss／Emerson「DML 083S」、Sanhua「FD-083-S」、Sporlan「C-083-S」 */
function drierModel(prefix, code, solder) {
  const p = prefix.toLowerCase()
  if (p === 'fd' || p === 'fdbi' || p === 'c' || p === 'hpc') return `${prefix.toUpperCase()}-${code}${solder ? '-S' : ''}`
  return `${prefix.toUpperCase()} ${code}${solder ? 'S' : ''}`
}
function drierEquivalents(code, solder, bi) {
  const list = bi ? (solder ? FD.biSolder : FD.biFlare) : solder ? FD.solder : FD.flare
  if (!list.includes(code)) return []
  if (bi) {
    const out = [{ brand: 'Sanhua', model: drierModel('FDBI', code, solder) }]
    if (!(solder && code === '303')) out.push({ brand: 'Danfoss', model: drierModel('DMB', code, solder) }, { brand: 'Danfoss', model: drierModel('DCB', code, solder) })
    out.push({ brand: 'Emerson', model: drierModel('BFK', code, solder) })
    return out
  }
  // Sporlan 75 立方吋叫 60x（Sanhua 對照表：757S＝C-607-S、759S＝C-609-S）
  const sporlan = { 757: '607', 759: '609' }[code] || code
  return [
    { brand: 'Sanhua', model: drierModel('FD', code, solder) },
    { brand: 'Danfoss', model: drierModel('DML', code, solder) },
    { brand: 'Danfoss', model: drierModel('DCL', code, solder) },
    { brand: 'Emerson', model: drierModel('EK', code, solder) },
    { brand: 'Emerson', model: drierModel('ADK', code, solder) },
    { brand: 'Sporlan', model: drierModel('C', sporlan, solder) },
  ]
}
function decodeDrier(text) {
  const m = DRIER_RE.exec(text)
  if (!m) return null
  const prefix = m[1] || m[4]
  const code = m[2] || m[5]
  const solder = !!(m[3] || m[6])
  const p = prefix.toLowerCase()
  const bi = BI.has(p)
  const vol = parseInt(code.slice(0, -1), 10)
  const pipe = PIPE[Number(code.slice(-1))]
  if (!vol || !pipe) return null
  const model = drierModel(prefix, code, solder)
  const facts = [
    `乾燥劑 ${vol} 立方吋（型號前兩碼）`,
    `接頭 ${pipe[0]}＝${pipe[1]}（最後一碼，以 1/8 吋計）`,
    solder ? '尾巴有 S＝焊接（ODF）' : '尾巴沒有 S＝喇叭口（SAE flare）',
  ]
  if (p === 'dml' || p === 'dmb') facts.push('核心：100% 分子篩（DML）')
  if (p === 'dcl' || p === 'dcb') facts.push('核心：80% 分子篩＋20% 活性氧化鋁（DCL）')
  if (bi) facts.push('雙向型：熱泵、冷媒會反向流的系統用')
  const equivalents = drierEquivalents(code, solder, bi).filter((e) => modelKey(e.model) !== modelKey(model))
  return {
    kind: 'drier',
    label: '乾燥過濾器',
    brand: BRAND_OF[p],
    model,
    spec: `${pipe[1]}・${solder ? '焊接' : '喇叭口'}`,
    title: `乾燥過濾器 ${vol} 立方吋・${pipe[1]}・${solder ? '焊接' : '喇叭口'}${bi ? '・雙向' : ''}`,
    facts,
    equivalents,
    note: equivalents.length
      ? '同尺寸、同接頭的各牌可以互換。Danfoss 建議：POE／PAG 冷凍油有添加劑時用 DML，不用 DCL。'
      : '這個尺寸不在對照表裡，請查原廠資料。',
    source: 'Danfoss DML／DCL 規格書、Sanhua 對照表',
  }
}

// ───────────── Danfoss 熱力膨脹閥 ─────────────
const TXV_REF = { x: 'R22', n: 'R134a', s: 'R404A／R507' }
const TXV_RE = /\bT(E)?\s?([XNS])\s?-?(2|5|12|20|55)(?![0-9])/i
function decodeTxv(text) {
  const m = TXV_RE.exec(text)
  if (!m) return null
  const ext = !!m[1]
  const ref = TXV_REF[m[2].toLowerCase()]
  const model = `T${ext ? 'E' : ''}${m[2].toUpperCase()} ${m[3]}`
  return {
    kind: 'txv',
    label: '膨脹閥',
    brand: 'Danfoss',
    model,
    spec: ref,
    title: `熱力膨脹閥・${ref} 用・${ext ? '外' : '內'}平衡`,
    facts: [`字母 ${m[2].toUpperCase()}＝${ref} 用`, ext ? '有 E＝外平衡（要接外平衡管）' : '沒有 E＝內平衡', m[3] === '2' ? '閥芯（孔口）另外選，依能力配' : `${m[3]} 號閥體`],
    equivalents: [],
    note: '換膨脹閥要看冷媒、內外平衡、能力（閥芯）都對。',
    source: 'Danfoss T2／TE2 規格書',
  }
}

// ───────────── Danfoss KP 壓力開關 ─────────────
const KP = {
  1: ['低壓開關', '可調範圍 -0.2～7.5 bar'],
  5: ['高壓開關', '可調範圍 8～32 bar'],
  15: ['高低壓開關（雙壓）', '低壓 -0.2～7.5 bar、高壓 8～32 bar'],
}
function decodeKp(text) {
  const m = /\bKP\s?-?(15|1|5)(?![0-9])/i.exec(text)
  if (!m) return null
  const [name, range] = KP[m[1]]
  return { kind: 'kp', label: '壓力開關', brand: 'Danfoss', model: `KP ${m[1]}`, spec: name, title: `壓力開關 KP ${m[1]}＝${name}`, facts: [range], equivalents: [], note: '', source: 'Danfoss KP 規格書' }
}

// ───────────── Danfoss EVR 電磁閥 ─────────────
function decodeEvr(text) {
  const m = /\bEVR\s?-?(\d{1,2})(?![0-9])/i.exec(text)
  if (!m) return null
  return { kind: 'evr', label: '電磁閥', brand: 'Danfoss', model: `EVR ${m[1]}`, spec: '', title: `電磁閥 EVR ${m[1]}`, facts: ['數字越大，閥越大', '線圈另外配（看電壓）'], equivalents: [], note: '', source: 'Danfoss EVR 型錄' }
}

// ───────────── 散熱器、冷排：排×支×鏡面 ─────────────
function decodeCoil(text) {
  const m = /(?<![0-9])(\d{1,2})\s*[x×X*＊]\s*(\d{1,2})\s*[x×X*＊]\s*(\d{3,4})(?![0-9])/.exec(String(text).normalize('NFKC'))
  if (!m) return null
  const [r, t, l] = [Number(m[1]), Number(m[2]), Number(m[3])]
  const facts = [`${r} 排（一層一層）`, `每排 ${t} 支（找平的彎頭，順著數）`, `鏡面（實內）${l} mm：有鰭片、風吹得到的長度`]
  if (t === 11) facts.push('11 支約 29 公分高，配 10 吋風車')
  if (t === 14) facts.push('14 支約 36.5 公分高，配 12 吋風車')
  if (r === 4 && t === 11 && l === 330) facts.push('老闆說：4×11×330 約 1 馬')
  return { kind: 'coil', label: '冷凝器', brand: '', model: '', spec: `${r}×${t}×${l}`, title: `${r} 排 × ${t} 支 × 鏡面 ${l} mm`, facts, equivalents: [], note: '壓縮機 1 馬配散熱器 2 馬（估價先 ×2）。', source: '上課錄音（錄音09、10）' }
}

// ───────────── 銅管、接頭的分數 ─────────────
function decodePipe(text) {
  const m = /(?:(\d)\s*吋\s*)?(\d)\s*分/.exec(String(text).normalize('NFKC'))
  if (!m) return null
  const eighths = Number(m[1] || 0) * 8 + Number(m[2])
  const p = PIPE[eighths]
  if (!p) return null
  return { kind: 'pipe', label: '', brand: '', model: '', spec: p[1], title: `${p[1]}＝${p[0]}＝外徑 ${p[2]} mm`, facts: ['1 分＝1/8 吋；1 吋＝8 分＝25.4 mm'], equivalents: [], note: '', source: '上課錄音（錄音05）' }
}

// ───────────── 冷媒 ─────────────
const REF_INFO = {
  22: 'HCFC 舊冷媒，蒙特婁議定書逐步淘汰中',
  32: 'HFC：2026 起環境部管制總量',
  '134a': 'HFC：2026 起環境部管制總量',
  '404a': 'HFC 混合冷媒：2026 起環境部管制總量',
  '407c': 'HFC 混合冷媒：2026 起環境部管制總量',
  '410a': 'HFC 混合冷媒：2026 起環境部管制總量',
  '417a': 'HFC 混合冷媒（含碳氫）：2026 起環境部管制總量',
  507: 'HFC 混合冷媒：2026 起環境部管制總量',
  '507a': 'HFC 混合冷媒：2026 起環境部管制總量',
  '448a': 'HFC／HFO 混合（含 HFC，在管制範圍）',
  '449a': 'HFC／HFO 混合（含 HFC，在管制範圍）',
  '513a': 'HFO／HFC 混合（含 R134a，在管制範圍）',
  '600a': '碳氫冷媒（異丁烷），可燃（A3）',
  290: '碳氫冷媒（丙烷），可燃（A3）',
  744: 'CO₂（二氧化碳）',
  '1234yf': 'HFO',
}
function decodeRefrigerant(text) {
  const m = /\bR-?\s?(1234yf|600a|134a|404A|407C|410A|417A|448A|449A|507A?|513A|22|32|290|744)(?![0-9a-z])/i.exec(text)
  if (!m) return null
  const code = m[1].toLowerCase()
  const name = `R${['134a', '600a', '1234yf'].includes(code) ? code : code.toUpperCase()}`
  return { kind: 'ref', label: '冷媒', brand: '', model: '', spec: name, title: `冷媒 ${name}`, facts: [REF_INFO[code]], equivalents: [], note: '', source: '環境部氫氟碳化物管理辦法' }
}

// ───────────── Danfoss 訂購碼（例如 023Z5040） ─────────────
function decodeDanfossCode(text) {
  const m = /\b(0\d{2}[A-Z]\d{4})\b/.exec(String(text).toUpperCase())
  if (!m) return null
  return { kind: 'code', label: '', brand: 'Danfoss', model: '', spec: '', code: m[1], title: `Danfoss 訂購碼 ${m[1]}`, facts: ['到 Danfoss Product Store 輸入這個號碼，可以查規格和文件'], equivalents: [], note: '', source: 'Danfoss Product Store' }
}

const DECODERS = [decodeDrier, decodeTxv, decodeKp, decodeEvr, decodeCoil, decodeRefrigerant, decodeDanfossCode, decodePipe]

/** 一段文字（型號、規格、標籤上的字）可能看得出好幾件事；同一種只留第一個 */
export function decode(...texts) {
  const text = texts.filter(Boolean).join(' ')
  const out = []
  for (const fn of DECODERS) {
    const r = fn(text)
    if (r && !out.some((x) => x.kind === r.kind)) out.push(r)
  }
  return out
}

/** 型號統一寫法（例如 AI 寫「DML083S」「dml 083 s」→「DML 083S」）；看不懂就原樣 */
export function normalizeModel(model) {
  const s = String(model ?? '').trim()
  if (!s) return s
  for (const fn of [decodeDrier, decodeTxv, decodeKp, decodeEvr]) {
    const r = fn(s)
    if (r?.model) return r.model
  }
  return s
}

/** 原廠資料連結（官方工具＋搜尋） */
export function linksFor({ brand = '', model = '', label = '' } = {}, decoded = []) {
  const q = [brand, model || label].filter(Boolean).join(' ')
  const out = []
  if (q) out.push({ title: `搜尋「${q}」規格書`, url: `https://www.google.com/search?q=${encodeURIComponent(`${q} datasheet`)}` })
  const kinds = new Set(decoded.map((d) => d.kind))
  const danfoss = /danfoss/i.test(brand) || decoded.some((d) => d.brand === 'Danfoss')
  if (danfoss) out.push({ title: 'Danfoss Product Store（輸入型號或訂購碼）', url: 'https://store.danfoss.com/en/' })
  if (kinds.has('drier')) out.push({ title: 'Sanhua 各牌對照表（PDF）', url: 'https://cdn.sanhuaeurope.co.uk/new_content/static/uploads/files/catalogue/sanhua-am-crossref-final-010421-web-pdf-1634564345.pdf' })
  if (/copeland|emerson/i.test(brand) || /壓縮機/.test(label)) out.push({ title: 'Copeland OPI（壓縮機規格、替代型號）', url: 'https://www.copeland.com/en-in/tools-resources/hvac-and-refrigeration-software-tools/online-product-information' })
  if (danfoss || kinds.has('txv') || kinds.has('drier')) out.push({ title: 'Danfoss Coolselector2（免費選型）', url: 'https://www.danfoss.com/en-in/service-and-support/downloads/dcs/coolselector-2/' })
  return out
}
