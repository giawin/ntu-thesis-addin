// personalize.mjs — bien ban VBA goc cua CuongBC195/add-in-office thanh ban NTU.
//
// Muc dich: dong bo tu kho goc (GitHub Actions) khong con phai "ap lai bang tay
// bay diem ca nhan hoa". Script nay lam dung viec do mot cach tat dinh.
//
// Dung: node personalize.mjs <nguon.bas> <dich.vba>
//
// Tinh chat: IDEMPOTENT. Chay tren ban goc hay tren ban da ca nhan hoa deu cho
// cung mot ket qua. Vi vay workflow co the chay lai bao nhieu lan cung duoc.

import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { dirname } from 'node:path';

// ---------------------------------------------------------------------------
// 1. Hang so ca nhan hoa
// ---------------------------------------------------------------------------
const APP_ID = 'NTUThesisAddin';                  // thay "OfficePartySuite"
export const VERSION = '1.0.2';                    // CN_PhienBan
const REPO = 'giawin/ntu-thesis-addin';
const MANIFEST_URL =
  `https://github.com/${REPO}/releases/latest/download/version.json`;
const ADDIN_FILE = 'giatwin_h_ntu_thesis_word.dotm'; // CN_DungFileAddIn
const AUTHOR_LINE = 'Created by Add-in NTU Thesis';
const UPDATE_TITLE = 'Cap nhat NTU Thesis Add-in';

// ---------------------------------------------------------------------------
// 2. Thay the theo chuoi (chay truoc khi xu ly than ham)
// ---------------------------------------------------------------------------
const STRING_RULES = [
  // Ten chuong trinh goc xuat hien o 9 cho: hang so TT_SETTING_APP, appName
  // trong CheckLicenseActivation, bon loi goi GetSetting/SaveSetting, thu muc
  // %APPDATA%\OfficePartySuite va hai dong lenh trong file .cmd tu cap nhat.
  // Thay mot lan cho ca 9 cho — dung mot phep thay toan cuc thay vi sau phep
  // thay riêng le, de khong bo sot cho nao (da tung bo sot dong 8314).
  ['OfficePartySuite', APP_ID],

  // Dong ghi cong trong MsgBox
  ['"Created by VPDU xa Thanh Phu',
   `"${AUTHOR_LINE}`],

  // Tieu de cac hop thoai cap nhat
  ['"Cap nhat Add-in"', `"${UPDATE_TITLE}"`],
];

const LINE_RULES = [
  // Phien ban add-in
  [/^CN_PhienBan = ".*"$/m,        `CN_PhienBan = "${VERSION}"`],
  // Dia chi manifest cap nhat
  [/^CN_ManifestUrl = ".*"$/m,     `CN_ManifestUrl = "${MANIFEST_URL}"`],
  // Ten file add-in phai khop de tu cap nhat dung cho
  [/^CN_DungFileAddIn = \(n = ".*"\)$/m,
   `CN_DungFileAddIn = (n = "${ADDIN_FILE}")`],
];

// ---------------------------------------------------------------------------
// 3. Thay toan bo than ham theo ten
// ---------------------------------------------------------------------------
// Moi phan tu: [ten ham, dong khai bao, cac dong than moi]
// Ham duoc tim theo dong khai bao, ket thuc o "End Function" dau tien sau do.
const FUNCTION_BODIES = [
  {
    decl: 'Private Function LicApiUrl() As String',
    body: ['LicApiUrl = ""'],
  },
  {
    decl: 'Private Function CheckLicenseActivation() As Boolean',
    body: [
      "    ' Ban NTU: bo co che ban quyen theo may cua tac gia goc.",
      '    CheckLicenseActivation = True',
    ],
  },
];

// ---------------------------------------------------------------------------
// 4. Xu ly
// ---------------------------------------------------------------------------

/** Chuan hoa xuong dong ve LF de so sanh on dinh. */
function toLf(text) {
  return text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
}

/** Thay than mot ham VBA, giu nguyen dong khai bao va dong "End Function". */
function replaceFunctionBody(lines, decl, newBody, log) {
  const start = lines.findIndex((l) => l.trim() === decl);
  if (start === -1) {
    throw new Error(`Khong tim thay khai bao ham: ${decl}`);
  }
  // Tim "End Function" dau tien sau dong khai bao.
  let end = -1;
  for (let i = start + 1; i < lines.length; i++) {
    if (lines[i].trim() === 'End Function') { end = i; break; }
  }
  if (end === -1) {
    throw new Error(`Khong tim thay "End Function" cua: ${decl}`);
  }
  const removed = end - start - 1;
  lines.splice(start + 1, removed, ...newBody);
  log.push(`  ${decl} — thay ${removed} dong than bang ${newBody.length} dong`);
}

export function personalize(sourceText, { verbose = false } = {}) {
  const log = [];
  let text = toLf(sourceText);

  // --- Thay theo chuoi ---
  for (const [from, to] of STRING_RULES) {
    const truoc = text;
    text = text.split(from).join(to);
    if (truoc !== text && verbose) log.push(`  chuoi: ${from.slice(0, 48)}`);
  }

  // --- Thay theo dong (regex) ---
  for (const [re, to] of LINE_RULES) {
    text = text.replace(re, to);
  }

  // --- Thay than ham ---
  const lines = text.split('\n');
  for (const { decl, body } of FUNCTION_BODIES) {
    replaceFunctionBody(lines, decl, body, log);
  }
  text = lines.join('\n');

  // --- Kiem tra hau dieu kien ---
  const mustNot = [
    'OfficePartySuite',
    'script.google.com',
    'AKfycbxSZyf',
    'Thanh Phu',
    'OfficePartySuite',
    '"Cap nhat Add-in"',
  ];
  for (const bad of mustNot) {
    if (text.includes(bad)) {
      throw new Error(`Con sot chuoi goc sau khi ca nhan hoa: ${bad}`);
    }
  }

  const mustHave = [
    `TT_SETTING_APP As String = "${APP_ID}"`,
    ADDIN_FILE,
    `CN_PhienBan = "${VERSION}"`,
    `CN_ManifestUrl = "${MANIFEST_URL}"`,
    AUTHOR_LINE,
    UPDATE_TITLE,
    'CheckLicenseActivation = True',
    'LicApiUrl = ""',
  ];
  for (const need of mustHave) {
    if (!text.includes(need)) {
      throw new Error(`Thieu gia tri ca nhan hoa bat buoc: ${need}`);
    }
  }

  // Dau ra phai la ASCII thuan: VBA luu ma nguon theo codepage ANSI, nen moi
  // ky tu > 127 phai duoc viet bang ChrW() tu truoc (xem vba2bas.mjs).
  const bad = [...text].findIndex((ch) => ch.charCodeAt(0) > 127);
  if (bad !== -1) {
    const around = text.slice(Math.max(0, bad - 40), bad + 40);
    throw new Error(
      `Ma nguon con ky tu ngoai ASCII (vi tri ${bad}): ...${around}...`
    );
  }

  // VBA can CRLF.
  const out = text.replace(/\n/g, '\r\n');
  return { text: out, log };
}

// ---------------------------------------------------------------------------
// 5. Chay truc tiep
// ---------------------------------------------------------------------------
const isMain = process.argv[1] &&
  (process.argv[1].endsWith('personalize.mjs') ||
   process.argv[1].endsWith('personalize.mjs'.replace(/\//g, '\\')));

if (isMain) {
  const [, , src, dst] = process.argv;
  if (!src || !dst) {
    console.error('Dung: node personalize.mjs <nguon.bas> <dich.vba>');
    process.exit(2);
  }
  const raw = readFileSync(src, 'utf8');
  const { text, log } = personalize(raw, { verbose: true });
  mkdirSync(dirname(dst), { recursive: true });
  writeFileSync(dst, text, 'utf8');
  console.log(`Nguon : ${src}`);
  console.log(`Dich  : ${dst}`);
  for (const line of log) console.log(line);
  console.log(`OK -> ${text.length} ky tu, ASCII thuan`);
}