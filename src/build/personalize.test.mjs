// personalize.test.mjs — tu kiem tra bo bien doi ca nhan hoa.
//
// Chay: node personalize.test.mjs
//
// Ba tinh chat phai giu:
//   1. Tao lai DUNG TUNG BYTE ban Module1.vba da kiem chung tu ban goc.
//   2. Idempotent: chay tren ban da ca nhan hoa khong doi gi.
//   3. Tu choi dau vao khong hop le (thieu ham, con chuoi goc, ky tu >127).

import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { personalize } from './personalize.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const repo = join(here, '..', '..');
const UPSTREAM = join(repo, 'src', 'core', 'word', 'Module1.upstream.bas');
const EXPECTED = join(repo, 'src', 'core', 'word', 'Module1.vba');

let pass = 0;
let fail = 0;

function check(ten, fn) {
  try {
    fn();
    console.log(`  PASS  ${ten}`);
    pass++;
  } catch (e) {
    console.error(`  FAIL  ${ten}`);
    console.error(`        ${e.message.split('\n')[0]}`);
    fail++;
  }
}

const upstream = readFileSync(UPSTREAM, 'utf8');
const expected = readFileSync(EXPECTED, 'utf8');
const built = personalize(upstream).text;

// --- 1. Tai tao chinh xac ---
check('tao lai byte-for-byte Module1.vba tu ban goc', () => {
  assert.equal(
    built.length,
    expected.length,
    `do dai lech: ${built.length} vs ${expected.length}`
  );
  assert.ok(
    built === expected,
    'noi dung khac ban da kiem chung — sua personalize.mjs cho khop, ' +
      'KHONG sua Module1.vba'
  );
});

// --- 2. Idempotent ---
check('chay hai lan cho cung ket qua', () => {
  const lan2 = personalize(built).text;
  assert.equal(lan2, built, 'lan hai khac lan mot');
});

// --- 3. Gia tri ca nhan hoa da co ---
check('khong con chuoi cua tac gia goc', () => {
  for (const bad of [
    'OfficePartySuite',
    'script.google.com',
    'AKfycbxSZyf',
    'Thanh Phu',
    'chuanhoa_vanbandang.dotm',
  ]) {
    assert.ok(!built.includes(bad), `con sot: ${bad}`);
  }
});

check('co du gia tri cua ban NTU', () => {
  for (const need of [
    'TT_SETTING_APP As String = "NTUThesisAddin"',
    'giatwin_h_ntu_thesis_word.dotm',
    'CN_PhienBan = "1.0.1"',
    'giawin/ntu-thesis-addin/releases/latest/download/version.json',
    'Created by Add-in NTU Thesis',
    'CheckLicenseActivation = True',
  ]) {
    assert.ok(built.includes(need), `thieu: ${need}`);
  }
});

check('ma nguon la ASCII thuan', () => {
  const bad = [...built].findIndex((c) => c.charCodeAt(0) > 127);
  assert.equal(bad, -1, `ky tu ngoai ASCII tai vi tri ${bad}`);
});

check('dung CRLF, khong co LF don le', () => {
  const lone = built.replace(/\r\n/g, '').includes('\n');
  assert.ok(!lone, 'con LF don le');
});

// --- 4. Tu choi dau vao hong ---
check('tu choi khi thieu ham CheckLicenseActivation', () => {
  const broken = upstream.replace(
    'Private Function CheckLicenseActivation() As Boolean',
    'Private Function KhongConDay() As Boolean'
  );
  assert.throws(() => personalize(broken), /Khong tim thay khai bao ham/);
});

check('tu choi khi con sot chuoi goc', () => {
  // Chen lai mot lan xuat cua ten cu vao noi khong bi luat thay the cham toi.
  // Cach kiem: doi gia tri APP_ID gia lap bang cach chay tren van ban co chuoi
  // goc nam trong mot ham khong duoc thay the.
  const coChuoiGoc = upstream.replace(
    'Private Function CN_SoSanh',
    'Private Function CN_SoSanh_Test()\n    Dim x: x = "Thanh Phu"\nEnd Function\n\nPrivate Function CN_SoSanh'
  );
  assert.throws(() => personalize(coChuoiGoc), /Con sot chuoi goc/);
});

check('tu choi khi co ky tu ngoai ASCII', () => {
  // Chen mot chu thich co dau tieng Viet — VBA luu ma nguon theo codepage ANSI
  // nen bat ky ky tu >127 nao cung phai bi tu choi truoc khi build.
  const unicode = upstream + '\n\' Ghi chu co dau: chuẩn hóa\n';
  assert.throws(() => personalize(unicode), /ngoai ASCII/);
});

console.log('');
if (fail > 0) {
  console.log(`THAT BAI: ${fail} loi, ${pass} dat`);
  process.exit(1);
}
console.log(`OK: ${pass}/${pass} kiem tra dat`);