// Test hoi quy cho vba2bas.mjs: bao dam khong bao gio sinh .bas hong compile.
// Chay: node vba2bas.test.mjs
import { convertText, validateBas } from './vba2bas.mjs';

let pass = 0, fail = 0;
function ok(name, cond, extra) {
  if (cond) { pass++; console.log('  dat  ' + name); }
  else { fail++; console.log('  HONG ' + name + (extra ? '  -> ' + extra : '')); }
}

// 1. Dau nhay kep escape "" phai sinh ra .bas hop le
const src1 = [
  'Public Sub T()',
  '    MsgBox "Chỗ viết ""bảng/hình dưới đây"" phải nêu số"',
  'End Sub',
  '',
].join('\r\n');
const out1 = convertText(src1);
ok('escape "" sinh .bas hop le', validateBas(out1).length === 0,
   JSON.stringify(validateBas(out1)[0] || ''));
// Bo cac cum ChrW(...) di, chi giu phan chu ASCII, de kiem tra dau nhay kep escape
const plain1 = out1.replace(/ChrW\([^)]*\)/g, '').replace(/&\s*&\s*/g, '&').replace(/\s+/g, ' ');
ok('escape "" giu du 2 cap dau nhay kep', (plain1.match(/""/g) || []).length >= 2,
   'thay ' + (plain1.match(/""/g) || []).length + ' cap: ' + plain1.split('MsgBox')[1]);
ok('khong con ky tu >127', /^[\x00-\x7F]*$/.test(out1));

// 2. Dong lech dau nhay kep (so dau nhay le) phai bi bat — day chinh la loi tung gap
const bad = 'x = "mot hai' + '\r\n';
ok('bat duoc dong lech nhay kep', validateBas(bad).length > 0, 'KHONG BAT DUOC');
ok('khong bao dong sach', validateBas('x = "sach"' + '\r\n').length === 0, 'bao sai');

// 3. Tieng Viet co dau -> ChrW, doc lai dung
const src3 = 'ten = "Đồ án tốt nghiệp"' + '\r\n';
const out3 = convertText(src3);
ok('tieng Viet qua ChrW', /ChrW\(/.test(out3) && validateBas(out3).length === 0, out3.trim());

// 4. Comment bo dau, khong con ChrW trong comment
const src4 = "' Đồ án tốt nghiệp — ghi chú" + '\r\nz = 1\r\n';
const out4 = convertText(src4);
ok('comment bo dau', /^' Do an tot nghiep/m.test(out4), out4.split('\r\n')[0]);

console.log('\n' + pass + ' dat, ' + fail + ' hong');
process.exit(fail ? 1 : 0);