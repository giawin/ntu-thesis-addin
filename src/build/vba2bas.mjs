// Chuyen file .vba (UTF-8, tieng Viet co dau) thanh .bas thuan ASCII:
// chuoi trong dau nhay kep -> "..." & ChrW(&Hxxxx) & "..."
// comment co dau -> bo dau (vi khong anh huong chay)
// Muc dich: VBA luu code theo codepage ANSI nen ghi thang tieng Viet se mat dau.
import fs from 'node:fs';
import path from 'node:path';

const FOLD = {
  à:'a',á:'a',ả:'a',ã:'a',ạ:'a',ă:'a',ằ:'a',ắ:'a',ẳ:'a',ẵ:'a',ặ:'a',â:'a',ầ:'a',ấ:'a',ẩ:'a',ẫ:'a',ậ:'a',
  è:'e',é:'e',ẻ:'e',ẽ:'e',ẹ:'e',ê:'e',ề:'e',ế:'e',ể:'e',ễ:'e',ệ:'e',
  ì:'i',í:'i',ỉ:'i',ĩ:'i',ị:'i',
  ò:'o',ó:'o',ỏ:'o',õ:'o',ọ:'o',ô:'o',ồ:'o',ố:'o',ổ:'o',ỗ:'o',ộ:'o',ơ:'o',ờ:'o',ớ:'o',ở:'o',ỡ:'o',ợ:'o',
  ù:'u',ú:'u',ủ:'u',ũ:'u',ụ:'u',ư:'u',ừ:'u',ứ:'u',ử:'u',ữ:'u',ự:'u',
  ỳ:'y',ý:'y',ỷ:'y',ỹ:'y',ỵ:'y',đ:'d',
  À:'A',Á:'A',Ả:'A',Ã:'A',Ạ:'A',Ă:'A',Ằ:'A',Ắ:'A',Ẳ:'A',Ẵ:'A',Ặ:'A',Â:'A',Ầ:'A',Ấ:'A',Ẩ:'A',Ẫ:'A',Ậ:'A',
  È:'E',É:'E',Ẻ:'E',Ẽ:'E',Ẹ:'E',Ê:'E',Ề:'E',Ế:'E',Ể:'E',Ễ:'E',Ệ:'E',
  Ì:'I',Í:'I',Ỉ:'I',Ĩ:'I',Ị:'I',
  Ò:'O',Ó:'O',Ỏ:'O',Õ:'O',Ọ:'O',Ô:'O',Ồ:'O',Ố:'O',Ổ:'O',Ỗ:'O',Ộ:'O',Ơ:'O',Ờ:'O',Ớ:'O',Ở:'O',Ỡ:'O',Ợ:'O',
  Ù:'U',Ú:'U',Ủ:'U',Ũ:'U',Ụ:'U',Ư:'U',Ừ:'U',Ứ:'U',Ử:'U',Ữ:'U',Ự:'U',
  Ỳ:'Y',Ý:'Y',Ỷ:'Y',Ỹ:'Y',Ỵ:'Y',Đ:'D',
};

const fold = (s) => s.replace(/[^\x00-\x7F]/g, (c) => FOLD[c] ?? '');

// Tach mot dong thanh cac doan: comment / chuoi / code
export function convertLine(line) {
  let out = '';
  let i = 0;
  while (i < line.length) {
    const c = line[i];
    if (c === '"') {
      // doc het chuoi (xu ly "" la escape)
      let j = i + 1;
      let body = '';
      while (j < line.length) {
        if (line[j] === '"') {
          if (line[j + 1] === '"') { body += '"'; j += 2; continue; }
          break;
        }
        body += line[j];
        j++;
      }
      out += '"' + escapeBody(body) + '"';
      i = j + 1;
      continue;
    }
    if (c === "'") {
      out += fold(line.slice(i));
      break;
    }
    out += c;
    i++;
  }
  return out;
}

function escapeBody(body) {
  // Dau nhay kep ben trong chuoi VBA phai viet la "" — neu khong se pha vo literal
  // (chunk() chen cac doan " & ChrW(...) & " nen phai escape TRUOC khi gom cum)
  body = body.replace(/"/g, '""');
  if (!/[^\x00-\x7F]/.test(body)) return body;
  // Gom cum ky tu khong-ASCII lien tiep thanh mot ChrW
  let out = '';
  let buf = '';
  for (const ch of body) {
    if (ch.codePointAt(0) > 0x7f) { buf += ch; continue; }
    if (buf) { out += chunk(buf); buf = ''; }
    out += ch;
  }
  if (buf) out += chunk(buf);
  return out;
}

function chunk(s) {
  // "abc" + ChrW(&H1ED3) + "xyz" — dau va cuoi noi bang & de nam trong literal
  return '" & ' + [...s].map((c) => `ChrW(&H${c.codePointAt(0).toString(16).toUpperCase()})`).join(' & ') + ' & "';
}

// Kiem tra moi dong co so dau nhay kep CHAN (ngoai comment) — lech chan = literal bi pha vo
export function validateBas(text) {
  const errs = [];
  text.split(/\r?\n/).forEach((line, idx) => {
    let quotes = 0;
    for (let i = 0; i < line.length; i++) {
      if (line[i] === "'") break;      // comment: phan con lai khong tinh
      if (line[i] === '"') quotes++;
    }
    if (quotes % 2 !== 0) errs.push(`dong ${idx + 1}: lech so dau nhay kep (${quotes}) -> ${line.slice(0, 90)}`);
  });
  return errs;
}

export function convertText(src) {
  return src.split(/\r?\n/).map(convertLine).join('\r\n');
}

if (process.argv[1] && import.meta.url.endsWith(path.basename(process.argv[1]))) {
  const [inp, outp] = process.argv.slice(2);
  if (!inp || !outp) {
    console.error('usage: node vba2bas.mjs <in.vba> <out.bas>');
    process.exit(2);
  }
  const out = convertText(fs.readFileSync(inp, 'utf8'));
  const bad = [...out].filter((c) => c.charCodeAt(0) > 0x7f);
  if (bad.length) {
    console.error(`CON: ${bad.length} ky tu khong phai ASCII sau khi chuyen`);
    process.exit(1);
  }
  const errs = validateBas(out);
  if (errs.length) {
    console.error('CON: ' + errs.length + ' dong lech dau nhay kep (se gay loi bien dich VBA):');
    errs.slice(0, 10).forEach((e) => console.error('  ' + e));
    process.exit(1);
  }
  fs.writeFileSync(outp, out, 'latin1');
  console.log(`OK -> ${outp} (${out.length} bytes, ASCII thuan)`);
}