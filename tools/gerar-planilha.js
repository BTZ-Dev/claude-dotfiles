/**
 * Gera planilha Yampi (.xlsx) a partir de um arquivo JSON de produtos.
 * 100% Node.js puro — sem npm, sem PowerShell, sem dependências externas.
 * Uso: node gerar-planilha.js <caminho_do_json>
 */

const fs   = require('fs');
const path = require('path');

const HEADERS = [
  'id','ativo','possui_variacoes','marca','codigo_erp','ncm','nome',
  'buscavel','produto_digital','categorias','colecoes','filtros',
  'variacoes','selos','slug','video','descricao','meses_de_garantia',
  'frete_customizado','valor_do_frete','especificacoes','medidas',
  'valor_de_presente','categoria_google','seo_titulo_pagina','seo_descricao',
  'seo_palavras_chave','link_canonico','termos_de_busca','link_produto',
  'link_foto_principal'
];

// ── CRC32 ─────────────────────────────────────────────────────────
const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let j = 0; j < 8; j++) c = (c & 1) ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
    t[i] = c;
  }
  return t;
})();

function crc32(buf) {
  let c = 0xFFFFFFFF;
  for (const b of buf) c = CRC_TABLE[(c ^ b) & 0xFF] ^ (c >>> 8);
  return (c ^ 0xFFFFFFFF) >>> 0;
}

// ── ZIP writer (STORED, sem compressão) ───────────────────────────
function writeZip(entries, outputPath) {
  const parts = [];
  const central = [];
  let offset = 0;

  for (const [name, content] of Object.entries(entries)) {
    const nameBuf = Buffer.from(name, 'utf8');
    const data    = Buffer.isBuffer(content) ? content : Buffer.from(content, 'utf8');
    const crc     = crc32(data);

    // Local file header (30 bytes + name)
    const lh = Buffer.alloc(30 + nameBuf.length);
    lh.writeUInt32LE(0x04034B50, 0);   // signature
    lh.writeUInt16LE(20,         4);   // version needed
    lh.writeUInt16LE(0,          6);   // flags
    lh.writeUInt16LE(0,          8);   // STORED
    lh.writeUInt16LE(0,          10);  // mod time
    lh.writeUInt16LE(0,          12);  // mod date
    lh.writeUInt32LE(crc,        14);
    lh.writeUInt32LE(data.length,18);  // compressed size
    lh.writeUInt32LE(data.length,22);  // uncompressed size
    lh.writeUInt16LE(nameBuf.length,26);
    lh.writeUInt16LE(0,          28);  // extra length
    nameBuf.copy(lh, 30);

    parts.push(lh, data);

    // Central directory entry (46 bytes + name)
    const cd = Buffer.alloc(46 + nameBuf.length);
    cd.writeUInt32LE(0x02014B50, 0);
    cd.writeUInt16LE(20,         4);
    cd.writeUInt16LE(20,         6);
    cd.writeUInt16LE(0,          8);
    cd.writeUInt16LE(0,          10);
    cd.writeUInt16LE(0,          12);
    cd.writeUInt16LE(0,          14);
    cd.writeUInt32LE(crc,        16);
    cd.writeUInt32LE(data.length,20);
    cd.writeUInt32LE(data.length,24);
    cd.writeUInt16LE(nameBuf.length,28);
    cd.writeUInt16LE(0,          30);  // extra
    cd.writeUInt16LE(0,          32);  // comment
    cd.writeUInt16LE(0,          34);  // disk start
    cd.writeUInt16LE(0,          36);  // internal attr
    cd.writeUInt32LE(0,          38);  // external attr
    cd.writeUInt32LE(offset,     42);  // local header offset
    nameBuf.copy(cd, 46);

    central.push(cd);
    offset += lh.length + data.length;
  }

  const cdBuf = Buffer.concat(central);
  const count  = central.length;

  // End of central directory (22 bytes)
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054B50, 0);
  eocd.writeUInt16LE(0,          4);
  eocd.writeUInt16LE(0,          6);
  eocd.writeUInt16LE(count,      8);
  eocd.writeUInt16LE(count,      10);
  eocd.writeUInt32LE(cdBuf.length, 12);
  eocd.writeUInt32LE(offset,     16);
  eocd.writeUInt16LE(0,          20);

  fs.writeFileSync(outputPath, Buffer.concat([...parts, cdBuf, eocd]));
}

// ── XML helpers ───────────────────────────────────────────────────
function esc(s) {
  return String(s == null ? '' : s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&apos;');
}

function colLetter(n) {
  let s = '';
  while (n > 0) { s = String.fromCharCode(64 + (n - 1) % 26 + 1) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// ── Gera o xlsx ───────────────────────────────────────────────────
function buildXlsx(produtos, outputPath) {
  // Shared strings
  const strings = [];
  const idx = {};
  function si(v) {
    const s = String(v == null ? '' : v);
    if (s in idx) return idx[s];
    idx[s] = strings.length;
    strings.push(s);
    return strings.length - 1;
  }
  HEADERS.forEach(h => si(h));
  produtos.forEach(p => HEADERS.forEach(h => si(p[h] || '')));

  // sheet1.xml
  let rows = `<row r="1">` +
    HEADERS.map((h, ci) => `<c r="${colLetter(ci+1)}1" t="s" s="1"><v>${si(h)}</v></c>`).join('') +
    `</row>`;
  produtos.forEach((p, ri) => {
    const rn = ri + 2;
    rows += `<row r="${rn}">` +
      HEADERS.map((h, ci) => `<c r="${colLetter(ci+1)}${rn}" t="s"><v>${si(p[h]||'')}</v></c>`).join('') +
      `</row>`;
  });

  const files = {
    '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`,

    '_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,

    'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Produtos" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,

    'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`,

    'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>${rows}</sheetData>
</worksheet>`,

    'xl/sharedStrings.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${strings.length}" uniqueCount="${strings.length}">
${strings.map(s => `<si><t xml:space="preserve">${esc(s)}</t></si>`).join('\n')}
</sst>`,

    'xl/styles.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/><color rgb="FFFFFFFF"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF6B4F3A"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
  </cellXfs>
</styleSheet>`,
  };

  writeZip(files, outputPath);
}

// ── Main ──────────────────────────────────────────────────────────
const args = process.argv.slice(2);
if (!args[0]) { console.error('Uso: node gerar-planilha.js <arquivo_produtos.json>'); process.exit(1); }

const jsonPath = path.resolve(args[0]);
if (!fs.existsSync(jsonPath)) { console.error('Arquivo não encontrado: ' + jsonPath); process.exit(1); }

const produtos = JSON.parse(fs.readFileSync(jsonPath, 'utf8'));
if (!Array.isArray(produtos)) { console.error('O JSON deve ser uma lista de produtos.'); process.exit(1); }

const pasta = path.dirname(path.resolve(__filename));
const hoje  = new Date().toISOString().slice(0, 10);
const saida = path.join(pasta, `yampi-randa-mundu-${hoje}.xlsx`);

// Copia todas as imagens do produto (STILLs + LOOKBOOK) para pasta de saída
const pastaImagens = path.join(pasta, `yampi-imagens-${hoje}`);
let imagensCopiaadas = 0;

const indexPath = path.join(pasta, 'drive-index.json');
const driveIndex = fs.existsSync(indexPath)
  ? JSON.parse(fs.readFileSync(indexPath, 'utf8').replace(/^﻿/, ''))
  : {};

produtos.forEach(p => {
  const codigo = String(p.codigo_erp || '');
  const entry  = driveIndex[codigo] || {};
  const todas  = [...(entry.stills || []), ...(entry.lookbook || [])];

  // Se não há índice, tenta pelo menos a foto principal
  if (todas.length === 0 && p.link_foto_principal) todas.push(p.link_foto_principal);

  todas.forEach(imgPath => {
    if (imgPath && fs.existsSync(imgPath)) {
      if (!fs.existsSync(pastaImagens)) fs.mkdirSync(pastaImagens, { recursive: true });
      fs.copyFileSync(imgPath, path.join(pastaImagens, path.basename(imgPath)));
      imagensCopiaadas++;
    }
  });
});

try {
  buildXlsx(produtos, saida);
  console.log('SUCESSO: ' + saida);
  if (imagensCopiaadas > 0) console.log(`IMAGENS: ${imagensCopiaadas} arquivo(s) copiado(s) para ${pastaImagens}`);
} catch (err) {
  console.error('ERRO: ' + err.message);
  process.exit(1);
}
