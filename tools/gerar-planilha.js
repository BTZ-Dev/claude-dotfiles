/**
 * Gera planilha Yampi (.xlsx) a partir de um arquivo JSON de produtos.
 * Não requer nenhum pacote npm — usa Node.js puro + PowerShell para criar o ZIP.
 * Uso: node gerar-planilha.js <caminho_do_json>
 */

const fs   = require('fs');
const path = require('path');
const os   = require('os');
const { execSync } = require('child_process');

const HEADERS = [
  'id','ativo','possui_variacoes','marca','codigo_erp','ncm','nome',
  'buscavel','produto_digital','categorias','colecoes','filtros',
  'variacoes','selos','slug','video','descricao','meses_de_garantia',
  'frete_customizado','valor_do_frete','especificacoes','medidas',
  'valor_de_presente','categoria_google','seo_titulo_pagina','seo_descricao',
  'seo_palavras_chave','link_canonico','termos_de_busca','link_produto',
  'link_foto_principal'
];

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

function buildXlsx(produtos, outputPath) {
  // ── Shared strings index ──────────────────────────────────────
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

  // ── sheet1.xml ────────────────────────────────────────────────
  let rows = '';
  // header row with bold+brown style (s="1")
  rows += `<row r="1">` +
    HEADERS.map((h, ci) => `<c r="${colLetter(ci+1)}1" t="s" s="1"><v>${si(h)}</v></c>`).join('') +
    `</row>`;
  // data rows
  produtos.forEach((p, ri) => {
    const rn = ri + 2;
    rows += `<row r="${rn}">` +
      HEADERS.map((h, ci) => `<c r="${colLetter(ci+1)}${rn}" t="s"><v>${si(p[h]||'')}</v></c>`).join('') +
      `</row>`;
  });

  const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>${rows}</sheetData>
</worksheet>`;

  // ── sharedStrings.xml ─────────────────────────────────────────
  const sharedXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${strings.length}" uniqueCount="${strings.length}">
${strings.map(s => `<si><t xml:space="preserve">${esc(s)}</t></si>`).join('\n')}
</sst>`;

  // ── styles.xml (normal + cabeçalho marrom) ────────────────────
  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
</styleSheet>`;

  // ── workbook.xml ──────────────────────────────────────────────
  const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Produtos" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;

  // ── [Content_Types].xml ───────────────────────────────────────
  const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`;

  const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

  const wbRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

  // ── Escreve arquivos em diretório temporário ──────────────────
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'xlsx-'));
  const mk  = (...p) => { fs.mkdirSync(path.join(tmp, ...p), { recursive: true }); };
  const wr  = (p, content) => fs.writeFileSync(path.join(tmp, p), content, 'utf8');

  mk('_rels');
  mk('xl', '_rels');
  mk('xl', 'worksheets');

  wr('[Content_Types].xml',           contentTypesXml);
  wr('_rels/.rels',                   rootRels);
  wr('xl/workbook.xml',               workbookXml);
  wr('xl/_rels/workbook.xml.rels',    wbRels);
  wr('xl/worksheets/sheet1.xml',      sheetXml);
  wr('xl/sharedStrings.xml',          sharedXml);
  wr('xl/styles.xml',                 stylesXml);

  // ── Compacta em .xlsx via PowerShell (sem npm) ────────────────
  const safeOut = outputPath.replace(/'/g, "''");
  const safeTmp = tmp.replace(/'/g, "''");

  const ps = `
Add-Type -AssemblyName System.IO.Compression.FileSystem
$out = '${safeOut}'
$src = '${safeTmp}'
if (Test-Path $out) { Remove-Item -Force $out }
$zip = [System.IO.Compression.ZipFile]::Open($out, [System.IO.Compression.ZipArchiveMode]::Create)
Get-ChildItem -Path $src -Recurse -File | ForEach-Object {
  $entry = $_.FullName.Substring($src.Length).TrimStart('\\').Replace('\\','/')
  [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $_.FullName, $entry) | Out-Null
}
$zip.Dispose()
Write-Host "OK"
`.trim();

  execSync(`powershell -NoProfile -Command "${ps.replace(/"/g, '\\"')}"`, { stdio: 'pipe' });

  // Limpa temp
  fs.rmSync(tmp, { recursive: true, force: true });
  return outputPath;
}

// ── Main ──────────────────────────────────────────────────────────
const args = process.argv.slice(2);
if (!args[0]) {
  console.error('Uso: node gerar-planilha.js <arquivo_produtos.json>');
  process.exit(1);
}

const jsonPath = path.resolve(args[0]);
if (!fs.existsSync(jsonPath)) {
  console.error('Arquivo não encontrado: ' + jsonPath);
  process.exit(1);
}

const produtos = JSON.parse(fs.readFileSync(jsonPath, 'utf8'));
if (!Array.isArray(produtos)) {
  console.error('O JSON deve ser uma lista de produtos.');
  process.exit(1);
}

const pasta  = path.dirname(path.resolve(__filename));
const hoje   = new Date().toISOString().slice(0, 10);
const saida  = path.join(pasta, `yampi-randa-mundu-${hoje}.xlsx`);

// Copia imagens locais referenciadas em link_foto_principal para pasta de saída
const pastaImagens = path.join(pasta, `yampi-imagens-${hoje}`);
let imagensCopiaadas = 0;
produtos.forEach(p => {
  const imgPath = p.link_foto_principal;
  if (imgPath && fs.existsSync(imgPath)) {
    if (!fs.existsSync(pastaImagens)) fs.mkdirSync(pastaImagens, { recursive: true });
    const dest = path.join(pastaImagens, path.basename(imgPath));
    fs.copyFileSync(imgPath, dest);
    imagensCopiaadas++;
  }
});

try {
  buildXlsx(produtos, saida);
  console.log('SUCESSO: ' + saida);
  if (imagensCopiaadas > 0) console.log(`IMAGENS: ${imagensCopiaadas} arquivo(s) copiado(s) para ${pastaImagens}`);
} catch (err) {
  console.error('ERRO: ' + err.message);
  process.exit(1);
}
