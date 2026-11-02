'use strict';
const JSZip = require('jszip');
const sax = require('sax');
const { createHash } = require('node:crypto');
const { inflateRawSync } = require('node:zlib');
const { requireCondition: check } = require('./errors.cjs');
const defaults = require('../config/development.json').limits;
const templatePolicy = require('../config/template-policy.json');
const {validateOperation,validateOperationResult}=require('./operation.cjs');
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const editablePart = /^word\/(document|header\d+|footer\d+)\.xml$/;
const sha256 = bytes => createHash('sha256').update(bytes).digest('hex');

// Le o diretorio central ANTES de descompactar. ZIP64, criptografia e nomes
// duplicados/ambiguos sao recusados neste subconjunto conservador de OOXML.
function checkZip(bytes, limits = defaults) {
  check(Buffer.isBuffer(bytes) && bytes.length <= limits.maxFileBytes, 'FILE_SIZE_LIMIT');
  let end = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
    if (bytes.readUInt32LE(i) === 0x06054b50 && i + 22 + bytes.readUInt16LE(i + 20) === bytes.length) { end = i; break; }
  }
  check(end >= 0, 'INVALID_ZIP');
  const count = bytes.readUInt16LE(end + 10);
  const size = bytes.readUInt32LE(end + 12), offset = bytes.readUInt32LE(end + 16);
  check(bytes.readUInt16LE(end + 4) === 0 && bytes.readUInt16LE(end + 6) === 0 && bytes.readUInt16LE(end + 8) === count, 'MULTIVOLUME_ZIP');
  check(count > 0 && count <= limits.maxEntries && count !== 65535 && size !== 0xffffffff && offset !== 0xffffffff, 'ZIP_LIMIT');
  check(offset + size === end, 'ZIP_DIRECTORY');
  let pos = offset, expanded = 0;
  const names = new Set();
  for (let i = 0; i < count; i++) {
    check(pos + 46 <= end && bytes.readUInt32LE(pos) === 0x02014b50, 'ZIP_DIRECTORY');
    const flags = bytes.readUInt16LE(pos + 8), method = bytes.readUInt16LE(pos + 10);
    const packed = bytes.readUInt32LE(pos + 20), unpacked = bytes.readUInt32LE(pos + 24);
    const n = bytes.readUInt16LE(pos + 28), extra = bytes.readUInt16LE(pos + 30), comment = bytes.readUInt16LE(pos + 32);
    check(pos + 46 + n + extra + comment <= end, 'ZIP_DIRECTORY');
    const name = bytes.subarray(pos + 46, pos + 46 + n).toString('utf8');
    check(/^[\x20-\x7e]+$/.test(name) && !name.startsWith('/') && !name.includes('\\') && !name.includes(':') && !name.split('/').some(x => x === '..' || x === '.'), 'ZIP_UNSAFE_PATH');
    check(!names.has(name.toLowerCase()), 'ZIP_DUPLICATE_ENTRY'); names.add(name.toLowerCase());
    check(!(flags & 1) && [0, 8].includes(method), 'ZIP_UNSUPPORTED');
    expanded += unpacked;
    check(unpacked <= limits.maxPartBytes && expanded <= limits.maxExpandedBytes, 'ZIP_EXPANSION_LIMIT');
    const local = bytes.readUInt32LE(pos + 42);
    check(local + 30 <= offset && bytes.readUInt32LE(local) === 0x04034b50, 'ZIP_LOCAL_HEADER');
    const localNameLength = bytes.readUInt16LE(local + 26), localExtraLength = bytes.readUInt16LE(local + 28);
    check(bytes.subarray(local + 30, local + 30 + localNameLength).toString('utf8') === name, 'ZIP_NAME_MISMATCH');
    check(bytes.readUInt16LE(local + 6) === flags && bytes.readUInt16LE(local + 8) === method && local + 30 + localNameLength + localExtraLength + packed <= offset, 'ZIP_LOCAL_HEADER');
    const payloadStart = local + 30 + localNameLength + localExtraLength;
    const payload = bytes.subarray(payloadStart, payloadStart + packed);
    let inflated;
    try { inflated = method === 0 ? payload : inflateRawSync(payload, { maxOutputLength: Math.max(1, unpacked) }); }
    catch { check(false, 'ZIP_EXPANSION_OR_DATA_ERROR'); }
    check(inflated.length === unpacked, 'ZIP_SIZE_MISMATCH');
    pos += 46 + n + extra + comment;
  }
  check(pos === end, 'ZIP_DIRECTORY');
  return { entries: count, expandedBytes: expanded };
}

function parseXml(xml) {
  check(!/<!DOCTYPE|<!ENTITY|<!\[CDATA\[/i.test(xml), 'XML_UNSUPPORTED_DECLARATION');
  const parser = sax.parser(true, { xmlns: true, position: true, strictEntities: true });
  const stack = [], nodes = [], paragraphs = [];
  parser.onerror = () => { throw new (require('./errors.cjs').Blocked)('INVALID_XML'); };
  parser.onopentag = tag => {
    const parent = stack.at(-1);
    const node = { name: tag.name, local: tag.local, uri: tag.uri, attrs: tag.attributes,
      start: parser.startTagPosition - 1, openEnd: parser.position, selfClosing: tag.isSelfClosing,
      parent, children: [], text: '' };
    if (parent) parent.children.push(node);
    stack.push(node); nodes.push(node);
    if (tag.uri === W && tag.local === 'p') paragraphs.push(node);
  };
  parser.ontext = text => { if (stack.length) stack.at(-1).text += text; };
  parser.onclosetag = () => {
    const node = stack.pop(); node.end = parser.position;
    node.contentEnd = node.selfClosing ? node.openEnd : xml.lastIndexOf('</', parser.position - 1);
  };
  parser.write(xml).close();
  check(stack.length === 0 && nodes.length > 0, 'INVALID_XML');
  return { nodes, paragraphs };
}

function paragraphsOf(xml) {
  const parsed = parseXml(xml);
  const fields=[],stack=[];
  const readOnlyContainers=parsed.nodes.filter(n=>n.uri===W&&['sdt','fldSimple','ins','del','moveFrom','moveTo'].includes(n.local));
  for(const n of parsed.nodes.filter(n=>n.uri===W&&n.local==='fldChar')){const type=Object.values(n.attrs).find(a=>a.local==='fldCharType')?.value;if(type==='begin')stack.push(n.start);if(type==='end'&&stack.length)fields.push({start:stack.pop(),end:n.end});}
  for(const start of stack)fields.push({start,end:xml.length});
  return parsed.paragraphs.map((p, index) => {
    const texts = parsed.nodes.filter(n => n.uri === W && n.local === 't' && ancestorParagraph(n) === p);
    let at = 0;
    for (const t of texts) { t.from = at; at += t.text.length; t.to = at; }
    // Nenhuma substituicao atravessa campos, quebras, hiperlinks ou text boxes.
    const unsafe = parsed.nodes.some(n => ancestorParagraph(n) === p && n.uri === W &&
      ['tab','br','cr','fldChar','instrText','drawing','pict','object','sdt','hyperlink','del','ins','moveFrom','moveTo','vanish'].includes(n.local));
    const readOnly=texts.filter(t=>fields.some(f=>f.start<t.start&&f.end>t.end)||readOnlyContainers.some(n=>n.start<t.start&&n.end>t.end)).map(t=>({start:t.from,end:t.to}));
    const boundaries=parsed.nodes.filter(n=>ancestorParagraph(n)===p&&n.uri===W&&['tab','br','cr','drawing','pict','object'].includes(n.local)).map(n=>texts.filter(t=>t.start<n.start).reduce((sum,t)=>sum+t.text.length,0));
    return { index, texts, text: texts.map(n => n.text).join(''), unsafe,readOnly,boundaries };
  });
}
function textRangeSafe(p,start,end){return !!p&&!p.readOnly.some(r=>r.start<end&&r.end>start)&&!p.boundaries.some(at=>start<at&&at<end);}
function ancestorParagraph(n) {
  for (let p = n.parent; p; p = p.parent) if (p.uri === W && p.local === 'p') return p;
  return null;
}

async function openDocx(bytes, limits = defaults) {
  const stats = checkZip(bytes, limits);
  let zip;
  try { zip = await JSZip.loadAsync(bytes, { checkCRC32: true, createFolders: false }); }
  catch { check(false, 'INVALID_ZIP_DATA'); }
  check(zip.file('[Content_Types].xml') && zip.file('_rels/.rels') && zip.file('word/document.xml'), 'NOT_DOCX');
  const parts = new Map();
  for (const [name, entry] of Object.entries(zip.files)) {
    if (entry.dir) continue;
    const data = await entry.async('nodebuffer');
    check(data.length <= limits.maxPartBytes, 'ZIP_EXPANSION_LIMIT');
    parts.set(name, data);
    if (/\.(xml|rels)$/.test(name)) parseXml(data.toString('utf8'));
  }
  check(parts.get('[Content_Types].xml').toString().includes('application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'), 'NOT_STANDARD_DOCX');
  return { zip, parts, stats, hash: sha256(bytes) };
}

function privacyFindings(doc) {
  const findings = new Set();
  for (const [name, bytes] of doc.parts) {
    if (/comments|people|customXml|docProps\/custom/i.test(name)) findings.add('COMMENTS_OR_CUSTOM_DATA');
    if (/embeddings\/|vbaProject|activeX|_xmlsignatures/i.test(name)) findings.add('ACTIVE_OR_EMBEDDED_CONTENT');
    if (/media\//.test(name)) findings.add('MEDIA_REQUIRES_LOCAL_REVIEW');
    if (!/\.(xml|rels)$/.test(name)) continue;
    const xml = bytes.toString('utf8');
    const { nodes } = parseXml(xml);
    if (nodes.some(n => n.uri === W && ['ins','del','moveFrom','moveTo','delText'].includes(n.local))) findings.add('TRACKED_CHANGES');
    if (nodes.some(n => n.uri === W && ['vanish','webHidden'].includes(n.local))) findings.add('HIDDEN_TEXT');
    if (nodes.some(n => Object.values(n.attrs).some(a => a.local === 'TargetMode' && a.value === 'External'))) findings.add('EXTERNAL_RELATIONSHIP');
    if (name === 'docProps/core.xml' && nodes.some(n => ['creator','lastModifiedBy','title','subject','description','keywords'].includes(n.local) && n.text.trim())) findings.add('IDENTIFYING_METADATA');
  }
  return [...findings].sort();
}

function escapeText(value) { return value.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }

async function applyAuthorized(bytes, profile, plan) {
  check(templatePolicy.mode==='literal-template'&&templatePolicy.useExactSelectedTemplate===true&&templatePolicy.preserveOriginalLogo===true&&templatePolicy.recreateDocument===false&&templatePolicy.rewriteClauses===false&&templatePolicy.createOrReplaceImages===false,'LITERAL_TEMPLATE_POLICY_REQUIRED');
  const doc = await openDocx(bytes);
  check(profile.scope === 'synthetic-only' && plan.scope === 'synthetic-only', 'SYNTHETIC_ONLY');
  check(doc.hash === profile.templateSha256 && doc.hash === plan.templateSha256, 'TEMPLATE_HASH_MISMATCH');
  check(profile.anonymizationApproved !== false, 'ANONYMIZATION_APPROVAL_REQUIRED');
  const media=[...doc.parts].filter(([part])=>part.startsWith('word/media/'));
  const approvedMedia=profile.approvedMedia;
  const mediaApproved=profile.anonymizationApproved===true&&Array.isArray(approvedMedia)&&approvedMedia.length===media.length&&
    new Set(approvedMedia.map(m=>m.part)).size===media.length&&media.every(([part,bytes])=>approvedMedia.some(m=>m.part===part&&m.sha256===sha256(bytes)));
  for(const [part,data] of doc.parts){
    const findings=privacyFindings({parts:new Map([[part,data]])});
    const reviewed=profile.anonymizationApproved===true&&(profile.reviewedPreservedParts||[]).some(r=>r.part===part&&r.sha256===sha256(data)&&r.reviewed===true);
    check(findings.filter(f=>!(f==='MEDIA_REQUIRES_LOCAL_REVIEW'&&mediaApproved)&&!(reviewed&&['COMMENTS_OR_CUSTOM_DATA','EXTERNAL_RELATIONSHIP'].includes(f))).length===0,'DOCUMENT_REQUIRES_PRIVACY_REVIEW');
  }
  check(Array.isArray(plan.pending) && plan.pending.length === 0, 'PENDING_INFORMATION');
  check(Array.isArray(plan.replacements) && Array.isArray(profile.fields), 'INVALID_PLAN');
  const fields = new Map(profile.fields.map(f => [f.id, f]));
  check(fields.size === profile.fields.length && fields.size > 0, 'INVALID_PROFILE');
  const replacements = new Map();
  for (const r of plan.replacements) {
    check(fields.has(r.fieldId) && !replacements.has(r.fieldId), 'UNKNOWN_OR_DUPLICATE_FIELD');
    check(typeof r.value === 'string' && r.value.trim().length > 0 && r.value.length <= 2000 && !/[\u0000-\u001f\u007f]/u.test(r.value) && !/[\uD800-\uDFFF]/u.test(r.value.replace(/[\uD800-\uDBFF][\uDC00-\uDFFF]/g, '')), 'INVALID_REPLACEMENT');
    check(r.source && typeof r.source.documentId === 'string' && /^[A-Za-z0-9_-]{1,80}$/.test(r.source.documentId) && Number.isInteger(r.source.page) && r.source.page > 0, 'SOURCE_REQUIRED');
    replacements.set(r.fieldId, r);
  }
  for (const f of fields.values()) check(!f.required || replacements.has(f.id), 'REQUIRED_FIELD_MISSING');
  const updates = new Map(), audit = [];
  for (const [id, r] of replacements) {
    const f = fields.get(id);
    check(editablePart.test(f.part) && doc.parts.has(f.part), 'PART_NOT_ALLOWED');
    const xml = doc.parts.get(f.part).toString('utf8');
    const p = paragraphsOf(xml)[f.paragraph];
    check(p && !p.unsafe, 'UNSAFE_PARAGRAPH');
    check(Number.isInteger(f.start) && Number.isInteger(f.end) && f.start >= 0 && f.end > f.start && p.text.slice(f.start, f.end) === f.oldText && f.end <= p.text.length, 'ANCHOR_MISMATCH');
    if(p.text.indexOf(f.oldText)!==p.text.lastIndexOf(f.oldText)){
      const context=f.anchorContext;
      check(context&&typeof context.before==='string'&&typeof context.after==='string'&&context.before.length+context.after.length>=16,'AMBIGUOUS_ANCHOR');
      const needle=context.before+f.oldText+context.after;
      check(p.text.indexOf(needle)===f.start-context.before.length&&p.text.lastIndexOf(needle)===p.text.indexOf(needle),'AMBIGUOUS_ANCHOR');
    }
    const same = audit.filter(a => a.part === f.part && a.paragraph === f.paragraph);
    check(!same.some(a => f.start < a.end && f.end > a.start), 'OVERLAPPING_REPLACEMENTS');
    const involved = p.texts.filter(t => t.from < f.end && t.to > f.start);
    check(involved.length > 0 && involved.every(t => !t.selfClosing && t.parent?.uri === W && t.parent.local === 'r' && t.parent.parent?.local === 'p'), 'UNSUPPORTED_TEXT_LOCATION');
    let remaining = r.value;
    for (let i = 0; i < involved.length; i++) {
      const t = involved[i], start = Math.max(0, f.start - t.from), end = Math.min(t.text.length, f.end - t.from);
      const piece = i === involved.length - 1 ? remaining : remaining.slice(0, end - start);
      remaining = remaining.slice(piece.length);
      const key = `${f.part}:${t.start}`;
      if (!updates.has(key)) updates.set(key, { part: f.part, node: t, edits: [] });
      updates.get(key).edits.push({ start, end, value: piece });
    }
    audit.push({ fieldId: id, part: f.part, paragraph: f.paragraph, start: f.start, end: f.end,
      sourceDocumentId: r.source.documentId, sourcePage: r.source.page });
  }
  const operation=validateOperation(doc,profile,plan,paragraphsOf);
  const byPart = new Map();
  for (const update of updates.values()) {
    let value = update.node.text;
    for (const e of update.edits.sort((a,b) => b.start-a.start)) value = value.slice(0,e.start) + e.value + value.slice(e.end);
    const preserve = Object.values(update.node.attrs).some(a => a.name === 'xml:space' && a.value === 'preserve');
    check(preserve || value.trim() === value, 'WHITESPACE_FORMAT_RISK');
    if (!byPart.has(update.part)) byPart.set(update.part, []);
    byPart.get(update.part).push({ start: update.node.openEnd, end: update.node.contentEnd, value: escapeText(value) });
  }
  const expected = new Map(doc.parts);
  for (const [part, changes] of byPart) {
    let xml = doc.parts.get(part).toString('utf8');
    for (const c of changes.sort((a,b) => b.start-a.start)) xml = xml.slice(0,c.start) + c.value + xml.slice(c.end);
    parseXml(xml);
    expected.set(part, Buffer.from(xml)); doc.zip.file(part, xml);
  }
  const output = await doc.zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  const result = await openDocx(output);
  check(result.parts.size === expected.size, 'PACKAGE_CHANGED');
  for (const [part, data] of expected) check(result.parts.get(part)?.equals(data), 'UNAUTHORIZED_PACKAGE_CHANGE');
  validateOperationResult(doc,result,profile,plan,paragraphsOf,parseXml);
  return { bytes: output, audit: { schemaVersion: 1, scope: 'synthetic-only', inputSha256: doc.hash, outputSha256: result.hash,
    changedParts: [...byPart.keys()], changes: audit, unchangedParts: doc.parts.size - byPart.size,
    validation: 'exact-package-comparison-passed', operation, legacyDataValidation:'all-mapped-occurrences-verified', visualValidation: 'not-performed-by-engine' } };
}

module.exports = { sha256, checkZip, parseXml, paragraphsOf, textRangeSafe, openDocx, privacyFindings, applyAuthorized };
