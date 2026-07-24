/* SlideMakingwithMe — core conversion engine
 * Word (.docx) -> ProPresenter 7 (.pro)
 * Everything runs in the browser; no server, no libraries.
 */

/* ---------------- protobuf wire format (generic) ---------------- */

function pbParse(bytes) {
  const items = [];
  let i = 0;
  const readVarint = () => {
    let r = 0n, s = 0n;
    for (;;) {
      const x = bytes[i++];
      if (x === undefined) throw new Error("varint overrun");
      r |= BigInt(x & 0x7f) << s;
      if (!(x & 0x80)) return r;
      s += 7n;
    }
  };
  while (i < bytes.length) {
    const tag = Number(readVarint());
    const f = tag >>> 3, wt = tag & 7;
    if (f === 0) throw new Error("field 0");
    if (wt === 0) items.push({ f, wt, v: readVarint() });
    else if (wt === 1) { items.push({ f, wt, v: bytes.slice(i, i + 8) }); i += 8; }
    else if (wt === 5) { items.push({ f, wt, v: bytes.slice(i, i + 4) }); i += 4; }
    else if (wt === 2) {
      const len = Number(readVarint());
      if (i + len > bytes.length) throw new Error("bytes overrun");
      items.push({ f, wt, v: bytes.slice(i, i + len) }); i += len;
    } else throw new Error("bad wire type " + wt);
  }
  return items;
}

function pbEncode(items) {
  const out = [];
  const pushVarint = (n) => {
    let v = BigInt(n);
    for (;;) {
      const b = Number(v & 0x7fn);
      v >>= 7n;
      if (v > 0n) out.push(b | 0x80); else { out.push(b); return; }
    }
  };
  for (const { f, wt, v } of items) {
    pushVarint(BigInt(f << 3 | wt));
    if (wt === 0) pushVarint(v);
    else if (wt === 1 || wt === 5) for (const b of v) out.push(b);
    else if (wt === 2) { pushVarint(BigInt(v.length)); for (const b of v) out.push(b); }
    else throw new Error("bad wire type");
  }
  return new Uint8Array(out);
}

const utf8 = (s) => new TextEncoder().encode(s);

function pbFirst(items, f) { return items.find(it => it.f === f); }

/* Replace the value of the first occurrence of field f (wire type 2). */
function pbSet(items, f, bytes) {
  const it = pbFirst(items, f);
  if (!it) throw new Error("field " + f + " not found");
  it.v = bytes;
}

function newUuid() {
  return crypto.randomUUID().toUpperCase();
}

function uuidMsg(u) {
  return pbEncode([{ f: 1, wt: 2, v: utf8(u) }]);
}

/* ---------------- .docx reading (zip + document.xml) ---------------- */

async function inflateRaw(data) {
  const ds = new DecompressionStream("deflate-raw");
  const stream = new Blob([data]).stream().pipeThrough(ds);
  const buf = await new Response(stream).arrayBuffer();
  return new Uint8Array(buf);
}

async function extractDocumentXml(arrayBuffer) {
  const b = new Uint8Array(arrayBuffer);
  const dv = new DataView(arrayBuffer);
  // locate End Of Central Directory record (PK\x05\x06)
  let eocd = -1;
  for (let i = b.length - 22; i >= Math.max(0, b.length - 65558); i--) {
    if (b[i] === 0x50 && b[i + 1] === 0x4b && b[i + 2] === 0x05 && b[i + 3] === 0x06) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error("Not a valid .docx file (zip directory not found).");
  const count = dv.getUint16(eocd + 10, true);
  let off = dv.getUint32(eocd + 16, true);
  const td = new TextDecoder();
  for (let n = 0; n < count; n++) {
    if (dv.getUint32(off, true) !== 0x02014b50) throw new Error("Bad zip central directory.");
    const method = dv.getUint16(off + 10, true);
    const csize = dv.getUint32(off + 20, true);
    const nameLen = dv.getUint16(off + 28, true);
    const extraLen = dv.getUint16(off + 30, true);
    const cmtLen = dv.getUint16(off + 32, true);
    const lho = dv.getUint32(off + 42, true);
    const name = td.decode(b.slice(off + 46, off + 46 + nameLen));
    if (name === "word/document.xml") {
      const lnl = dv.getUint16(lho + 26, true);
      const lel = dv.getUint16(lho + 28, true);
      const start = lho + 30 + lnl + lel;
      const comp = b.slice(start, start + csize);
      const raw = method === 8 ? await inflateRaw(comp) : comp;
      return new TextDecoder().decode(raw);
    }
    off += 46 + nameLen + extraLen + cmtLen;
  }
  throw new Error("word/document.xml not found — is this a .docx file?");
}

function decodeXmlEntities(s) {
  return s
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(parseInt(d, 10)))
    .replace(/&lt;/g, "<").replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

/* Returns the document's paragraphs as plain-text strings. */
function docxParagraphs(xml) {
  const paras = [];
  const pRe = /<w:p[ >][\s\S]*?<\/w:p>/g;
  let pm;
  while ((pm = pRe.exec(xml)) !== null) {
    const inner = pm[0];
    let text = "";
    const runRe = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:(tab|br|cr)\b[^>]*\/>/g;
    let rm;
    while ((rm = runRe.exec(inner)) !== null) {
      if (rm[1] !== undefined) text += decodeXmlEntities(rm[1]);
      else if (rm[2] === "tab") text += "\t";
      else text += "\n";
    }
    paras.push(text);
  }
  return paras;
}

/* ---------------- sermon slide building (P Jon Choi format) ---------------- */

const SCRIPTURE_RE = /^((?:[1-3]\s*)?[A-Za-z]+\.?\s*\d+\s*:\s*\d+(?:\s*[-–—]\s*\d+(?::\d+)?)?)(\s*)/;
const MAX_SLIDE_CHARS = 300;

/* Split verse text into chunks at verse-number boundaries so that
 * each chunk stays under MAX_SLIDE_CHARS (a single long verse is kept whole). */
function splitVerses(body, max = MAX_SLIDE_CHARS) {
  if (body.length <= max) return [body];
  // boundaries sit right before a standalone verse number ("... of them. 7  And the LORD ...")
  const bounds = [];
  const re = /(^|\s)(\d{1,3})(?=\s)/g;
  let m;
  while ((m = re.exec(body)) !== null) {
    const pos = m.index + m[1].length;
    if (pos > 0) bounds.push(pos);
  }
  const segs = [];
  let prev = 0;
  for (const p of bounds) { if (p > prev) { segs.push(body.slice(prev, p)); prev = p; } }
  segs.push(body.slice(prev));
  const chunks = [];
  let acc = "";
  for (const seg of segs) {
    if (acc && (acc + seg).trimEnd().length > max) { chunks.push(acc.trimEnd()); acc = seg; }
    else acc += seg;
  }
  if (acc.trim()) chunks.push(acc.trimEnd());
  return chunks;
}

/* paragraphs -> array of slide texts (plain text, \n for line breaks, \t for tabs) */
function buildSermonSlides(paragraphs) {
  const paras = paragraphs
    .map(p => p.replace(/\u00A0/g, " "))
    .filter(p => p.trim() !== "");

  const slides = [];
  let title = null, textRef = null;
  const titleExtras = []; // subtitle lines between Title: and Text: (e.g. series name)
  const rest = [];
  for (const p of paras) {
    const t = p.trim();
    if (/^title\s*:/i.test(t) && title === null) title = p.replace(/^\s+/, "");
    else if (/^text\s*:/i.test(t) && textRef === null) textRef = t.replace(/^text\s*:\s*/i, "");
    else if (/^background/i.test(t)) continue; // production note, not a slide
    else if (title !== null && textRef === null && !SCRIPTURE_RE.test(t)) titleExtras.push(t);
    else rest.push(p);
  }
  if (title !== null || textRef !== null || titleExtras.length) {
    slides.push([title, ...titleExtras, textRef].filter(x => x !== null && x !== "").join("\n"));
  }
  let pointNum = 0;
  for (const p of rest) {
    const t = p.trim();
    const m = t.match(SCRIPTURE_RE);
    if (m) {
      const ref = m[1];
      const ws = m[2];
      const body = t.slice(m[0].length);
      const chunks = splitVerses(body);
      slides.push(ref + ws + "\n" + chunks[0]);
      for (let i = 1; i < chunks.length; i++) slides.push(chunks[i]);
    } else if (/^\d{1,3}\s/.test(t)) {
      // paragraph starting with a bare verse number: scripture continued from the previous slide
      for (const chunk of splitVerses(t)) slides.push(chunk);
    } else {
      pointNum += 1;
      slides.push(pointNum + ".\t" + t);
    }
  }
  return slides;
}

/* ---------------- RTF generation (matches the ProPresenter template) ---------------- */

const CP1252_MAP = {
  0x2018: 0x91, 0x2019: 0x92, 0x201C: 0x93, 0x201D: 0x94,
  0x2013: 0x96, 0x2014: 0x97, 0x2026: 0x85, 0x2022: 0x95,
  0x20AC: 0x80, 0x201A: 0x82, 0x0192: 0x83, 0x201E: 0x84,
  0x2020: 0x86, 0x2021: 0x87, 0x02C6: 0x88, 0x2030: 0x89,
  0x0160: 0x8A, 0x2039: 0x8B, 0x0152: 0x8C, 0x017D: 0x8E,
  0x02DC: 0x98, 0x2122: 0x99, 0x0161: 0x9A, 0x203A: 0x9B,
  0x0153: 0x9C, 0x017E: 0x9E, 0x0178: 0x9F,
};

function rtfEscape(text) {
  let out = "";
  for (const ch of text) {
    const c = ch.codePointAt(0);
    if (ch === "\\") out += "\\\\";
    else if (ch === "{") out += "\\{";
    else if (ch === "}") out += "\\}";
    else if (ch === "\n") out += "\\\n";
    else if (ch === "\t") out += "\t";
    else if (c === 0xA0) out += " ";
    else if (c < 0x80) out += ch;
    else if (CP1252_MAP[c]) out += "\\'" + CP1252_MAP[c].toString(16);
    else if (c < 0x100) out += "\\'" + c.toString(16).padStart(2, "0");
    else out += "\\uc0\\u" + (c > 0x7fff ? c - 0x10000 : c) + " ";
  }
  return out;
}

function makeRtf(text) {
  return (
    "{\\rtf1\\ansi\\ansicpg1252\\cocoartf2757\n" +
    "\\cocoatextscaling0\\cocoaplatform0{\\fonttbl\\f0\\fnil\\fcharset0 Montserrat-Regular;}\n" +
    "{\\colortbl;\\red255\\green255\\blue255;\\red255\\green255\\blue255;\\red0\\green0\\blue0;}\n" +
    "{\\*\\expandedcolortbl;;\\cssrgb\\c100000\\c100000\\c100000;\\cssrgb\\c0\\c0\\c0;}\n" +
    "\\deftab1680\n" +
    "\\pard\\pardeftab1680\\slleading340\\pardirnatural\\partightenfactor0\n" +
    "\n" +
    "\\f0\\fs96 \\cf2 \\outl0\\strokewidth-20 \\strokec3 " + rtfEscape(text) + "}"
  );
}

/* ---------------- .pro assembly from the embedded template ---------------- */

/* Clone the template cue, swapping in fresh UUIDs and the slide's text. */
function buildCue(templateCueBytes, text) {
  const cueUuid = newUuid();
  const cue = pbParse(templateCueBytes);
  pbSet(cue, 1, uuidMsg(cueUuid));

  const action = pbParse(pbFirst(cue, 10).v);
  pbSet(action, 1, uuidMsg(newUuid()));

  const slideWrap = pbParse(pbFirst(action, 23).v);   // Action.SlideType
  const slide = pbParse(pbFirst(slideWrap, 2).v);      // PresentationSlide
  const base = pbParse(pbFirst(slide, 1).v);           // BaseSlide
  pbSet(base, 7, uuidMsg(newUuid()));

  const element = pbParse(pbFirst(base, 1).v);         // Slide.Element
  const textEl = pbParse(pbFirst(element, 1).v);       // Graphics.Element
  pbSet(textEl, 1, uuidMsg(newUuid()));
  pbSet(textEl, 2, utf8(text));

  const textAttrs = pbParse(pbFirst(textEl, 13).v);    // Graphics.Text
  pbSet(textAttrs, 5, utf8(makeRtf(text)));

  // re-encode the chain bottom-up
  pbSet(textEl, 13, pbEncode(textAttrs));
  pbSet(element, 1, pbEncode(textEl));
  pbSet(base, 1, pbEncode(element));
  pbSet(slide, 1, pbEncode(base));
  pbSet(slideWrap, 2, pbEncode(slide));
  pbSet(action, 23, pbEncode(slideWrap));
  pbSet(cue, 10, pbEncode(action));

  return { uuid: cueUuid, bytes: pbEncode(cue) };
}

/* slides: array of plain-text strings; returns Uint8Array of the .pro file */
function buildPro(templateBytes, slides, presentationName) {
  const top = pbParse(templateBytes);
  const templateCueBytes = pbFirst(top, 13).v;

  const cues = slides.map(text => buildCue(templateCueBytes, text));

  // rebuild the cue group (slide order)
  const groupItems = pbParse(pbFirst(top, 12).v);
  const groupInfo = pbParse(pbFirst(groupItems, 1).v);
  pbSet(groupInfo, 1, uuidMsg(newUuid()));
  const newGroup = [
    { f: 1, wt: 2, v: pbEncode(groupInfo) },
    ...cues.map(c => ({ f: 2, wt: 2, v: uuidMsg(c.uuid) })),
  ];

  const out = [];
  let cuesInserted = false;
  for (const it of top) {
    if (it.f === 2) out.push({ f: 2, wt: 2, v: uuidMsg(newUuid()) });
    else if (it.f === 3) out.push({ f: 3, wt: 2, v: utf8(presentationName) });
    else if (it.f === 12) out.push({ f: 12, wt: 2, v: pbEncode(newGroup) });
    else if (it.f === 13) {
      if (!cuesInserted) {
        cuesInserted = true;
        for (const c of cues) out.push({ f: 13, wt: 2, v: c.bytes });
      }
    } else out.push(it);
  }
  return pbEncode(out);
}

/* ---------------- top-level pipeline ---------------- */

async function convertSermonDocx(arrayBuffer, presentationName, templateB64) {
  const xml = await extractDocumentXml(arrayBuffer);
  const paragraphs = docxParagraphs(xml);
  const slides = buildSermonSlides(paragraphs);
  if (slides.length === 0) throw new Error("No slide content found in the document.");
  const templateBytes = Uint8Array.from(atob(templateB64), c => c.charCodeAt(0));
  const pro = buildPro(templateBytes, slides, presentationName);
  return { slides, pro };
}
