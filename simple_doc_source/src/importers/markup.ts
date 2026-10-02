/**
 * A small, lenient HTML/XML parser that builds a plain node tree.
 *
 * It never touches a live DOM, so nothing in an imported file can run a script,
 * load a stylesheet, fetch a picture or resolve an external entity. HTML mode follows
 * the parts of the HTML parsing rules that matter for documents (void elements, raw
 * text elements, implied end tags for p/li/dt/dd/tr/td/option); XML mode is strict
 * about nesting and keeps prefixed names ("text:p") as they are.
 */

export interface DomText {
  type: "text";
  text: string;
  /** Keep the text exactly (no whitespace collapsing). */
  pre?: boolean;
}

export interface DomElement {
  type: "element";
  /** Lower-case in HTML mode; as written (with prefix) in XML mode. */
  name: string;
  attrs: Record<string, string>;
  children: DomNode[];
  parent: DomElement | null;
  /** Pre-parsed CSS declarations (front-ends may set this instead of a style attribute). */
  css?: Record<string, string>;
}

export type DomNode = DomText | DomElement;

export interface ParseOptions {
  xml?: boolean;
}

const VOID_ELEMENTS = new Set(["area", "base", "basefont", "bgsound", "br", "col", "embed", "frame", "hr", "img", "input", "keygen", "link", "meta", "param", "source", "track", "wbr"]);
const RAW_TEXT_ELEMENTS = new Set(["script", "style", "textarea", "title", "xmp", "iframe", "noembed", "noframes", "noscript", "plaintext"]);
// Opening one of these closes an open <p> (HTML "close a p element").
const CLOSES_P = new Set([
  "address", "article", "aside", "blockquote", "center", "details", "dialog", "dir", "div", "dl", "fieldset", "figcaption", "figure",
  "footer", "form", "h1", "h2", "h3", "h4", "h5", "h6", "header", "hgroup", "hr", "main", "menu", "nav", "ol", "p", "pre", "section",
  "summary", "table", "ul", "listing",
]);
// Elements that stop the search for an open element to close implicitly.
const SCOPE_BOUNDARY = new Set(["table", "td", "th", "caption", "html", "body", "object", "marquee", "applet", "template"]);
const LIST_BOUNDARY = new Set(["ul", "ol", "menu", "dir", ...SCOPE_BOUNDARY]);

const XML_ENTITIES: Record<string, string> = { amp: "&", lt: "<", gt: ">", quot: "\"", apos: "'" };
// Latin-1 names in code-point order from U+00A0.
const LATIN1_NAMES = (
  "nbsp iexcl cent pound curren yen brvbar sect uml copy ordf laquo not shy reg macr deg plusmn sup2 sup3 acute micro para middot cedil " +
  "sup1 ordm raquo frac14 frac12 frac34 iquest Agrave Aacute Acirc Atilde Auml Aring AElig Ccedil Egrave Eacute Ecirc Euml Igrave Iacute " +
  "Icirc Iuml ETH Ntilde Ograve Oacute Ocirc Otilde Ouml times Oslash Ugrave Uacute Ucirc Uuml Yacute THORN szlig agrave aacute acirc atilde " +
  "auml aring aelig ccedil egrave eacute ecirc euml igrave iacute icirc iuml eth ntilde ograve oacute ocirc otilde ouml divide oslash ugrave " +
  "uacute ucirc uuml yacute thorn yuml"
).split(" ");
const GREEK = "Alpha Beta Gamma Delta Epsilon Zeta Eta Theta Iota Kappa Lambda Mu Nu Xi Omicron Pi Rho _ Sigma Tau Upsilon Phi Chi Psi Omega".split(" ");
const HTML_ENTITIES: Record<string, string> = (() => {
  const table: Record<string, string> = { ...XML_ENTITIES };
  LATIN1_NAMES.forEach((name, index) => { table[name] = String.fromCharCode(0xa0 + index); });
  GREEK.forEach((name, index) => {
    if (name === "_") return;
    table[name] = String.fromCharCode(0x391 + index);
    table[name.toLowerCase()] = String.fromCharCode(0x3b1 + index);
  });
  Object.assign(table, {
    sigmaf: "ς", thetasym: "ϑ", upsih: "ϒ", piv: "ϖ", OElig: "Œ", oelig: "œ", Scaron: "Š", scaron: "š", Yuml: "Ÿ", fnof: "ƒ", circ: "ˆ", tilde: "˜",
    ensp: " ", emsp: " ", thinsp: " ", zwnj: "\u200c", zwj: "\u200d", lrm: "\u200e", rlm: "\u200f", ndash: "–", mdash: "—",
    lsquo: "‘", rsquo: "’", sbquo: "‚", ldquo: "“", rdquo: "”", bdquo: "„", dagger: "†", Dagger: "‡", bull: "•", hellip: "…", permil: "‰",
    prime: "′", Prime: "″", lsaquo: "‹", rsaquo: "›", oline: "‾", frasl: "⁄", euro: "€", image: "ℑ", weierp: "℘", real: "ℜ", trade: "™",
    alefsym: "ℵ", larr: "←", uarr: "↑", rarr: "→", darr: "↓", harr: "↔", crarr: "↵", lArr: "⇐", uArr: "⇑", rArr: "⇒", dArr: "⇓", hArr: "⇔",
    forall: "∀", part: "∂", exist: "∃", empty: "∅", nabla: "∇", isin: "∈", notin: "∉", ni: "∋", prod: "∏", sum: "∑", minus: "−", lowast: "∗",
    radic: "√", prop: "∝", infin: "∞", ang: "∠", and: "∧", or: "∨", cap: "∩", cup: "∪", int: "∫", there4: "∴", sim: "∼", cong: "≅", asymp: "≈",
    ne: "≠", equiv: "≡", le: "≤", ge: "≥", sub: "⊂", sup: "⊃", nsub: "⊄", sube: "⊆", supe: "⊇", oplus: "⊕", otimes: "⊗", perp: "⊥", sdot: "⋅",
    lceil: "⌈", rceil: "⌉", lfloor: "⌊", rfloor: "⌋", lang: "⟨", rang: "⟩", loz: "◊", spades: "♠", clubs: "♣", hearts: "♥", diams: "♦",
    check: "✓", cross: "✗", star: "☆", starf: "★", Tab: "\t", NewLine: "\n", nbsp: "\u00a0", hyphen: "‐", dash: "‐", horbar: "―",
  });
  return table;
})();
// Entities a browser also accepts without the closing semicolon (the common ones).
const LEGACY_NO_SEMICOLON = new Set(["amp", "lt", "gt", "quot", "nbsp", "copy", "reg", "AMP", "LT", "GT", "QUOT"]);
// Windows-1252 meanings of C1 numeric references (HTML spec).
const C1_REPLACEMENTS: Record<number, number> = {
  0x80: 0x20ac, 0x82: 0x201a, 0x83: 0x0192, 0x84: 0x201e, 0x85: 0x2026, 0x86: 0x2020, 0x87: 0x2021, 0x88: 0x02c6, 0x89: 0x2030, 0x8a: 0x0160,
  0x8b: 0x2039, 0x8c: 0x0152, 0x8e: 0x017d, 0x91: 0x2018, 0x92: 0x2019, 0x93: 0x201c, 0x94: 0x201d, 0x95: 0x2022, 0x96: 0x2013, 0x97: 0x2014,
  0x98: 0x02dc, 0x99: 0x2122, 0x9a: 0x0161, 0x9b: 0x203a, 0x9c: 0x0153, 0x9e: 0x017e, 0x9f: 0x0178,
};

function codePointText(code: number): string {
  if (!Number.isFinite(code) || code <= 0 || code > 0x10ffff || (code >= 0xd800 && code <= 0xdfff)) return "�";
  return String.fromCodePoint(C1_REPLACEMENTS[code] ?? code);
}

/** Decodes character references in text or attribute values. */
export function decodeEntities(text: string, xml = false): string {
  if (!text.includes("&")) return text;
  const table = xml ? XML_ENTITIES : HTML_ENTITIES;
  return text.replace(/&(#[xX][0-9a-fA-F]{1,8}|#[0-9]{1,9}|[A-Za-z][A-Za-z0-9]{0,31})(;?)/g, (whole, name: string, semicolon: string) => {
    if (name[0] === "#") {
      const code = name[1] === "x" || name[1] === "X" ? parseInt(name.slice(2), 16) : parseInt(name.slice(1), 10);
      return codePointText(code);
    }
    if (semicolon && Object.prototype.hasOwnProperty.call(table, name)) return table[name];
    if (!xml && !semicolon && LEGACY_NO_SEMICOLON.has(name)) return HTML_ENTITIES[name.toLowerCase()] ?? whole;
    return whole;
  });
}

function element(name: string, attrs: Record<string, string>, parent: DomElement | null): DomElement {
  return { type: "element", name, attrs, children: [], parent };
}

const NAME_START = /[A-Za-z_:]/;

/** Parses HTML (default) or XML into a tree rooted at a "#document" element. */
export function parseMarkup(source: string, options: ParseOptions = {}): DomElement {
  const xml = options.xml === true;
  const root = element("#document", {}, null);
  const stack: DomElement[] = [root];
  const current = () => stack[stack.length - 1];
  const length = source.length;
  let index = 0;

  const appendText = (text: string) => {
    if (!text) return;
    const parent = current();
    const last = parent.children[parent.children.length - 1];
    if (last && last.type === "text" && !last.pre) last.text += text;
    else parent.children.push({ type: "text", text });
  };
  const hasInScope = (name: string, boundary: Set<string>) => {
    for (let i = stack.length - 1; i > 0; i -= 1) {
      if (stack[i].name === name) return i;
      if (boundary.has(stack[i].name)) return -1;
    }
    return -1;
  };
  const closeTo = (position: number) => {
    if (position > 0) stack.length = position;
  };
  const impliedEnd = (name: string) => {
    if (CLOSES_P.has(name)) closeTo(hasInScope("p", SCOPE_BOUNDARY));
    if (name === "li") closeTo(hasInScope("li", LIST_BOUNDARY));
    if (name === "dt" || name === "dd") {
      const dd = hasInScope("dd", LIST_BOUNDARY);
      const dt = hasInScope("dt", LIST_BOUNDARY);
      closeTo(Math.max(dd, dt));
    }
    if (name === "option") closeTo(hasInScope("option", SCOPE_BOUNDARY));
    if (name === "tr" || name === "thead" || name === "tbody" || name === "tfoot") {
      closeTo(hasInScope("td", new Set(["table"])));
      closeTo(hasInScope("th", new Set(["table"])));
      if (name === "tr") closeTo(hasInScope("tr", new Set(["table"])));
      else {
        closeTo(hasInScope("tr", new Set(["table"])));
        for (const section of ["thead", "tbody", "tfoot"]) closeTo(hasInScope(section, new Set(["table"])));
      }
    }
    if (name === "td" || name === "th") {
      closeTo(hasInScope("td", new Set(["table", "tr"])));
      closeTo(hasInScope("th", new Set(["table", "tr"])));
    }
  };

  while (index < length) {
    const lt = source.indexOf("<", index);
    if (lt < 0) {
      appendText(decodeEntities(source.slice(index), xml));
      break;
    }
    if (lt > index) appendText(decodeEntities(source.slice(index, lt), xml));
    index = lt;
    const next = source[index + 1];
    if (source.startsWith("<!--", index)) {
      const end = source.indexOf("-->", index + 4);
      index = end < 0 ? length : end + 3;
      continue;
    }
    if (source.startsWith("<![CDATA[", index)) {
      const end = source.indexOf("]]>", index + 9);
      const text = source.slice(index + 9, end < 0 ? length : end);
      if (xml) {
        const parent = current();
        parent.children.push({ type: "text", text, pre: false });
      }
      index = end < 0 ? length : end + 3;
      continue;
    }
    // Word's "downlevel-revealed" blocks (<![if !supportLists]>…<![endif]>) repeat what the
    // markup already says (list markers, note numbers); keep !vml fallbacks (pictures).
    if (!xml && /^<!\[if\s+!support/i.test(source.slice(index, index + 20))) {
      const end = source.slice(index).search(/<!\[endif\]>/i);
      index = end < 0 ? length : index + end + 10;
      continue;
    }
    if (next === "!" || next === "?") {
      const end = source.indexOf(">", index + 2);
      index = end < 0 ? length : end + 1;
      continue;
    }
    if (next === "/") {
      const match = /^<\/([^\s/>]+)[^>]*>?/.exec(source.slice(index, index + 300));
      if (!match || !NAME_START.test(match[1][0] ?? "")) {
        // "</" that is not an end tag: HTML treats it as a bogus comment.
        const end = source.indexOf(">", index + 2);
        index = end < 0 ? length : end + 1;
        continue;
      }
      index += match[0].length;
      const name = xml ? match[1] : match[1].toLowerCase();
      for (let i = stack.length - 1; i > 0; i -= 1) {
        if (stack[i].name === name) {
          stack.length = i;
          break;
        }
        if (!xml && SCOPE_BOUNDARY.has(stack[i].name) && name !== stack[i].name && (name === "p" || name === "li")) break;
      }
      continue;
    }
    if (!next || !NAME_START.test(next)) {
      appendText("<");
      index += 1;
      continue;
    }
    // Start tag.
    let cursor = index + 1;
    while (cursor < length && !/[\s/>]/.test(source[cursor])) cursor += 1;
    const rawName = source.slice(index + 1, cursor);
    const name = xml ? rawName : rawName.toLowerCase();
    const attrs: Record<string, string> = {};
    let selfClosing = false;
    while (cursor < length) {
      while (cursor < length && /\s/.test(source[cursor])) cursor += 1;
      const ch = source[cursor];
      if (ch === ">") {
        cursor += 1;
        break;
      }
      if (ch === "/") {
        if (source[cursor + 1] === ">") {
          selfClosing = true;
          cursor += 2;
          break;
        }
        cursor += 1;
        continue;
      }
      if (ch === undefined) break;
      let nameEnd = cursor;
      while (nameEnd < length && !/[\s=/>]/.test(source[nameEnd])) nameEnd += 1;
      if (nameEnd === cursor) nameEnd += 1;
      const attrName = xml ? source.slice(cursor, nameEnd) : source.slice(cursor, nameEnd).toLowerCase();
      cursor = nameEnd;
      while (cursor < length && /\s/.test(source[cursor])) cursor += 1;
      let value = "";
      if (source[cursor] === "=") {
        cursor += 1;
        while (cursor < length && /\s/.test(source[cursor])) cursor += 1;
        const quote = source[cursor];
        if (quote === "\"" || quote === "'") {
          const end = source.indexOf(quote, cursor + 1);
          value = source.slice(cursor + 1, end < 0 ? length : end);
          cursor = end < 0 ? length : end + 1;
        } else {
          let end = cursor;
          while (end < length && !/[\s>]/.test(source[end])) end += 1;
          value = source.slice(cursor, end);
          cursor = end;
        }
      }
      if (!(attrName in attrs)) attrs[attrName] = decodeEntities(value, xml);
    }
    index = cursor;
    if (!xml) impliedEnd(name);
    const node = element(name, attrs, current());
    current().children.push(node);
    if (xml) {
      if (!selfClosing) stack.push(node);
      continue;
    }
    if (VOID_ELEMENTS.has(name) || (selfClosing && !RAW_TEXT_ELEMENTS.has(name))) continue;
    if (RAW_TEXT_ELEMENTS.has(name)) {
      const closePattern = new RegExp(`</${name}\\s*>`, "i");
      const rest = source.slice(index);
      const match = closePattern.exec(rest);
      const raw = match ? rest.slice(0, match.index) : rest;
      if (raw) node.children.push({ type: "text", text: name === "title" || name === "textarea" ? decodeEntities(raw) : raw, pre: true });
      index += match ? match.index + match[0].length : rest.length;
      continue;
    }
    stack.push(node);
  }
  return root;
}

export function isElement(node: DomNode | null | undefined): node is DomElement {
  return !!node && node.type === "element";
}

/** Concatenated text of a subtree. */
export function textContent(node: DomNode): string {
  if (node.type === "text") return node.text;
  let text = "";
  for (const child of node.children) text += textContent(child);
  return text;
}

/** Element children with the given name (or all element children). */
export function childElements(node: DomElement, name?: string): DomElement[] {
  const result: DomElement[] = [];
  for (const child of node.children) if (child.type === "element" && (name === undefined || child.name === name)) result.push(child);
  return result;
}

/** Depth-first search for elements (pre-order). */
export function findElements(node: DomElement, predicate: (element: DomElement) => boolean, limit = Infinity): DomElement[] {
  const result: DomElement[] = [];
  const pending: DomElement[] = [node];
  while (pending.length && result.length < limit) {
    const current = pending.pop()!;
    if (current !== node && predicate(current)) result.push(current);
    for (let i = current.children.length - 1; i >= 0; i -= 1) {
      const child = current.children[i];
      if (child.type === "element") pending.push(child);
    }
  }
  return result;
}

export function firstElement(node: DomElement, name: string): DomElement | null {
  return findElements(node, (candidate) => candidate.name === name, 1)[0] ?? null;
}

export function hasClass(node: DomElement, name: string): boolean {
  const value = node.attrs.class;
  return !!value && value.split(/\s+/).some((token) => token.toLowerCase() === name.toLowerCase());
}

/** Builds an element for front-ends that synthesize trees (RTF, ODT). */
export function createElement(name: string, attrs: Record<string, string> = {}, children: DomNode[] = [], css?: Record<string, string>): DomElement {
  const node = element(name, attrs, null);
  if (css) node.css = css;
  for (const child of children) appendChild(node, child);
  return node;
}

export function appendChild(parent: DomElement, child: DomNode): void {
  if (child.type === "element") child.parent = parent;
  parent.children.push(child);
}
