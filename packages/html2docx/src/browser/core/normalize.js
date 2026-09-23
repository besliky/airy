  /**
   * Dirty-HTML hardening (BUG-1650): an inline formatting tag left open
   * across an implicit block boundary (<h1>text <b>bold</h1><p>...,
   * unclosed <p>/<li> after an open <b>/<i>/<span>) makes the HTML parser
   * reconstruct that inline element around every following block, so the
   * parsed tree holds <p>/<ul>/<table> INSIDE an inline <b>. Downstream
   * classification then flattens the whole document tail into one
   * run-paragraph and the lists/tables disappear.
   *
   * Undo that artifact before classification: split any inline element whose
   * direct children include block-level boxes into
   * (clone, hoisted blocks, clone) — the minimal re-association the implied
   * end tags (p closed before blocks, li closed before li/ul/ol, open
   * elements closed at the end of the document) were meant to produce.
   * Well-formed documents never place block boxes inside a non-<a> inline
   * element (<a> has a transparent content model, so an anchor wrapper is
   * valid authoring and stays untouched), and the walk rewrites nothing —
   * extraction stays bit-identical for them.
   */
  function normalizeUnclosedInlineWrappers() {
    const HTML_NS = 'http://www.w3.org/1999/xhtml';
    const isBlockBox = (node) => {
      if (node.nodeType !== Node.ELEMENT_NODE) return false;
      if (node.namespaceURI !== HTML_NS || SKIP_TAGS.has(node.tagName)) return false;
      const display = cs(node).display;
      return (
        display !== 'none' &&
        display !== 'contents' &&
        display !== 'inline' &&
        !display.startsWith('inline-')
      );
    };
    const isMisnestedWrapper = (el) =>
      el.namespaceURI === HTML_NS &&
      el.tagName !== 'A' &&
      !SKIP_TAGS.has(el.tagName) &&
      cs(el).display === 'inline' &&
      [...el.children].some(isBlockBox);
    const splitWrapper = (el) => {
      const frag = document.createDocumentFragment();
      let inlineRun = [];
      const flushRun = () => {
        if (!inlineRun.length) return;
        const clone = el.cloneNode(false);
        while (inlineRun.length) clone.appendChild(inlineRun.shift());
        frag.appendChild(clone);
      };
      for (const node of [...el.childNodes]) {
        if (isBlockBox(node)) {
          flushRun();
          frag.appendChild(node);
        } else {
          inlineRun.push(node);
        }
      }
      flushRun();
      el.parentNode.replaceChild(frag, el);
    };
    // Post-order: a child split can hoist blocks into the element itself,
    // which must then be re-tested as a wrapper too (b > i > p nesting).
    const visit = (el) => {
      for (const child of [...el.children]) visit(child);
      if (isMisnestedWrapper(el)) splitWrapper(el);
    };
    if (document.body) visit(document.body);
  }
  normalizeUnclosedInlineWrappers();
