/* Preserve signature alignment across the browser editor and Outlook's HTML renderer. */
(function (root) {
  "use strict";

  function alignment(value) {
    var v = String(value || "").trim().toLowerCase();
    return /^(left|center|right|justify)$/.test(v) ? v : "";
  }

  function stabilize(box, maxWidth) {
    var blocks = box.querySelectorAll("p,div,td,th");
    for (var i = 0; i < blocks.length; i++) {
      var block = blocks[i];
      var align = alignment(block.style.textAlign || block.getAttribute("align"));
      if (align) {
        block.style.textAlign = align;
        block.setAttribute("align", align); // honored by Outlook's Word renderer too
      }
    }
    var imgs = box.querySelectorAll("img");
    for (var j = 0; j < imgs.length; j++) {
      var img = imgs[j];
      img.style.maxWidth = "100%";
      var explicit = img.getAttribute("width") || img.style.width;
      if (!explicit && img.naturalWidth) {
        // Freeze display size before compression changes the image's intrinsic size.
        var width = Math.min(img.naturalWidth, maxWidth || 320);
        img.setAttribute("width", String(width));
        if (!img.getAttribute("height") && !img.style.height && img.naturalHeight) {
          img.setAttribute("height", String(Math.max(1, Math.round(width * img.naturalHeight / img.naturalWidth))));
        }
      }
      if (/^\d+(\.\d+)?px$/i.test(img.style.width)) {
        img.setAttribute("width", String(parseFloat(img.style.width)));
      }
      if (/^\d+(\.\d+)?px$/i.test(img.style.height)) {
        img.setAttribute("height", String(parseFloat(img.style.height)));
      }
    }
    return box;
  }

  function imagesBelowText(box) {
    var imgs = Array.from(box.querySelectorAll("img"));
    if (!imgs.length) { return box; }
    var footer = document.createElement("div");
    footer.setAttribute("data-raa-image-footer", "true");
    footer.setAttribute("align", "left");
    footer.style.cssText = "text-align:left;clear:both;margin:0;padding:0";
    imgs.forEach(function (img) {
      // Keep an image-only link clickable when moving a linked logo.
      var link = img.closest("a");
      var content = link && !link.textContent.trim() && link.querySelectorAll("img").length === 1 ? link : img;
      var parent = content.parentElement;
      content.remove();
      // Remove empty image paragraphs/wrappers, but preserve the text's tables.
      while (parent && parent !== box && /^(P|DIV|SPAN|A)$/.test(parent.tagName) &&
          !parent.textContent.trim() && !parent.querySelector("img,table,hr,input")) {
        var next = parent.parentElement;
        parent.remove();
        parent = next;
      }
      img.removeAttribute("align");
      img.style.display = "block";
      img.style.cssFloat = "none";
      img.style.position = "static";
      ["left", "right", "top", "bottom"].forEach(function (key) { img.style.removeProperty(key); });
      img.style.marginLeft = "0";
      img.style.marginRight = "0";
      var line = document.createElement("p");
      line.setAttribute("align", "left");
      line.style.cssText = "text-align:left;clear:both;margin:0;padding:0";
      line.appendChild(content);
      footer.appendChild(line);
    });
    box.appendChild(footer);
    return box;
  }

  // Outlook/Word clipboard HTML may put centering in a stylesheet that is
  // removed during paste cleanup. Materialize that alignment before removing it.
  function clipboardAlignment(box, html) {
    var initial = new Set();
    var els = box.querySelectorAll("*");
    for (var i = 0; i < els.length; i++) {
      if (els[i].style.textAlign || els[i].getAttribute("align")) { initial.add(els[i]); }
    }
    var styles = String(html).match(/<style\b[^>]*>[\s\S]*?<\/style>/gi) || [];
    styles.forEach(function (style) {
      var css = style.replace(/<\/?style\b[^>]*>/gi, "").replace(/\/\*[\s\S]*?\*\//g, "").replace(/<!--|-->/g, "");
      css.replace(/([^{}]+)\{([^{}]*)\}/g, function (_, selector, declarations) {
        var probe = document.createElement("span");
        probe.style.cssText = declarations;
        var align = alignment(probe.style.textAlign);
        if (!align) { return _; }
        try {
          var matches = box.querySelectorAll(selector.trim());
          for (var j = 0; j < matches.length; j++) {
            if (!initial.has(matches[j])) { matches[j].style.textAlign = align; }
          }
        } catch (e) { /* Ignore Word-specific selectors unsupported by the browser. */ }
        return _;
      });
    });
  }

  function prepare(html) {
    var box = document.createElement("div");
    box.innerHTML = html;
    stabilize(box, 320);
    imagesBelowText(box);
    // Persist inherited alignment too: a quoted document's broad p/div rules
    // must not override the signature when the two fragments share a body.
    var blocks = box.querySelectorAll("p,div,td,th");
    for (var i = 0; i < blocks.length; i++) {
      var node = blocks[i];
      var align = "";
      while (node && node !== box && !align) {
        align = alignment(node.style.textAlign || node.getAttribute("align"));
        node = node.parentElement;
      }
      align = align || "left";
      blocks[i].style.textAlign = align;
      blocks[i].setAttribute("align", align);
    }
    // An explicit container keeps quoted-thread CSS/alignment from repositioning
    // the signature. Its descendants retain their own chosen alignment.
    return '<div align="left" style="text-align:left;margin:0;padding:0">' + box.innerHTML + "</div>";
  }

  root.SignatureLayout = { stabilize: stabilize, imagesBelowText: imagesBelowText, clipboardAlignment: clipboardAlignment, prepare: prepare };
})(typeof self !== "undefined" ? self : this);
