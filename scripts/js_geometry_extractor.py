from typing import Final

JS_SLIDE_GEOMETRY_EXTRACTOR: Final[str] = """
(() => {
    const targetId = '%SLIDE_ID%';
    const slideIdx0 = %SLIDE_IDX_0%;
    const s = document.getElementById(targetId) || document.querySelector('.slide[data-active="true"]') || document.querySelector('.slide.active') || document.querySelector('.slide');
    if (!s) return null;

    try {
        for (const fnName of ['updateSlide', 'showSlide', 'goToSlide', 'setSlide', 'renderSlide']) {
            if (typeof window[fnName] === 'function') {
                window[fnName](slideIdx0);
                break;
            }
        }
    } catch (e) {}

    try {
        s.style.transition = 'none';
        s.style.transform = 'none';
        s.style.opacity = '1';
        s.style.visibility = 'visible';
        s.style.display = 'flex';
    } catch (e) {}

    try {
        s.querySelectorAll('canvas').forEach(cv => {
            if (window.Chart && typeof window.Chart.getChart === 'function') {
                const inst = window.Chart.getChart(cv);
                if (inst) {
                    inst.resize();
                    inst.update('none');
                }
            }
        });
    } catch (e) {}

    const slideRect = s.getBoundingClientRect();
    const slideW = slideRect.width || 1920;
    const slideH = slideRect.height || 1080;

    const toIn = (rect) => {
        const relLeft = Math.max(0, rect.left - slideRect.left);
        const relTop = Math.max(0, rect.top - slideRect.top);
        const inLeft = (relLeft / slideW) * 13.333;
        const inTop = (relTop / slideH) * 7.5;
        const inWidth = Math.min(13.333 - inLeft, (rect.width / slideW) * 13.333);
        const inHeight = Math.min(7.5 - inTop, (rect.height / slideH) * 7.5);
        return {
            left: Number(inLeft.toFixed(3)),
            top: Number(inTop.toFixed(3)),
            width: Number(inWidth.toFixed(3)),
            height: Number(inHeight.toFixed(3)),
            pxWidth: Math.round(rect.width),
            pxHeight: Math.round(rect.height)
        };
    };

    const parseRgba = (str) => {
        if (!str || str === 'transparent' || str === 'none') return null;
        const m = str.match(/rgba?\\(\\s*(\\d+)[,\\s]+(\\d+)[,\\s]+(\\d+)(?:[,\\s/]+([\\d.]+))?\\s*\\)/);
        if (m) {
            return {
                r: parseInt(m[1], 10),
                g: parseInt(m[2], 10),
                b: parseInt(m[3], 10),
                a: m[4] !== undefined ? parseFloat(m[4]) : 1.0
            };
        }
        const hm = str.trim().match(/^#([0-9a-fA-F]{3,8})$/);
        if (hm) {
            let hex = hm[1];
            if (hex.length === 3) {
                hex = hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
            }
            if (hex.length >= 6) {
                return {
                    r: parseInt(hex.slice(0, 2), 16),
                    g: parseInt(hex.slice(2, 4), 16),
                    b: parseInt(hex.slice(4, 6), 16),
                    a: hex.length === 8 ? parseInt(hex.slice(6, 8), 16) / 255.0 : 1.0
                };
            }
        }
        return null;
    };

    const parseLinearGradient = (bgImg) => {
        if (!bgImg || bgImg === 'none' || !bgImg.includes('linear-gradient')) return null;
        const lgIdx = bgImg.indexOf('linear-gradient(');
        if (lgIdx < 0) return null;
        let depth = 0;
        let body = '';
        for (let i = lgIdx + 16; i < bgImg.length; i++) {
            const ch = bgImg[i];
            if (ch === '(') depth++;
            else if (ch === ')') {
                if (depth === 0) break;
                depth--;
            }
            body += ch;
        }
        if (!body) return null;
        const parts = [];
        let cur = '';
        depth = 0;
        for (let i = 0; i < body.length; i++) {
            const ch = body[i];
            if (ch === '(') depth++;
            else if (ch === ')') depth--;
            if (ch === ',' && depth === 0) {
                parts.push(cur.trim());
                cur = '';
            } else {
                cur += ch;
            }
        }
        if (cur.trim()) parts.push(cur.trim());
        if (parts.length < 2) return null;

        let angle = 180;
        let startIdx = 0;
        const first = parts[0].toLowerCase();
        if (first.endsWith('deg')) {
            angle = parseFloat(first) || 180;
            startIdx = 1;
        } else if (first.startsWith('to ')) {
            startIdx = 1;
            if (first.includes('right') && first.includes('bottom')) angle = 135;
            else if (first.includes('right') && first.includes('top')) angle = 45;
            else if (first.includes('left') && first.includes('bottom')) angle = 225;
            else if (first.includes('left') && first.includes('top')) angle = 315;
            else if (first.includes('right')) angle = 90;
            else if (first.includes('left')) angle = 270;
            else if (first.includes('top')) angle = 0;
            else angle = 180;
        }

        const rawStops = [];
        for (let i = startIdx; i < parts.length; i++) {
            const p = parts[i];
            const cm = p.match(/(rgba?\\([^)]+\\)|#[0-9a-fA-F]{3,8})\\s*(?:([\\d.]+)%?)?/);
            if (cm) {
                rawStops.push({
                    color: cm[1],
                    pos: cm[2] !== undefined ? parseFloat(cm[2]) : null
                });
            }
        }
        if (rawStops.length < 2) return null;
        if (rawStops[0].pos === null) rawStops[0].pos = 0;
        if (rawStops[rawStops.length - 1].pos === null) rawStops[rawStops.length - 1].pos = 100;
        for (let i = 1; i < rawStops.length - 1; i++) {
            if (rawStops[i].pos === null) {
                let prevIdx = i - 1;
                let nextIdx = i + 1;
                while (nextIdx < rawStops.length - 1 && rawStops[nextIdx].pos === null) nextIdx++;
                const p0 = rawStops[prevIdx].pos;
                const p1 = rawStops[nextIdx].pos;
                rawStops[i].pos = Number((p0 + ((p1 - p0) * (i - prevIdx)) / (nextIdx - prevIdx)).toFixed(1));
            }
        }
        return {
            type: 'linear',
            angle: angle,
            stops: rawStops
        };
    };

    const extractBgColor = (cs) => {
        let bg = cs.backgroundColor;
        if ((!bg || bg === 'rgba(0, 0, 0, 0)' || bg === 'transparent') && cs.backgroundImage && cs.backgroundImage !== 'none') {
            const bgImg = cs.backgroundImage;
            const isTextClip = cs.webkitBackgroundClip === 'text' || cs.backgroundClip === 'text';
            if (!isTextClip) {
                const m = bgImg.match(/rgba?\\(\\s*\\d+\\s*,\\s*\\d+\\s*,\\s*\\d+(?:\\s*,\\s*[\\d.]+)?\\s*\\)|#[0-9a-fA-F]{3,8}/);
                if (m) bg = m[0];
            }
        }
        return bg;
    };

    const isCodeElement = (el, cs) => {
        if (!el || !el.tagName) return false;
        const tag = el.tagName.toLowerCase();
        if (tag === 'code' || tag === 'kbd' || tag === 'samp' || tag === 'pre') return true;
        const cls = typeof el.className === 'string' ? el.className : '';
        if (/\\b(mono|code)\\b/i.test(cls)) return true;
        const ff = (cs && cs.fontFamily) ? cs.fontFamily.toLowerCase() : '';
        return ff.includes('mono') || ff.includes('consolas') || ff.includes('courier');
    };

    const getStyles = (el, pseudo = null) => {
        if (!el) return null;
        const cs = window.getComputedStyle(el, pseudo);
        const pxSize = parseFloat(cs.fontSize) || 16;
        const scaledPt = Number(Math.max(8.5, Math.min(44, pxSize * 0.56)).toFixed(1));
        const br = parseFloat(cs.borderRadius) || parseFloat(cs.borderTopLeftRadius) || 0;
        const bw = parseFloat(cs.borderWidth) || parseFloat(cs.borderBottomWidth) || parseFloat(cs.borderLeftWidth) || 0;
        const btw = parseFloat(cs.borderTopWidth) || 0;
        const blw = parseFloat(cs.borderLeftWidth) || 0;
        let color = cs.color;
        if ((color === 'rgba(0, 0, 0, 0)' || color === 'transparent') && cs.backgroundImage && cs.backgroundImage !== 'none') {
            const m = cs.backgroundImage.match(/rgba?\\(\\s*\\d+\\s*,\\s*\\d+\\s*,\\s*\\d+(?:\\s*,\\s*[\\d.]+)?\\s*\\)|#[0-9a-fA-F]{3,8}/);
            if (m) color = m[0];
        }
        const grad = parseLinearGradient(cs.backgroundImage);
        const hasShadow = Boolean(cs.boxShadow && cs.boxShadow !== 'none');
        return {
            color: color,
            backgroundColor: extractBgColor(cs),
            gradient: grad,
            hasShadow: hasShadow,
            isCode: pseudo ? false : isCodeElement(el, cs),
            fontSizePt: scaledPt,
            fontWeight: cs.fontWeight,
            isBold: parseInt(cs.fontWeight) >= 600 || cs.fontWeight === 'bold',
            borderRadius: br,
            isRounded: br >= 4,
            borderTopColor: cs.borderTopColor,
            borderTopWidth: btw,
            borderLeftColor: cs.borderLeftColor,
            borderLeftWidth: blw,
            borderColor: bw > 0 ? (cs.borderColor || cs.borderBottomColor || cs.borderLeftColor) : null,
            borderWidth: bw,
            paddingLeft: parseFloat(cs.paddingLeft) || 0,
            paddingTop: parseFloat(cs.paddingTop) || 0,
            textAlign: cs.textAlign,
            display: cs.display
        };
    };

    const bgLayers = [];
    let curr = s;
    while (curr && curr !== document) {
        const cs = window.getComputedStyle(curr);
        const parsed = parseRgba(cs.backgroundColor);
        if (parsed && parsed.a > 0) {
            bgLayers.push(parsed);
            if (parsed.a >= 0.99) break;
        }
        curr = curr.parentElement;
    }

    let bgRgb = { r: 255, g: 255, b: 255 };
    if (bgLayers.length > 0) {
        const bottom = bgLayers[bgLayers.length - 1];
        bgRgb = { r: bottom.r, g: bottom.g, b: bottom.b };
        for (let i = bgLayers.length - 2; i >= 0; i--) {
            const top = bgLayers[i];
            const a = Math.min(1, Math.max(0, top.a));
            bgRgb = {
                r: Math.round(top.r * a + bgRgb.r * (1 - a)),
                g: Math.round(top.g * a + bgRgb.g * (1 - a)),
                b: Math.round(top.b * a + bgRgb.b * (1 - a))
            };
        }
    }
    const slideBg = `rgb(${bgRgb.r}, ${bgRgb.g}, ${bgRgb.b})`;

    let slideBgGradient = null;
    const sCs = window.getComputedStyle(s);
    if (sCs.backgroundImage && sCs.backgroundImage !== 'none') {
        const linG = parseLinearGradient(sCs.backgroundImage);
        if (linG) {
            slideBgGradient = linG;
        } else if (sCs.backgroundImage.includes('radial-gradient')) {
            const rgMatches = Array.from(sCs.backgroundImage.matchAll(/rgba?\\(\\s*(\\d+)[,\\s]+(\\d+)[,\\s]+(\\d+)(?:[,\\s/]+([\\d.]+))?\\s*\\)/g));
            const nonZero = rgMatches.map(m => ({
                r: parseInt(m[1], 10),
                g: parseInt(m[2], 10),
                b: parseInt(m[3], 10),
                a: m[4] !== undefined ? parseFloat(m[4]) : 1.0
            })).filter(c => c.a > 0.01);
            if (nonZero.length >= 1) {
                const c1 = nonZero[0];
                const c2 = nonZero.length > 1 ? nonZero[1] : null;
                const blendOnBg = (c) => {
                    const a = Math.min(0.25, Math.max(0.02, c.a));
                    return `rgb(${Math.round(c.r * a + bgRgb.r * (1 - a))}, ${Math.round(c.g * a + bgRgb.g * (1 - a))}, ${Math.round(c.b * a + bgRgb.b * (1 - a))})`;
                };
                slideBgGradient = {
                    type: 'linear',
                    angle: 135,
                    stops: [
                        { color: c2 ? blendOnBg(c2) : slideBg, pos: 0 },
                        { color: slideBg, pos: 50 },
                        { color: blendOnBg(c1), pos: 100 }
                    ]
                };
            }
        }
    }

    const extractRichRuns = (rootEl, rootBgColor = null) => {
        if (!rootEl) return null;
        const rootSt = getStyles(rootEl);
        const effRootBg = rootBgColor || rootSt.backgroundColor || slideBg;
        const rootFw = parseInt(rootSt.fontWeight) || 400;
        const hasBoldChild = Boolean(rootEl.querySelector('b, strong'));
        const boldThreshold = (rootFw === 600 && hasBoldChild) ? 700 : 600;
        const rootIsBold = rootFw >= boldThreshold || rootSt.fontWeight === 'bold';
        const runs = [];
        let pendingNewline = false;

        const walk = (node, currEl, currSt) => {
            if (node.nodeType === 3) {
                let raw = node.textContent || '';
                raw = raw.replace(/[\\r\\n\\t]+/g, ' ').replace(/\\s{2,}/g, ' ');
                if (!raw.trim()) {
                    if (raw.includes(' ') && runs.length > 0 && !pendingNewline && !runs[runs.length - 1].text.endsWith(' ')) {
                        runs[runs.length - 1].text += ' ';
                    }
                    return;
                }
                if (pendingNewline) {
                    raw = raw.replace(/^\\s+/, '');
                } else if (runs.length === 0) {
                    raw = raw.replace(/^\\s+/, '');
                }
                const elBg = currSt.backgroundColor;
                const hasInlineBg = currEl !== rootEl && elBg && elBg !== 'transparent' && elBg !== 'rgba(0, 0, 0, 0)' && elBg !== effRootBg && elBg !== slideBg;
                const cls = (currEl && typeof currEl.className === 'string') ? currEl.className : '';
                const isBadge = Boolean(hasInlineBg && (currSt.borderRadius >= 4 || /badge|tag|score|pill|chip/i.test(cls)) && !currSt.isCode);
                const currFw = parseInt(currSt.fontWeight) || 400;
                const currTag = (currEl && currEl.tagName) ? currEl.tagName.toLowerCase() : '';
                const runIsBold = currFw >= boldThreshold || currSt.fontWeight === 'bold' || currTag === 'b' || currTag === 'strong';
                runs.push({
                    text: raw,
                    color: currSt.color,
                    bgColor: hasInlineBg ? elBg : null,
                    fontSizePt: currSt.fontSizePt,
                    isBold: runIsBold,
                    isCode: currSt.isCode,
                    isBadge: isBadge,
                    newline: pendingNewline
                });
                pendingNewline = false;
            } else if (node.nodeType === 1) {
                const tag = node.tagName.toLowerCase();
                if (tag === 'br') {
                    if (runs.length > 0) pendingNewline = true;
                    return;
                }
                const st = getStyles(node);
                if (st.display === 'none') return;
                const isBlock = st.display === 'block' || tag === 'div' || tag === 'p';
                if (isBlock && runs.length > 0) {
                    pendingNewline = true;
                }
                Array.from(node.childNodes).forEach(ch => walk(ch, node, st));
                if (isBlock && runs.length > 0) {
                    pendingNewline = true;
                }
            }
        };

        Array.from(rootEl.childNodes).forEach(ch => walk(ch, rootEl, rootSt));
        if (runs.length > 0) {
            runs[runs.length - 1].text = runs[runs.length - 1].text.replace(/\\s+$/, '');
        }
        const hasRichFormatting = runs.length > 1 || runs.some(r => r.bgColor || r.isCode || r.isBadge || r.newline || r.isBold !== rootIsBold || r.color !== rootSt.color || r.fontSizePt !== rootSt.fontSizePt);
        return hasRichFormatting ? runs : null;
    };

    const headings = Array.from(s.querySelectorAll('h1, h2, h3, h4, h5, h6')).filter(el => {
        if (el.closest('header') || el.closest('footer') || el.closest('.foot')) return false;
        const cs = window.getComputedStyle(el);
        return cs.display !== 'none' && cs.visibility !== 'hidden';
    });

    let mainH = headings.find(h => h.tagName.toLowerCase() === 'h1');
    if (!mainH && headings.length > 0) mainH = headings[0];
    if (!mainH) {
        const prominent = Array.from(s.querySelectorAll('div, p, span')).filter(el => {
            const cs = window.getComputedStyle(el);
            const fs = parseFloat(cs.fontSize) || 0;
            const fw = parseInt(cs.fontWeight) || 400;
            const r = el.getBoundingClientRect();
            return fs >= 22 && fw >= 600 && r.top < 400 && el.innerText.trim().length > 0;
        });
        if (prominent.length > 0) mainH = prominent[0];
    }

    let subDesc = Array.from(s.querySelectorAll('p.subtitle, p.premise, p.lead, p.cover-subtitle, h2.subtitle, [class*="subtitle"], [class*="premise"], p.text-slate-400')).find(el => {
        if (!mainH) return true;
        if (el === mainH || mainH.contains(el) || el.contains(mainH)) return false;
        if (el.closest('.glass-card, .card, .panel, .info-card, .timeline, .pipe, .steps, .pcol, table')) return false;
        const rIn = toIn(el.getBoundingClientRect());
        if (rIn.top > 3.2) return false;
        return el.innerText.trim().length > 0;
    }) || null;
    if (!subDesc && mainH) {
        const headWrap = mainH.closest('.slide-head');
        const cand = headWrap ? (headWrap.querySelector('p') || headWrap.nextElementSibling) : mainH.nextElementSibling;
        if (cand && cand !== mainH && cand.tagName.toLowerCase() === 'p') {
            const rIn = toIn(cand.getBoundingClientRect());
            const txt = cand.innerText ? cand.innerText.trim() : '';
            if (txt && rIn.top <= 2.8 && !txt.startsWith('※') && !txt.startsWith('*') && !txt.startsWith('출처')) {
                subDesc = cand;
            }
        }
    }

    const headerBadges = [];
    const headerBadgeEls = [];
    const seenBadges = new Set();
    const mainHTopIn = mainH ? toIn(mainH.getBoundingClientRect()).top : 1.0;
    const maxBadgeTopIn = Math.max(1.45, mainHTopIn + 0.05);
    const addHeaderBadge = (el) => {
        if (!el || el === mainH || el === subDesc) return;
        if (mainH && (mainH.contains(el) || el.contains(mainH))) return;
        if (subDesc && (subDesc.contains(el) || el.contains(subDesc))) return;
        if (el.closest('.glass-card, .card, .panel, .info-card, table')) return;
        const txt = el.innerText ? el.innerText.trim().replace(/\\s+/g, ' ') : '';
        if (!txt || txt.length > 60 || seenBadges.has(txt)) return;
        const rIn = toIn(el.getBoundingClientRect());
        if (rIn.width < 0.2 || rIn.height < 0.1 || rIn.top > maxBadgeTopIn) return;
        seenBadges.add(txt);
        headerBadgeEls.push(el);
        headerBadges.push({
            text: txt,
            rect: rIn,
            styles: getStyles(el)
        });
    };

    s.querySelectorAll('.kicker, [class*="kicker"], .slide-head .eyebrow, .slide-head .badge-pill, .slide-head [class*="tag"], .slide-head [class*="badge"], .slide-head [class*="pill"], .slide-tag, [class*="eyebrow"]').forEach(addHeaderBadge);

    s.querySelectorAll('div, span').forEach(el => {
        if (el.children.length > 2) return;
        if (el.querySelector('h1, h2, h3, p, div')) return;
        const rIn = toIn(el.getBoundingClientRect());
        if (rIn.top > maxBadgeTopIn || rIn.height > 0.45 || rIn.width > 5.5) return;
        const cls = typeof el.className === 'string' ? el.className : '';
        const cs = window.getComputedStyle(el);
        const bg = extractBgColor(cs);
        const hasPillBg = bg && bg !== 'rgba(0, 0, 0, 0)' && bg !== 'transparent' && bg !== slideBg;
        const isEyebrowAboveTitle = rIn.top < mainHTopIn && (cls.includes('uppercase') || cls.includes('tracking-') || cs.textTransform === 'uppercase' || parseInt(cs.fontWeight) >= 600);
        const isHeaderPill = rIn.top <= 1.35 && (hasPillBg || cls.includes('badge') || cls.includes('rounded'));
        if (isEyebrowAboveTitle || isHeaderPill) {
            addHeaderBadge(el);
        }
    });

    const num = s.querySelector('.slide-number, [class*="slide-counter"], [class*="foot__n"]');

    const extractSvgPrimitives = (svgEl) => {
        if (svgEl.querySelector('path, image, use, foreignObject')) return null;
        const primNodes = Array.from(svgEl.querySelectorAll('rect, circle, ellipse, polygon, line, text')).filter(n => !n.closest('defs, clipPath, mask, filter, pattern'));
        if (primNodes.length === 0) return null;

        const prims = [];
        primNodes.forEach(node => {
            const nr = node.getBoundingClientRect();
            if (nr.width < 1 && nr.height < 1) return;
            const ncs = window.getComputedStyle(node);
            if (ncs.display === 'none' || ncs.visibility === 'hidden') return;
            const tag = node.tagName.toLowerCase();

            let effOpacity = 1.0;
            let cur = node;
            while (cur && cur !== svgEl) {
                const ccs = window.getComputedStyle(cur);
                const op = parseFloat(ccs.opacity);
                if (!isNaN(op)) effOpacity *= op;
                cur = cur.parentElement;
            }
            const fillOp = parseFloat(ncs.fillOpacity);
            if (!isNaN(fillOp)) effOpacity *= fillOp;

            let fillColor = null;
            let fillGradient = null;
            const rawFillAttr = node.getAttribute('fill') || '';
            const compFill = ncs.fill || rawFillAttr;
            if (rawFillAttr.startsWith('url(') || compFill.startsWith('url(')) {
                const urlStr = rawFillAttr.startsWith('url(') ? rawFillAttr : compFill;
                const idMatch = urlStr.match(/#([^)'"\\s]+)/);
                if (idMatch) {
                    const gradEl = svgEl.querySelector('#' + CSS.escape(idMatch[1])) || document.getElementById(idMatch[1]);
                    if (gradEl && gradEl.tagName.toLowerCase().includes('gradient')) {
                        const x1 = parseFloat(gradEl.getAttribute('x1') || '0');
                        const y1 = parseFloat(gradEl.getAttribute('y1') || '0');
                        const x2 = parseFloat(gradEl.getAttribute('x2') || '1');
                        const y2 = parseFloat(gradEl.getAttribute('y2') || '0');
                        const angle = Math.round((Math.atan2(y2 - y1, x2 - x1) * 180) / Math.PI + 90);
                        const stops = Array.from(gradEl.querySelectorAll('stop')).map((stEl, idx, arr) => {
                            let off = stEl.getAttribute('offset') || (idx === 0 ? '0' : '1');
                            let pos = parseFloat(off);
                            if (!off.includes('%') && pos <= 1.0) pos = pos * 100;
                            const stCs = window.getComputedStyle(stEl);
                            const stCol = stCs.stopColor || stEl.getAttribute('stop-color') || '#000000';
                            return { color: stCol, pos: Math.round(pos) };
                        });
                        if (stops.length >= 2) {
                            fillGradient = { type: 'linear', angle: (angle + 360) % 360, stops: stops };
                        }
                    }
                }
            } else if (compFill && compFill !== 'none' && compFill !== 'transparent' && compFill !== 'rgba(0, 0, 0, 0)') {
                fillColor = compFill;
            }

            let strokeColor = null;
            const compStroke = ncs.stroke || node.getAttribute('stroke');
            const strokeW = parseFloat(ncs.strokeWidth || node.getAttribute('stroke-width') || '0');
            if (compStroke && compStroke !== 'none' && strokeW > 0) {
                strokeColor = compStroke;
            }

            let rxRatio = 0;
            if (tag === 'rect') {
                const rx = parseFloat(node.getAttribute('rx') || node.getAttribute('ry') || ncs.rx || '0');
                const rw = parseFloat(node.getAttribute('width') || nr.width || '1');
                const rh = parseFloat(node.getAttribute('height') || nr.height || '1');
                const minDim = Math.min(rw, rh);
                if (rx > 0 && minDim > 0) {
                    rxRatio = Number(Math.min(0.5, rx / minDim).toFixed(3));
                }
            }

            let normPoints = null;
            if (tag === 'polygon') {
                const ptsAttr = (node.getAttribute('points') || '').trim();
                const nums = ptsAttr.split(/[\\s,]+/).map(parseFloat).filter(n => !isNaN(n));
                if (nums.length >= 6) {
                    const xs = [];
                    const ys = [];
                    for (let i = 0; i < nums.length - 1; i += 2) {
                        xs.push(nums[i]);
                        ys.push(nums[i + 1]);
                    }
                    const minX = Math.min(...xs);
                    const maxX = Math.max(...xs);
                    const minY = Math.min(...ys);
                    const maxY = Math.max(...ys);
                    const bw = Math.max(1, maxX - minX);
                    const bh = Math.max(1, maxY - minY);
                    normPoints = xs.map((x, idx) => [
                        Number(((x - minX) / bw).toFixed(4)),
                        Number(((ys[idx] - minY) / bh).toFixed(4))
                    ]);
                }
            }

            prims.push({
                tag: tag,
                rect: toIn(nr),
                fill: fillColor,
                gradient: fillGradient,
                stroke: strokeColor,
                strokeWidth: strokeW,
                opacity: Number(effOpacity.toFixed(2)),
                rxRatio: rxRatio,
                normPoints: normPoints,
                text: tag === 'text' ? (node.textContent || '').trim() : null,
                fontSizePt: tag === 'text' ? Number(Math.max(8, (parseFloat(ncs.fontSize) || 14) * 0.56).toFixed(1)) : null
            });
        });
        return prims.length > 0 ? prims : null;
    };

    const images = [];
    s.querySelectorAll('svg, img').forEach(el => {
        if (el.parentElement && el.parentElement.closest('svg')) return;
        const cs = window.getComputedStyle(el);
        if (cs.display === 'none' || cs.visibility === 'hidden' || parseFloat(cs.opacity) === 0) return;
        const r = el.getBoundingClientRect();
        if (r.width < 28 || r.height < 28) return;
        const tag = el.tagName.toLowerCase();
        images.push({
            tagName: tag,
            rect: toIn(r),
            svgPrimitives: tag === 'svg' ? extractSvgPrimitives(el) : null
        });
    });

    const charts = [];
    s.querySelectorAll('canvas').forEach(cv => {
        const r = cv.getBoundingClientRect();
        if (r.width < 30 || r.height < 30) return;
        let dataUrl = null;
        try {
            dataUrl = cv.toDataURL('image/png');
        } catch (e) {}

        let chartCfg = null;
        try {
            const inst = (window.Chart && typeof window.Chart.getChart === 'function') ? window.Chart.getChart(cv) : null;
            const rawCfg = (inst && inst.config) ? inst.config : cv.__chartConfig;
            if (rawCfg) {
                const cType = rawCfg.type || 'bar';
                const opts = rawCfg.options || {};
                const indexAxis = opts.indexAxis || 'x';
                const labels = Array.isArray(rawCfg.data?.labels) ? rawCfg.data.labels.map(String) : [];
                const datasets = (rawCfg.data?.datasets || []).map(ds => ({
                    label: ds.label ? String(ds.label) : '',
                    data: Array.isArray(ds.data) ? ds.data.map(v => typeof v === 'number' ? v : (parseFloat(v) || 0)) : [],
                    borderColor: ds.borderColor || null,
                    backgroundColor: ds.backgroundColor || null,
                    borderDash: Array.isArray(ds.borderDash) ? ds.borderDash : null,
                    borderWidth: ds.borderWidth || 2,
                    tension: ds.tension || 0,
                    fill: !!ds.fill
                }));
                chartCfg = {
                    type: cType,
                    indexAxis: indexAxis,
                    labels: labels,
                    datasets: datasets,
                    data: { labels: labels, datasets: datasets },
                    legendDisplay: opts.plugins?.legend?.display !== false,
                    legendPosition: opts.plugins?.legend?.position || 'top',
                    cutout: opts.cutout || null,
                    xDisplay: opts.scales?.x?.display !== false,
                    yDisplay: opts.scales?.y?.display !== false
                };
            }
        } catch (e) {}

        charts.push({
            id: cv.id || '',
            rect: toIn(r),
            dataUrl: dataUrl,
            chartConfig: chartCfg
        });
    });

    const hlBoxes = [];
    const hlBoxEls = [];
    const seenHl = new Set();
    s.querySelectorAll('.foot__pocket, .highlight-box, .banner, [class*="highlight-box"], [class*="callout"]').forEach(hl => {
        const txt = hl.innerText ? hl.innerText.trim().replace(/\\s*\\n\\s*/g, ' ') : '';
        if (!txt || seenHl.has(txt)) return;
        seenHl.add(txt);
        hlBoxEls.push(hl);
        const hlSt = getStyles(hl);
        hlBoxes.push({
            rect: toIn(hl.getBoundingClientRect()),
            styles: hlSt,
            text: txt,
            runs: extractRichRuns(hl, hlSt.backgroundColor)
        });
    });
    if (hlBoxes.length === 0) {
        s.querySelectorAll('*').forEach(hl => {
            if (hl.closest('table') || hl.querySelector('table')) return;
            if (hl.querySelector('.foot__pocket, .banner, [class*="highlight-box"], [class*="callout"]')) return;
            const txt = hl.innerText ? hl.innerText.trim().replace(/\\s*\\n\\s*/g, ' ') : '';
            if (!txt || seenHl.has(txt)) return;
            const isCallout = txt.startsWith('💡') || txt.startsWith('📌') || txt.startsWith('★') || 
                              txt.startsWith('⚠️') || txt.startsWith('ℹ️') || txt.includes('핵심 메시지') ||
                              txt.includes('핵심 요약') || txt.includes('참고') || txt.includes('결론');
            if (isCallout && hl.offsetHeight < 320 && hl.offsetWidth > 150) {
                seenHl.add(txt);
                hlBoxEls.push(hl);
                const hlSt = getStyles(hl);
                hlBoxes.push({
                    rect: toIn(hl.getBoundingClientRect()),
                    styles: hlSt,
                    text: txt,
                    runs: extractRichRuns(hl, hlSt.backgroundColor)
                });
            }
        });
    }

    const cardClassTokens = new Set([
        'panel', 'card', 'glass-card', 'info-card', 'pillar', 'ma-pillar',
        'runtime', 'sub-card', 'bridge-badge', 'bridge-col', 'f-item',
        'serving-box', 'attach-card', 'oss-card', 'template-box',
        'viz', 'pcol', 'step', 'node', 'side'
    ]);

    const hasExplicitCardClass = (el) => {
        const cls = (el && typeof el.className === 'string') ? el.className : '';
        return cls.split(/\\s+/).some(t => cardClassTokens.has(t));
    };

    const isTopContainer = (el) => {
        if (el === s || el.contains(s) || el.querySelector('table') || el.closest('table')) return false;
        if (el.closest('.slide-head') || el.closest('header') || el.closest('.foot') || el.closest('footer')) return false;
        if (headerBadgeEls.some(hb => hb === el || hb.contains(el) || el.contains(hb))) return false;
        if (hlBoxEls.some(hb => hb === el || hb.contains(el) || el.contains(hb))) return false;
        if (num && (el === num || num.contains(el) || el.contains(num))) return false;
        if (mainH && (el === mainH || el.contains(mainH) || mainH.contains(el))) return false;
        if (subDesc && (el === subDesc || el.contains(subDesc) || subDesc.contains(el))) return false;
        const tag = el.tagName ? el.tagName.toLowerCase() : '';
        if (tag === 'code' || tag === 'b' || tag === 'strong' || tag === 'em' || tag === 'small') return false;
        if (el.closest('p, li, h1, h2, h3, h4, h5, h6, pre, code')) return false;

        const cls = el.className || '';
        if (typeof cls !== 'string') return false;

        const r = el.getBoundingClientRect();
        if (r.width < 60 || r.height < 32 || r.width > slideW * 0.98 || r.height > slideH * 0.98) return false;
        const rIn = toIn(r);

        const cs = window.getComputedStyle(el);
        const bg = extractBgColor(cs);
        const hasBg = bg !== 'rgba(0, 0, 0, 0)' && bg !== 'transparent' && bg !== slideBg;
        const hasRadius = (parseFloat(cs.borderRadius) || parseFloat(cs.borderTopLeftRadius) || 0) >= 4;
        const hasFullBorder = parseFloat(cs.borderBottomWidth) > 0 && parseFloat(cs.borderLeftWidth) > 0;
        const hasAccentBorder = parseFloat(cs.borderTopWidth) >= 2 || parseFloat(cs.borderLeftWidth) >= 3;
        const hasBorder = hasFullBorder || hasAccentBorder;
        const hasShadow = cs.boxShadow && cs.boxShadow !== 'none';
        const hasCardClass = hasExplicitCardClass(el);

        if (!hasBg && !hasRadius && !hasCardClass && rIn.top > 6.2 && rIn.height < 0.65) return false;

        if (!hasBg && !hasRadius && !hasBorder && !hasShadow && !hasCardClass) return false;

        return true;
    };

    let allCandidateContainers = Array.from(s.querySelectorAll('*')).filter(isTopContainer);
    let topContainers = allCandidateContainers.filter(c => !allCandidateContainers.some(p => p !== c && p.contains(c) && isTopContainer(p)));
    if (topContainers.length === 0) topContainers = allCandidateContainers;

    const decorShapes = [];
    s.querySelectorAll('*').forEach(el => {
        if (el.closest('svg, table, pre, code')) return;
        const elRect = el.getBoundingClientRect();
        if (elRect.width <= 0 || elRect.height <= 0) return;

        for (const pseudo of ['::before', '::after']) {
            const pcs = window.getComputedStyle(el, pseudo);
            if (!pcs || pcs.content === 'none' || pcs.content === 'normal' || pcs.display === 'none') continue;
            const pw = parseFloat(pcs.width) || (pcs.left !== 'auto' && pcs.right !== 'auto' ? elRect.width - (parseFloat(pcs.left) || 0) - (parseFloat(pcs.right) || 0) : 0);
            const ph = parseFloat(pcs.height) || 0;
            if (pw < 4 || ph < 3) continue;
            const pBg = extractBgColor(pcs);
            const pGrad = parseLinearGradient(pcs.backgroundImage);
            const hasPFill = pGrad || (pBg && pBg !== 'rgba(0, 0, 0, 0)' && pBg !== 'transparent');
            const rawContent = pcs.content.replace(/^["']|["']$/g, '').trim();
            if (!hasPFill && !rawContent) continue;

            const pLeft = elRect.left + (parseFloat(pcs.left) || 0);
            const pTop = elRect.top + (parseFloat(pcs.top) || 0);
            const pRectIn = toIn({ left: pLeft, top: pTop, width: pw, height: ph });
            const pSt = getStyles(el, pseudo);
            decorShapes.push({
                rect: pRectIn,
                styles: pSt,
                text: rawContent || '',
                isCircle: pSt.borderRadius >= Math.min(pw, ph) * 0.45
            });
        }

        if (topContainers.includes(el)) return;
        const tag = el.tagName ? el.tagName.toLowerCase() : '';
        if (!['div', 'span', 'i', 'b'].includes(tag)) return;
        if ((el.innerText || '').trim().length > 0) return;
        if (el.querySelector('svg, img, canvas, table')) return;
        if (elRect.width < 6 || elRect.height < 6 || elRect.width > 300 || elRect.height > 300) return;
        const st = getStyles(el);
        const hasFill = st.gradient || (st.backgroundColor && st.backgroundColor !== 'rgba(0, 0, 0, 0)' && st.backgroundColor !== 'transparent' && st.backgroundColor !== slideBg);
        if (!hasFill && st.borderWidth <= 0) return;
        decorShapes.push({
            rect: toIn(elRect),
            styles: st,
            text: '',
            isCircle: st.borderRadius >= Math.min(elRect.width, elRect.height) * 0.45
        });
    });

    const cards = [];
    topContainers.forEach(c => {
        if (c.querySelector('table') || c.closest('table')) return;
        const cDomRect = c.getBoundingClientRect();
        const r = toIn(cDomRect);
        const st = getStyles(c);
        const hasCanvas = !!c.querySelector('canvas');
        const hasSvgOrImg = !!c.querySelector('svg, img');

        const bridgeBadgeEl = c.querySelector('.bridge-badge') || (c.classList.contains('bridge-badge') ? c : null);
        const isBridge = (c.className && c.className.includes('bridge')) || !!bridgeBadgeEl;
        let bridgeData = null;
        if (bridgeBadgeEl) {
            const bArrow = bridgeBadgeEl.querySelector('[class*="arrow"]');
            const bTxt = bridgeBadgeEl.querySelector('[class*="txt"]');
            bridgeData = {
                arrow: bArrow ? bArrow.innerText.trim() : '➔',
                text: bTxt ? bTxt.innerText.trim().replace(/\\n/g, ' ') : bridgeBadgeEl.innerText.trim().replace(/\\n/g, ' '),
                rect: toIn(bridgeBadgeEl.getBoundingClientRect()),
                styles: getStyles(bridgeBadgeEl)
            };
        }

        const isAttachCard = (c.className && c.className.includes('attach-card')) || !!c.querySelector('.attach-card__mech');
        let attachData = null;
        if (isAttachCard) {
            const tgtEl = c.querySelector('.attach-card__target, [class*="target"]');
            const modEl = c.querySelector('.attach-card__mode, [class*="mode"]');
            const mchEl = c.querySelector('.attach-card__mech, [class*="mech"]');
            const pEl = c.querySelector('p');
            attachData = {
                target: tgtEl ? tgtEl.innerText.trim() : '',
                mode: modEl ? modEl.innerText.trim() : '',
                mech: mchEl ? mchEl.innerText.trim() : '',
                desc: pEl ? pEl.innerText.trim() : ''
            };
        }

        const pipeStepEls = Array.from(c.querySelectorAll('.ma-pipe-step, .pipeline-step, [class*="pipe-step"], [class*="step-item"]')).filter(el => !el.matches('.pipeline-step-num, .pipeline-step-content, [class*="step-num"], [class*="step-content"]'));
        const flowEls = Array.from(c.querySelectorAll('.flow, [class*="flow"]')).filter(fl => fl.querySelector('.flow__node, [class*="node"]'));
        const fItemEls = Array.from(c.querySelectorAll('.f-item, [class*="f-item"]')).filter(fi => fi !== c);

        const rawSubCardCandidates = Array.from(c.querySelectorAll('div')).filter(sc => {
            if (sc === c || sc.querySelector('canvas, table, svg')) return false;
            if (pipeStepEls.some(ps => ps === sc || ps.contains(sc) || sc.contains(ps))) return false;
            if (flowEls.some(fl => fl === sc || fl.contains(sc) || sc.contains(fl))) return false;
            if (fItemEls.some(fi => fi === sc || fi.contains(sc))) return false;
            const scCls = typeof sc.className === 'string' ? sc.className : '';
            const isExplicitSub = /\\b(sub-card|serving-box|oss-card|template-box|nbar)\\b/.test(scCls) || scCls.includes('sub-card');
            if (isExplicitSub) return true;
            const scCs = window.getComputedStyle(sc);
            const scBg = extractBgColor(scCs);
            const scGrad = parseLinearGradient(scCs.backgroundImage);
            const scBr = parseFloat(scCs.borderRadius) || parseFloat(scCs.borderTopLeftRadius) || 0;
            const scR = sc.getBoundingClientRect();
            const txt = sc.innerText ? sc.innerText.trim() : '';
            if ((scGrad || (scBg && scBg !== 'rgba(0, 0, 0, 0)' && scBg !== 'transparent' && scBg !== st.backgroundColor)) && scBr >= 4 && scR.width > 90 && scR.height >= 28 && (txt.length > 8 || scGrad)) {
                return true;
            }
            return false;
        });
        const subCardEls = rawSubCardCandidates.filter(sc => !rawSubCardCandidates.some(p => p !== sc && p.contains(sc)));

        const positionedLabelEls = [];
        const subCards = subCardEls.map(sc => {
            const scSt = getStyles(sc);
            const absChildren = Array.from(sc.children).filter(ch => {
                const chCs = window.getComputedStyle(ch);
                return chCs.position === 'absolute' && (ch.innerText || '').trim().length > 0;
            });
            if (absChildren.length > 0 && scSt.gradient) {
                absChildren.forEach(ch => positionedLabelEls.push(ch));
                return {
                    rect: toIn(sc.getBoundingClientRect()),
                    styles: scSt,
                    head: '',
                    headStyles: scSt,
                    badge: '',
                    items: []
                };
            }

            const scBadge = sc.querySelector('.badge, [class*="badge"]');
            let scHead = sc.querySelector('.sub-card__head, h4, h5, [class*="font-bold"], b, strong');
            if (scHead && (scHead.tagName.toLowerCase() === 'b' || scHead.tagName.toLowerCase() === 'strong') && scHead.closest('p, li')) {
                scHead = null;
            }
            if (!scHead) {
                const firstChildP = sc.querySelector('p, div');
                if (firstChildP && firstChildP !== sc && !firstChildP.querySelector('p, div')) {
                    const fcCs = window.getComputedStyle(firstChildP);
                    if (parseInt(fcCs.fontWeight) >= 700 || fcCs.fontWeight === 'bold') {
                        scHead = firstChildP;
                    }
                }
            }
            const headTxt = scHead ? scHead.innerText.replace(scBadge ? scBadge.innerText : '', '').trim() : '';
            const headStyles = scHead ? getStyles(scHead) : scSt;
            const scList = [];
            sc.querySelectorAll('li, p, div').forEach(el => {
                if (el === scHead || (scHead && (scHead.contains(el) || el.contains(scHead)))) return;
                if (scBadge && (el === scBadge || scBadge.contains(el) || el.contains(scBadge))) return;
                if (el.querySelector('li, p, div')) return;
                const t = el.innerText ? el.innerText.trim() : '';
                if (t && t !== headTxt && !scList.includes(t)) {
                    scList.push(t);
                }
            });
            if (scList.length === 0) {
                let remTxt = sc.innerText ? sc.innerText.trim() : '';
                if (headTxt && remTxt.startsWith(headTxt)) {
                    remTxt = remTxt.slice(headTxt.length).trim();
                }
                if (remTxt) scList.push(remTxt);
            }
            return {
                rect: toIn(sc.getBoundingClientRect()),
                styles: scSt,
                head: headTxt,
                headStyles: headStyles,
                badge: scBadge ? scBadge.innerText.trim() : '',
                items: scList
            };
        });

        let cHeading = c.querySelector('h2, h3, h4, h5, .panel__title, .card-title, .pillar-title, .runtime__name, [class*="title"]');
        if (cHeading && subCardEls.some(sc => sc.contains(cHeading))) {
            cHeading = null;
        }
        let tagEl = c.querySelector('.panel__tag, [class*="panel__tag"], span[class*="tag"], .ma-pillar__num, .date, .no');
        if (tagEl && subCardEls.some(sc => sc.contains(tagEl))) {
            tagEl = null;
        }

        if (!tagEl && cHeading) {
            const hTop = cHeading.getBoundingClientRect().top;
            const eyebrowCandidates = Array.from(c.querySelectorAll('div, span')).filter(el => {
                if (el === cHeading || el.contains(cHeading) || cHeading.contains(el)) return false;
                if (subCardEls.some(sc => sc.contains(el))) return false;
                if (el.children.length > 1) return false;
                const er = el.getBoundingClientRect();
                const txt = el.innerText ? el.innerText.trim() : '';
                if (!txt || txt.length > 32 || txt.length <= 1) return false;
                const cs = window.getComputedStyle(el);
                return er.top <= hTop && er.height < 42 && parseFloat(cs.fontSize) <= 24 && parseInt(cs.fontWeight) >= 600;
            });
            if (eyebrowCandidates.length > 0) {
                tagEl = eyebrowCandidates[0];
            }
        }

        if (!cHeading && !hasSvgOrImg) {
            const boldCandidates = Array.from(c.querySelectorAll('div, span, b, strong')).filter(el => {
                if (el === tagEl || (tagEl && (tagEl.contains(el) || el.contains(tagEl)))) return false;
                if (subCardEls.some(sc => sc.contains(el) || el.contains(sc))) return false;
                if (el.closest('li, p, table, pre, code')) return false;
                if (el.querySelector('div, p, ul, table, canvas, svg')) return false;
                const er = el.getBoundingClientRect();
                if ((er.top - cDomRect.top) > Math.min(110, cDomRect.height * 0.35)) return false;
                const txt = el.innerText ? el.innerText.trim() : '';
                if (!txt || txt.length <= 1 || txt.length > 65) return false;
                const cs = window.getComputedStyle(el);
                return parseInt(cs.fontWeight) >= 700 || el.tagName.toLowerCase() === 'b' || el.tagName.toLowerCase() === 'strong';
            });
            if (boldCandidates.length > 0) {
                cHeading = boldCandidates[0];
            }
        }

        const tagData = tagEl ? {
            text: tagEl.innerText.trim(),
            rect: toIn(tagEl.getBoundingClientRect()),
            styles: getStyles(tagEl)
        } : null;

        const cardTitle = cHeading ? cHeading.innerText.trim() : '';
        const titleRect = cHeading ? toIn(cHeading.getBoundingClientRect()) : null;
        const titleStyles = cHeading ? getStyles(cHeading) : null;
        const titleRuns = cHeading ? extractRichRuns(cHeading, st.backgroundColor) : null;

        const cDescEl = c.querySelector('.panel__desc, .card-desc, .runtime__role, p.one');
        const cardDesc = (cDescEl && (!subCardEls.some(sc => sc.contains(cDescEl)))) ? cDescEl.innerText.trim() : '';
        const cardDescObj = (cDescEl && cardDesc) ? {
            text: cardDesc,
            rect: toIn(cDescEl.getBoundingClientRect()),
            styles: getStyles(cDescEl),
            runs: extractRichRuns(cDescEl, st.backgroundColor)
        } : null;

        const pipeSteps = pipeStepEls.map(stEl => {
            const numEl = stEl.querySelector('.pipeline-step-num, [class*="step-num"]');
            const h = stEl.querySelector('h3, h4, h5, [class*="title"], strong');
            const p = stEl.querySelector('p, span:last-child');
            return {
                rect: toIn(stEl.getBoundingClientRect()),
                styles: getStyles(stEl),
                stepNum: stEl.getAttribute('data-step') || (numEl ? numEl.innerText.trim() : ''),
                title: h ? h.innerText.trim() : '',
                titleStyles: h ? getStyles(h) : null,
                desc: p && p !== h ? p.innerText.trim() : '',
                descStyles: p && p !== h ? getStyles(p) : null
            };
        });

        const flows = flowEls.map(fl => {
            const nodes = Array.from(fl.querySelectorAll('.flow__node, [class*="node"], .flow__arrow, [class*="arrow"]')).map(n => {
                const nCls = typeof n.className === 'string' ? n.className : '';
                const isArrow = n.matches('.flow__arrow, [class*="arrow"]');
                return {
                    text: n.innerText.trim(),
                    rect: toIn(n.getBoundingClientRect()),
                    styles: getStyles(n),
                    isArrow: isArrow,
                    isHighlight: /hl|active|primary|accent/i.test(nCls)
                };
            });
            const arrowEl = fl.querySelector('.flow__arrow, [class*="arrow"]');
            return {
                rect: toIn(fl.getBoundingClientRect()),
                styles: getStyles(fl),
                nodes: nodes,
                arrow: arrowEl ? arrowEl.innerText.trim() : '→'
            };
        }).filter(f => f.nodes.length > 0);

        const fItems = fItemEls.map(fi => {
            const ic = fi.querySelector('.f-item__icon, [class*="icon"]');
            const tx = fi.querySelector('div:last-child, span:last-child, p');
            return {
                rect: toIn(fi.getBoundingClientRect()),
                styles: getStyles(fi),
                icon: ic ? ic.innerText.trim() : '',
                iconRect: ic ? toIn(ic.getBoundingClientRect()) : null,
                iconStyles: ic ? getStyles(ic) : null,
                text: tx && tx !== ic ? tx.innerText.trim() : fi.innerText.replace(ic ? ic.innerText : '', '').trim(),
                textRect: tx && tx !== ic ? toIn(tx.getBoundingClientRect()) : null,
                textStyles: tx && tx !== ic ? getStyles(tx) : null
            };
        });

        const codeBlockEls = [];
        const codeBlocks = [];
        c.querySelectorAll('pre, .file, .code-block, [class*="code-block"]').forEach(cb => {
            if (cb.tagName.toLowerCase() === 'code' && cb.closest('li, p')) return;
            if (subCardEls.some(sc => sc.contains(cb) || cb.contains(sc))) return;
            const txt = cb.innerText.trim();
            if (txt) {
                const cbSt = getStyles(cb);
                codeBlockEls.push(cb);
                codeBlocks.push({
                    text: txt,
                    rect: toIn(cb.getBoundingClientRect()),
                    styles: cbSt,
                    runs: extractRichRuns(cb, cbSt.backgroundColor)
                });
            }
        });

        const badgeEls = Array.from(c.querySelectorAll('span.badge, [class*="badge"], [class*="tag"], [class*="lang"], span')).filter(b => {
            if (b === tagEl || (tagEl && (tagEl.contains(b) || b.contains(tagEl)))) return false;
            if (cHeading && (b === cHeading || b.contains(cHeading) || cHeading.contains(b))) return false;
            if (cDescEl && (b === cDescEl || b.contains(cDescEl) || cDescEl.contains(b))) return false;
            if (subCardEls.some(sc => sc.contains(b))) return false;
            if (flowEls.some(fl => fl.contains(b))) return false;
            if (fItemEls.some(fi => fi.contains(b))) return false;
            if (codeBlockEls.some(cb => cb === b || cb.contains(b) || b.contains(cb))) return false;
            if (isAttachCard) return false;
            const txt = (b.innerText || '').trim();
            if (!txt || txt.length > 40) return false;
            const cls = typeof b.className === 'string' ? b.className : '';
            if (/badge|tag|lang/i.test(cls)) return true;
            const bCs = window.getComputedStyle(b);
            if (isCodeElement(b, bCs)) return false;
            const bBg = extractBgColor(bCs);
            const hasPillBg = bBg && bBg !== 'rgba(0, 0, 0, 0)' && bBg !== 'transparent' && bBg !== st.backgroundColor && bBg !== slideBg;
            const bBr = parseFloat(bCs.borderRadius) || parseFloat(bCs.borderTopLeftRadius) || 0;
            const bR = b.getBoundingClientRect();
            return Boolean(hasPillBg && bBr >= 4 && bR.width >= 20 && bR.height >= 20);
        });
        const badges = badgeEls.map(b => ({
            text: b.innerText.trim(),
            rect: toIn(b.getBoundingClientRect()),
            styles: getStyles(b)
        }));

        const paras = [];
        const seenTexts = new Set();
        if (cardTitle) seenTexts.add(cardTitle);
        if (cardDesc) seenTexts.add(cardDesc);
        if (tagData) seenTexts.add(tagData.text);
        badges.forEach(b => seenTexts.add(b.text));

        positionedLabelEls.forEach(pl => {
            const pTxt = (pl.innerText || '').trim();
            if (!pTxt) return;
            const plCs = window.getComputedStyle(pl);
            const plSt = getStyles(pl);
            if (plCs.transform && plCs.transform !== 'none') {
                plSt.textAlign = 'center';
            }
            paras.push({
                text: pTxt,
                isList: false,
                hasCustomMarker: true,
                rect: toIn(pl.getBoundingClientRect()),
                styles: plSt,
                runs: null
            });
        });

        if (fItems.length === 0 && pipeSteps.length === 0 && !isBridge && !isAttachCard) {
            const leafBlockEls = [];
            c.querySelectorAll('p, li, div, span').forEach(el => {
                if (el === c || el === cHeading || el === cDescEl || el === tagEl) return;
                if (cHeading && (cHeading.contains(el) || el.contains(cHeading))) return;
                if (cDescEl && (cDescEl.contains(el) || el.contains(cDescEl))) return;
                if (tagEl && (tagEl.contains(el) || el.contains(tagEl))) return;
                if (badgeEls.some(b => b === el || b.contains(el))) return;
                if (codeBlockEls.some(cb => cb === el || cb.contains(el) || el.contains(cb))) return;
                if (subCardEls.some(sc => sc === el || sc.contains(el) || el.contains(sc))) return;
                if (flowEls.some(fl => fl.contains(el) || el.contains(fl))) return;
                if (fItemEls.some(fi => fi.contains(el) || el.contains(fi))) return;
                if (pipeStepEls.some(ps => ps.contains(el) || el.contains(ps))) return;

                let targetEl = el;
                let hasChildBadge = false;
                if (el.tagName.toLowerCase() === 'li') {
                    const childBadge = badgeEls.find(b => el.contains(b));
                    if (childBadge) {
                        hasChildBadge = true;
                        const textSib = Array.from(el.children).find(ch => ch !== childBadge && !ch.contains(childBadge));
                        if (textSib) targetEl = textSib;
                    }
                } else if (badgeEls.some(b => el.contains(b))) {
                    return;
                }

                if (targetEl.querySelector('p, li, div, ul, ol, table, pre, canvas, svg')) return;
                if (targetEl.tagName.toLowerCase() === 'span' && !hasChildBadge) {
                    if (leafBlockEls.some(lb => lb.contains(targetEl))) return;
                    if (targetEl.querySelector('span')) return;
                }
                const rawTxt = targetEl.innerText ? targetEl.innerText.trim().replace(/\\s*\\n\\s*/g, '  ') : '';
                if (!rawTxt || seenTexts.has(rawTxt)) return;
                if (cardTitle && (cardTitle.includes(rawTxt) || rawTxt.includes(cardTitle)) && rawTxt.length <= cardTitle.length + 4) return;
                seenTexts.add(rawTxt);
                leafBlockEls.push(el);

                const elSt = getStyles(targetEl);
                let pRect = toIn(targetEl.getBoundingClientRect());
                let hasCustomMarker = hasChildBadge;
                if (el.tagName.toLowerCase() === 'li' && !hasChildBadge) {
                    const bcs = window.getComputedStyle(el, '::before');
                    const padL = parseFloat(window.getComputedStyle(el).paddingLeft) || 0;
                    if (padL >= 16 && bcs && bcs.content !== 'none' && bcs.content !== 'normal') {
                        const shiftIn = Number(((padL / slideW) * 13.333).toFixed(3));
                        pRect = {
                            left: Number((pRect.left + shiftIn).toFixed(3)),
                            top: pRect.top,
                            width: Number(Math.max(0.4, pRect.width - shiftIn).toFixed(3)),
                            height: pRect.height,
                            pxWidth: Math.max(20, pRect.pxWidth - Math.round(padL)),
                            pxHeight: pRect.pxHeight
                        };
                        hasCustomMarker = true;
                    }
                }

                paras.push({
                    text: rawTxt,
                    isList: el.tagName.toLowerCase() === 'li',
                    hasCustomMarker: hasCustomMarker,
                    rect: pRect,
                    styles: elSt,
                    runs: extractRichRuns(targetEl, st.backgroundColor)
                });
            });

            if (paras.length === 0 && subCards.length === 0 && codeBlocks.length === 0) {
                let directTxt = c.innerText ? c.innerText.trim() : '';
                if (cardTitle && directTxt.startsWith(cardTitle)) {
                    directTxt = directTxt.slice(cardTitle.length).trim();
                }
                if (tagData && directTxt.startsWith(tagData.text)) {
                    directTxt = directTxt.slice(tagData.text.length).trim();
                }
                directTxt = directTxt.replace(/\\s*\\n\\s*/g, '  ').trim();
                if (directTxt && !seenTexts.has(directTxt)) {
                    let fbRect = r;
                    if (titleRect) {
                        const newTop = Number((titleRect.top + titleRect.height + 0.04).toFixed(3));
                        fbRect = {
                            left: titleRect.left,
                            top: newTop,
                            width: titleRect.width,
                            height: Number(Math.max(0.25, (r.top + r.height) - newTop - 0.08).toFixed(3)),
                            pxWidth: titleRect.pxWidth,
                            pxHeight: Math.round(r.pxHeight * 0.6)
                        };
                    }
                    paras.push({
                        text: directTxt,
                        isList: false,
                        hasCustomMarker: false,
                        rect: fbRect,
                        styles: st
                    });
                }
            }
        }

        cards.push({
            rect: r,
            styles: st,
            isRounded: st.isRounded,
            borderRadius: st.borderRadius,
            hasTopAccent: st.borderTopWidth >= 2,
            topAccentColor: st.borderTopColor,
            hasLeftAccent: st.borderLeftWidth >= 3,
            leftAccentColor: st.borderLeftColor,
            hasCanvas: hasCanvas,
            hasSvgOrImg: hasSvgOrImg,
            isBridge: isBridge,
            bridge: bridgeData,
            isAttachCard: isAttachCard,
            attachData: attachData,
            tag: tagData,
            title: cardTitle,
            titleRect: titleRect,
            titleStyles: titleStyles,
            titleRuns: titleRuns,
            desc: cardDescObj,
            pipelineSteps: pipeSteps,
            subCards: subCards,
            flows: flows,
            fItems: fItems,
            badges: badges,
            codeBlocks: codeBlocks,
            paragraphs: paras
        });
    });

    const tables = [];
    s.querySelectorAll('table').forEach(tbl => {
        const parentCard = tbl.closest('.info-card, .card, .panel, [class*="rounded"]');
        const cardTitleEl = parentCard ? parentCard.querySelector('h3, h4, [class*="font-bold"], [class*="title"]') : null;
        const cardTitle = cardTitleEl ? cardTitleEl.innerText.trim() : '';
        const cardTitleColor = cardTitleEl ? getStyles(cardTitleEl).color : null;
        const tblDomRect = tbl.getBoundingClientRect();
        const tRect = toIn(tblDomRect);
        const tblSt = getStyles(tbl);

        const headers = [];
        const colWidthsPx = [];
        tbl.querySelectorAll('thead th, tr:first-child th').forEach(th => {
            const st = getStyles(th);
            colWidthsPx.push(th.getBoundingClientRect().width);
            const trEl = th.closest('tr');
            const theadEl = th.closest('thead');
            let thBg = st.backgroundColor;
            if (!thBg || thBg === 'transparent' || thBg === 'rgba(0, 0, 0, 0)') {
                thBg = trEl ? getStyles(trEl).backgroundColor : null;
            }
            if (!thBg || thBg === 'transparent' || thBg === 'rgba(0, 0, 0, 0)') {
                thBg = theadEl ? getStyles(theadEl).backgroundColor : null;
            }
            headers.push({
                text: th.innerText.trim(),
                color: st.color,
                bgColor: thBg && thBg !== 'transparent' && thBg !== 'rgba(0, 0, 0, 0)' ? thBg : null,
                fontWeight: st.fontWeight,
                fontSizePt: st.fontSizePt
            });
        });

        let rowBorderColor = null;
        const rows = [];
        tbl.querySelectorAll('tbody tr, tr:not(:first-child)').forEach(tr => {
            if (tr.querySelector('th') && !tr.querySelector('td') && tr.parentElement.tagName === 'THEAD') return;
            const trSt = getStyles(tr);
            const rowCells = [];
            const tds = Array.from(tr.querySelectorAll('td, th'));
            if (colWidthsPx.length === 0 && tds.length > 0) {
                tds.forEach(td => colWidthsPx.push(td.getBoundingClientRect().width));
            }
            tds.forEach(td => {
                if (!rowBorderColor) {
                    const tdCs = window.getComputedStyle(td);
                    if (parseFloat(tdCs.borderBottomWidth) > 0 && tdCs.borderBottomColor && tdCs.borderBottomColor !== 'transparent') {
                        rowBorderColor = tdCs.borderBottomColor;
                    }
                }
                const mb = td.querySelector('.math-badge, .score, [class*="badge"], [class*="tag"]');
                const st = getStyles(td);
                let cellBg = st.backgroundColor;
                if (!cellBg || cellBg === 'transparent' || cellBg === 'rgba(0, 0, 0, 0)') {
                    cellBg = trSt.backgroundColor;
                }
                if (cellBg === 'transparent' || cellBg === 'rgba(0, 0, 0, 0)') {
                    cellBg = null;
                }
                const effCellBg = cellBg || tblSt.backgroundColor || 'rgb(255, 255, 255)';
                const cellRuns = extractRichRuns(td, effCellBg);
                const onlyBadge = Boolean(mb && td.children.length === 1 && td.innerText.trim() === mb.innerText.trim());
                const mbSt = mb ? getStyles(mb) : null;
                rowCells.push({
                    text: td.innerText.trim(),
                    color: st.color,
                    bgColor: cellBg,
                    fontSizePt: st.fontSizePt,
                    is_bold: parseInt(st.fontWeight) >= 600 || st.fontWeight === 'bold',
                    has_badge: onlyBadge,
                    badge_text: onlyBadge ? mb.innerText.trim() : null,
                    badge_color: mbSt ? mbSt.color : null,
                    badge_bg: mbSt ? mbSt.backgroundColor : null,
                    runs: cellRuns
                });
            });
            if (rowCells.length) rows.push(rowCells);
        });

        const totalColW = colWidthsPx.reduce((a, b) => a + b, 0);
        const colRatios = totalColW > 0 ? colWidthsPx.map(w => Number((w / totalColW).toFixed(4))) : null;

        tables.push({
            rect: tRect,
            styles: tblSt,
            borderRadius: tblSt.borderRadius,
            hasShadow: tblSt.hasShadow,
            rowBorderColor: rowBorderColor,
            colRatios: colRatios,
            cardTitle: cardTitle,
            cardTitleColor: cardTitleColor,
            inCard: Boolean(parentCard || tblSt.borderRadius >= 8 || tblSt.hasShadow),
            tableIsCardItself: Boolean(!parentCard && (tblSt.borderRadius >= 8 || tblSt.hasShadow)),
            headers: headers,
            rows: rows,
            num_cols: Math.max(headers.length, ...rows.map(r => r.length), 0),
            num_rows: (headers.length ? 1 : 0) + rows.length
        });
    });

    const footnotes = [];
    const footnoteEls = [];
    const seenFn = new Set();
    s.querySelectorAll('.foot__n, .footnote, [class*="footnote"], [class*="bottom-note"]').forEach(el => {
        const txt = el.innerText ? el.innerText.trim() : '';
        if (!txt || seenFn.has(txt)) return;
        seenFn.add(txt);
        footnoteEls.push(el);
        footnotes.push({
            rect: toIn(el.getBoundingClientRect()),
            styles: getStyles(el),
            text: txt
        });
    });
    if (footnotes.length === 0) {
        s.querySelectorAll('p, div, span, small').forEach(el => {
            if (el === subDesc || el === mainH) return;
            if (el.querySelector('p, div, span, table')) return;
            if (topContainers.some(tc => tc.contains(el))) return;
            if (hlBoxEls.some(hb => hb.contains(el))) return;
            const txt = el.innerText ? el.innerText.trim() : '';
            if (!txt || seenFn.has(txt)) return;
            const r = toIn(el.getBoundingClientRect());
            const isFootnote = txt.startsWith('*') || txt.startsWith('※') || txt.startsWith('†') || 
                               txt.startsWith('1)') || txt.startsWith('참조') || txt.startsWith('출처') ||
                               (r.top > 6.4 && r.height < 0.55);
            if (isFootnote && r.top > 3.5) {
                seenFn.add(txt);
                footnoteEls.push(el);
                footnotes.push({
                    rect: r,
                    styles: getStyles(el),
                    text: txt
                });
            }
        });
    }

    const desc = [];
    const seenDesc = new Set();
    s.querySelectorAll('p.cover-desc, p.intro-desc, div.meta, div.arrow, span.x').forEach(el => {
        if (el === subDesc || el === mainH) return;
        if (topContainers.some(tc => tc === el || tc.contains(el))) return;
        if (hlBoxEls.some(hb => hb === el || hb.contains(el))) return;
        if (footnoteEls.some(fn => fn === el || fn.contains(el))) return;
        if (headerBadgeEls.some(hb => hb === el || hb.contains(el))) return;
        const txt = el.innerText ? el.innerText.trim() : '';
        if (txt && !seenDesc.has(txt)) {
            seenDesc.add(txt);
            desc.push({
                text: txt,
                rect: toIn(el.getBoundingClientRect()),
                styles: getStyles(el)
            });
        }
    });

    return {
        id: targetId,
        slideBgColor: slideBg,
        slideBgGradient: slideBgGradient,
        headerBadges: headerBadges,
        tag: headerBadges.length > 0 ? headerBadges[0] : null,
        num: num ? { text: num.innerText.trim(), rect: toIn(num.getBoundingClientRect()), styles: getStyles(num) } : null,
        title: mainH ? { text: mainH.innerText.trim(), rect: toIn(mainH.getBoundingClientRect()), styles: getStyles(mainH), runs: extractRichRuns(mainH, slideBg) } : null,
        sub: subDesc ? { text: subDesc.innerText.trim(), rect: toIn(subDesc.getBoundingClientRect()), styles: getStyles(subDesc), runs: extractRichRuns(subDesc, slideBg) } : null,
        desc: desc,
        decorShapes: decorShapes,
        images: images,
        cards: cards,
        charts: charts,
        tables: tables,
        highlightBoxes: hlBoxes,
        footnotes: footnotes
    };
})()
"""
