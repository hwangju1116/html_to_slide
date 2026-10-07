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
        s.style.opacity = '1';
        s.style.visibility = 'visible';
        s.style.display = 'flex';
        const sCs = window.getComputedStyle(s);
        const hasHiddenChild = Array.from(s.children).some(ch => window.getComputedStyle(ch).display === 'none');
        if (sCs.display === 'flex' && sCs.flexDirection === 'column' && sCs.justifyContent === 'space-between' && hasHiddenChild) {
            const hasCover = !!s.querySelector('.cover-slide-content');
            if (!hasCover) {
                s.style.justifyContent = 'flex-start';
                s.style.gap = '18px';
            }
        }
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

    const getStyles = (el) => {
        if (!el) return null;
        const cs = window.getComputedStyle(el);
        const pxSize = parseFloat(cs.fontSize) || 16;
        const scaledPt = Number(Math.max(8.5, Math.min(44, pxSize * 0.54)).toFixed(1));
        const br = parseFloat(cs.borderRadius) || parseFloat(cs.borderTopLeftRadius) || 0;
        const bw = parseFloat(cs.borderWidth) || parseFloat(cs.borderBottomWidth) || parseFloat(cs.borderLeftWidth) || 0;
        const btw = parseFloat(cs.borderTopWidth) || 0;
        const blw = parseFloat(cs.borderLeftWidth) || 0;
        const bbw = parseFloat(cs.borderBottomWidth) || 0;
        let color = cs.color;
        if ((color === 'rgba(0, 0, 0, 0)' || color === 'transparent') && cs.backgroundImage && cs.backgroundImage !== 'none') {
            const m = cs.backgroundImage.match(/rgba?\\(\\s*\\d+\\s*,\\s*\\d+\\s*,\\s*\\d+(?:\\s*,\\s*[\\d.]+)?\\s*\\)|#[0-9a-fA-F]{3,8}/);
            if (m) color = m[0];
        }
        return {
            color: color,
            backgroundColor: extractBgColor(cs),
            fontSizePt: scaledPt,
            fontWeight: cs.fontWeight,
            isBold: parseInt(cs.fontWeight) >= 600 || cs.fontWeight === 'bold',
            borderRadius: br,
            isRounded: br >= 4,
            borderTopColor: cs.borderTopColor,
            borderTopWidth: btw,
            borderLeftColor: cs.borderLeftColor,
            borderLeftWidth: blw,
            borderBottomColor: cs.borderBottomColor,
            borderBottomWidth: bbw,
            borderColor: bw > 0 ? (cs.borderColor || cs.borderBottomColor || cs.borderLeftColor) : null,
            borderWidth: bw,
            textAlign: cs.textAlign,
            display: cs.display
        };
    };

    const parseRgba = (str) => {
        if (!str || str === 'transparent' || str === 'none') return null;
        const m = str.match(/rgba?\\(\\s*(\\d+)[,\\s]+(\\d+)[,\\s]+(\\d+)(?:[,\\s/]+([\\d.]+))?\\s*\\)/);
        if (!m) return null;
        return {
            r: parseInt(m[1], 10),
            g: parseInt(m[2], 10),
            b: parseInt(m[3], 10),
            a: m[4] !== undefined ? parseFloat(m[4]) : 1.0
        };
    };

    const bgLayers = [];
    let curr = s;
    while (curr && curr !== document) {
        const cs = window.getComputedStyle(curr);
        const parsed = parseRgba(extractBgColor(cs));
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

    let headerDividerTop = null;
    let headerDividerColor = null;
    const slideHeaderEl = s.querySelector('.slide-header, .slide-head');
    if (slideHeaderEl) {
        const shCs = window.getComputedStyle(slideHeaderEl);
        if ((parseFloat(shCs.borderBottomWidth) || 0) >= 1) {
            const shRectIn = toIn(slideHeaderEl.getBoundingClientRect());
            headerDividerTop = Number((shRectIn.top + shRectIn.height).toFixed(3));
            headerDividerColor = shCs.borderBottomColor;
        }
    }

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

    let titleRuns = [];
    if (mainH && mainH.children.length > 0) {
        const mColor = getStyles(mainH).color;
        let hasColoredChild = false;
        mainH.querySelectorAll('span, b, strong, em').forEach(ch => {
            if (getStyles(ch).color !== mColor) hasColoredChild = true;
        });
        if (hasColoredChild) {
            mainH.childNodes.forEach(node => {
                if (node.nodeType === 3) {
                    const t = node.textContent;
                    if (t) titleRuns.push({ text: t, color: mColor });
                } else if (node.nodeType === 1) {
                    if (node.tagName.toLowerCase() === 'br') {
                        titleRuns.push({ text: '\\n', color: mColor });
                    } else {
                        const t = node.innerText || node.textContent || '';
                        if (t) titleRuns.push({ text: t, color: getStyles(node).color });
                    }
                }
            });
        }
    }

    let subDesc = Array.from(s.querySelectorAll('p.subtitle, p.premise, p.lead, p.cover-subtitle, h2.subtitle, [class*="subtitle"], [class*="premise"], p.text-slate-400')).find(el => {
        if (!mainH) return true;
        if (el === mainH || mainH.contains(el) || el.contains(mainH)) return false;
        if (el.closest('.glass-card, .card, .panel, .info-card, .timeline, .pipe, .steps, .pcol, table')) return false;
        const rIn = toIn(el.getBoundingClientRect());
        if (rIn.top > 4.2) return false;
        return el.innerText.trim().length > 0;
    }) || null;
    if (!subDesc && mainH) {
        const headWrap = mainH.closest('.slide-head');
        const cand = headWrap ? (headWrap.querySelector('p') || headWrap.nextElementSibling) : mainH.nextElementSibling;
        if (cand && cand !== mainH && cand.tagName.toLowerCase() === 'p') {
            const rIn = toIn(cand.getBoundingClientRect());
            const txt = cand.innerText ? cand.innerText.trim() : '';
            if (txt && rIn.top <= 3.2 && !txt.startsWith('※') && !txt.startsWith('*') && !txt.startsWith('출처')) {
                subDesc = cand;
            }
        }
    }

    const num = s.querySelector('.slide-number, [class*="slide-counter"], [class*="foot__n"]');

    const headerBadges = [];
    const headerBadgeEls = [];
    const seenBadges = new Set();
    const mainHTopIn = mainH ? toIn(mainH.getBoundingClientRect()).top : 1.0;
    const maxBadgeTopIn = Math.max(1.45, mainHTopIn + 0.05);
    const addHeaderBadge = (el) => {
        if (!el || el === mainH || el === subDesc || el === num) return;
        if (num && (num.contains(el) || el.contains(num))) return;
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

    const images = [];
    s.querySelectorAll('svg, img').forEach(el => {
        if (el.parentElement && el.parentElement.closest('svg')) return;
        const cs = window.getComputedStyle(el);
        if (cs.display === 'none' || cs.visibility === 'hidden' || parseFloat(cs.opacity) === 0) return;
        const r = el.getBoundingClientRect();
        if (r.width < 28 || r.height < 28) return;
        images.push({
            tagName: el.tagName.toLowerCase(),
            rect: toIn(r)
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
    const rawHlCandidates = Array.from(s.querySelectorAll('.foot__pocket, .highlight-box, .banner, [class*="highlight-box"], [class*="callout"]'));
    const topHlCandidates = rawHlCandidates.filter(hl => !rawHlCandidates.some(p => p !== hl && p.contains(hl)));
    topHlCandidates.forEach(hl => {
        const subBadgeEl = hl.querySelector('.banner-sub, [class*="banner-sub"], .badge');
        let subBadge = null;
        if (subBadgeEl && !subBadgeEl.closest('p, li')) {
            subBadge = {
                text: subBadgeEl.innerText.trim(),
                rect: toIn(subBadgeEl.getBoundingClientRect()),
                styles: getStyles(subBadgeEl)
            };
        }
        const mainTextEl = hl.querySelector('.banner-text, [class*="banner-text"]') || hl;
        let txt = hl.innerText ? hl.innerText.trim() : '';
        if (subBadge && subBadge.text && txt.endsWith(subBadge.text)) {
            txt = txt.slice(0, txt.length - subBadge.text.length).trim();
        }
        txt = txt.replace(/\\s*\\n\\s*/g, ' ');
        if (!txt || seenHl.has(txt)) return;
        seenHl.add(txt);
        hlBoxEls.push(hl);

        const hlParas = [];
        hl.querySelectorAll('p').forEach(pEl => {
            const pTxt = pEl.innerText ? pEl.innerText.trim().replace(/\\s*\\n\\s*/g, ' ') : '';
            if (pTxt) {
                hlParas.push({
                    text: pTxt,
                    styles: getStyles(pEl)
                });
            }
        });

        const hlSt = getStyles(hl);
        hlBoxes.push({
            rect: toIn(hl.getBoundingClientRect()),
            styles: hlSt,
            textStyles: getStyles(mainTextEl),
            text: txt,
            paragraphs: hlParas,
            hasLeftAccent: hlSt.borderLeftWidth >= 3,
            leftAccentColor: hlSt.borderLeftColor,
            subBadge: subBadge
        });
    });
    if (hlBoxes.length === 0) {
        s.querySelectorAll('*').forEach(hl => {
            if (hl.closest('table') || hl.querySelector('table')) return;
            if (hl.closest('.slide-header, .slide-head, header, footer, .foot, .controls-footer')) return;
            if (hl.querySelector('.slide-header, .slide-head, h1, h2')) return;
            if (hl === mainH || hl === subDesc || hl === num) return;
            if (headerBadgeEls.some(hb => hb === hl || hb.contains(hl) || hl.contains(hb))) return;
            if (hl.querySelector('.foot__pocket, .banner, [class*="highlight-box"], [class*="callout"]')) return;
            const rIn = toIn(hl.getBoundingClientRect());
            if (rIn.top < 1.8) return;
            const txt = hl.innerText ? hl.innerText.trim().replace(/\\s*\\n\\s*/g, ' ') : '';
            if (!txt || seenHl.has(txt)) return;
            const isCallout = txt.startsWith('💡') || txt.startsWith('📌') || txt.startsWith('★') || 
                              txt.startsWith('⚠️') || txt.startsWith('ℹ️') || txt.includes('핵심 메시지') ||
                              txt.includes('참고') || txt.includes('결론');
            if (isCallout && hl.offsetHeight < 320 && hl.offsetWidth > 150) {
                seenHl.add(txt);
                hlBoxEls.push(hl);
                const hlSt = getStyles(hl);
                hlBoxes.push({
                    rect: rIn,
                    styles: hlSt,
                    textStyles: hlSt,
                    text: txt,
                    paragraphs: [],
                    hasLeftAccent: hlSt.borderLeftWidth >= 3,
                    leftAccentColor: hlSt.borderLeftColor,
                    subBadge: null
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
        if (el.closest('.slide-head') || el.closest('.slide-header') || el.closest('header') || el.closest('.foot') || el.closest('footer')) return false;
        if (headerBadgeEls.some(hb => hb === el || hb.contains(el) || el.contains(hb))) return false;
        if (hlBoxEls.some(hb => hb === el || hb.contains(el))) return false;
        if (num && (el === num || num.contains(el) || el.contains(num))) return false;
        if (mainH && (el === mainH || el.contains(mainH))) return false;
        if (subDesc && (el === subDesc || el.contains(subDesc))) return false;

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

    const cards = [];
    topContainers.forEach(c => {
        if (c.querySelector('table') || c.closest('table')) return;
        const r = toIn(c.getBoundingClientRect());
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

        const rawPipeSteps = Array.from(c.querySelectorAll('.ma-pipe-step, .pipeline-step, [class*="pipe-step"], [class*="pipeline-step"], [class*="step-item"]')).filter(stEl => {
            const cls = typeof stEl.className === 'string' ? stEl.className : '';
            return !/\\b(pipe-step|pipeline-step|step-item)[-_]{1,2}(title|desc|num|icon|body|text|content)\\b/.test(cls);
        });
        const pipeStepEls = rawPipeSteps.filter(stEl => !rawPipeSteps.some(p => p !== stEl && p.contains(stEl)));
        const pipeSteps = pipeStepEls.map(stEl => {
            const numEl = stEl.querySelector('.pipeline-step-num, [class*="step-num"]');
            const h = stEl.querySelector('h3, h4, h5, [class*="title"], strong');
            const p = stEl.querySelector('p, span:last-child');
            return {
                rect: toIn(stEl.getBoundingClientRect()),
                styles: getStyles(stEl),
                stepNum: numEl ? numEl.innerText.trim() : (stEl.getAttribute('data-step') || ''),
                numRect: numEl ? toIn(numEl.getBoundingClientRect()) : null,
                numStyles: numEl ? getStyles(numEl) : null,
                title: h ? h.innerText.trim() : '',
                titleRect: h ? toIn(h.getBoundingClientRect()) : null,
                titleStyles: h ? getStyles(h) : null,
                desc: p && p !== h ? p.innerText.trim() : '',
                descRect: p && p !== h ? toIn(p.getBoundingClientRect()) : null,
                descStyles: p && p !== h ? getStyles(p) : null
            };
        });

        const rawFlows = Array.from(c.querySelectorAll('.flow, [class*="flow"]')).filter(fl => {
            const cls = typeof fl.className === 'string' ? fl.className : '';
            if (cls.includes('pipeline-flow')) return false;
            return !/\\bflow[-_]{1,2}(node|arrow|label|title|desc|text)\\b/.test(cls);
        });
        const flowEls = rawFlows.filter(fl => !rawFlows.some(p => p !== fl && p.contains(fl)));
        const flows = flowEls.map(fl => {
            const rawNodes = Array.from(fl.querySelectorAll('.flow__node, [class*="node"]'));
            const topNodes = rawNodes.filter(n => !rawNodes.some(p => p !== n && p.contains(n)));
            const nodes = topNodes.map(n => ({
                text: n.innerText.trim(),
                rect: toIn(n.getBoundingClientRect()),
                styles: getStyles(n)
            }));
            const arrowEl = fl.querySelector('.flow__arrow, [class*="arrow"]');
            return {
                rect: toIn(fl.getBoundingClientRect()),
                nodes: nodes,
                arrow: arrowEl ? arrowEl.innerText.trim() : '→'
            };
        }).filter(f => f.nodes.length > 0);

        const rawFItems = Array.from(c.querySelectorAll('.f-item, [class*="f-item"]')).filter(fi => {
            if (fi === c) return false;
            const cls = typeof fi.className === 'string' ? fi.className : '';
            if (/\\bf-item[-_]{1,2}(title|desc|icon|body|text|content)\\b/.test(cls)) return false;
            return true;
        });
        const fItemEls = rawFItems.filter(fi => !rawFItems.some(p => p !== fi && p.contains(fi)));
        const fItems = fItemEls.map(fi => {
            const ic = fi.querySelector('.f-item__icon, .icon, [class*="icon"]');
            const tEl = fi.querySelector('.f-item-title, .f-item__title, h4, h5');
            const dEl = fi.querySelector('.f-item-desc, .f-item__desc, p');
            const tx = fi.querySelector('div:last-child, span:last-child, p');
            const titleTxt = tEl ? tEl.innerText.trim() : '';
            const descTxt = dEl && dEl !== tEl ? dEl.innerText.trim() : '';
            let fullTxt = '';
            if (titleTxt && descTxt) {
                fullTxt = titleTxt + '\\n' + descTxt;
            } else if (tx && tx !== ic) {
                fullTxt = tx.innerText.trim();
            } else {
                fullTxt = fi.innerText.replace(ic ? ic.innerText : '', '').trim();
            }
            const textAnchorEl = tEl || (tx && tx !== ic ? tx : fi);
            return {
                rect: toIn(fi.getBoundingClientRect()),
                styles: getStyles(fi),
                icon: ic ? ic.innerText.trim() : '',
                iconRect: ic ? toIn(ic.getBoundingClientRect()) : null,
                iconStyles: ic ? getStyles(ic) : null,
                title: titleTxt,
                titleStyles: tEl ? getStyles(tEl) : null,
                desc: descTxt,
                descStyles: dEl ? getStyles(dEl) : null,
                text: fullTxt,
                textRect: toIn(textAnchorEl.getBoundingClientRect()),
                textStyles: getStyles(dEl || tEl || textAnchorEl)
            };
        });

        const rawSubCardCandidates = Array.from(c.querySelectorAll('div')).filter(sc => {
            if (sc === c || sc.querySelector('canvas, table, svg')) return false;
            if (fItemEls.some(fi => fi === sc || fi.contains(sc) || sc.contains(fi))) return false;
            if (pipeStepEls.some(ps => ps === sc || ps.contains(sc) || sc.contains(ps))) return false;
            if (flowEls.some(fl => fl === sc || fl.contains(sc) || sc.contains(fl))) return false;
            if (hlBoxEls.some(hb => hb === sc || hb.contains(sc) || sc.contains(hb))) return false;
            const scCls = typeof sc.className === 'string' ? sc.className : '';
            if (/\\b(badge|tag|pill|icon)\\b/.test(scCls) && !scCls.includes('card') && !scCls.includes('box')) return false;
            const isExplicitSub = /\\b(sub-card|serving-box|oss-card|template-box)\\b/.test(scCls) || scCls.includes('sub-card');
            if (isExplicitSub) return true;
            const scCs = window.getComputedStyle(sc);
            const scBg = extractBgColor(scCs);
            const scBr = parseFloat(scCs.borderRadius) || parseFloat(scCs.borderTopLeftRadius) || 0;
            const scR = sc.getBoundingClientRect();
            const txt = sc.innerText ? sc.innerText.trim() : '';
            if (scBg && scBg !== 'rgba(0, 0, 0, 0)' && scBg !== 'transparent' && scBg !== st.backgroundColor && scBr >= 4 && scR.width > 90 && scR.height >= 28 && txt.length > 8) {
                return true;
            }
            return false;
        });
        const subCardEls = rawSubCardCandidates.filter(sc => !rawSubCardCandidates.some(p => p !== sc && p.contains(sc)));

        const subCards = subCardEls.map(sc => {
            const scBadge = Array.from(sc.querySelectorAll('.badge, [class*="badge"]')).find(b => {
                const parentP = b.closest('p, li');
                return !parentP || parentP.innerText.trim() === b.innerText.trim();
            }) || null;
            let scHead = sc.querySelector('.sub-card__head, h3, h4, h5, [class*="font-bold"], [class*="head"], [class*="title"]');
            const scFullTxt = sc.innerText ? sc.innerText.trim() : '';
            if (!scHead && sc.firstElementChild) {
                const fc = sc.firstElementChild;
                const fcCs = window.getComputedStyle(fc);
                const fcTxt = fc.innerText ? fc.innerText.trim() : '';
                if (parseInt(fcCs.fontWeight) >= 600 && fcTxt && fcTxt.length <= 60 && scFullTxt.startsWith(fcTxt) && scFullTxt !== fcTxt && (fcCs.display !== 'inline' || sc.children.length >= 2)) {
                    scHead = fc;
                }
            }
            if (!scHead) {
                const bCand = sc.querySelector('b, strong');
                if (bCand) {
                    const bTxt = bCand.innerText ? bCand.innerText.trim() : '';
                    const pWrap = bCand.closest('p, li, div');
                    if (bTxt && scFullTxt.startsWith(bTxt) && (!pWrap || pWrap.innerText.trim() === bTxt)) {
                        scHead = bCand;
                    }
                }
            }
            const headTxt = scHead ? scHead.innerText.replace(scBadge ? scBadge.innerText : '', '').trim() : '';
            const headStyles = scHead ? getStyles(scHead) : getStyles(sc);
            const scList = [];
            const scListStyles = [];
            sc.querySelectorAll('li, p, div').forEach(el => {
                if (el === scHead || (scHead && (scHead.contains(el) || el.contains(scHead)))) return;
                if (scBadge && (el === scBadge || scBadge.contains(el))) return;
                if (el.querySelector('li, p, div')) return;
                const t = el.innerText ? el.innerText.trim() : '';
                if (t && t !== headTxt && !scList.includes(t)) {
                    scList.push(t);
                    scListStyles.push(getStyles(el));
                }
            });
            if (scList.length === 0) {
                let remTxt = sc.innerText ? sc.innerText.trim() : '';
                if (headTxt && remTxt.startsWith(headTxt)) {
                    remTxt = remTxt.slice(headTxt.length).trim();
                }
                if (remTxt) {
                    scList.push(remTxt);
                    scListStyles.push(getStyles(sc));
                }
            }
            const scSt = getStyles(sc);
            return {
                rect: toIn(sc.getBoundingClientRect()),
                styles: scSt,
                hasLeftAccent: scSt.borderLeftWidth >= 3,
                leftAccentColor: scSt.borderLeftColor,
                head: headTxt,
                headStyles: headStyles,
                badge: scBadge ? scBadge.innerText.trim() : '',
                items: scList,
                itemStyles: scListStyles
            };
        });

        let cHeading = c.querySelector('h2, h3, h4, h5, .panel__title, .card-title, .pillar-title, .runtime__name, [class*="title"]');
        if (cHeading && (
            subCardEls.some(sc => sc.contains(cHeading)) ||
            fItemEls.some(fi => fi.contains(cHeading)) ||
            pipeStepEls.some(ps => ps.contains(cHeading)) ||
            flowEls.some(fl => fl.contains(cHeading)) ||
            hlBoxEls.some(hb => hb.contains(cHeading))
        )) {
            cHeading = null;
        }
        let tagEl = c.querySelector('.panel__tag, [class*="panel__tag"], span[class*="tag"], .ma-pillar__num, .date, .no');
        if (tagEl && (
            subCardEls.some(sc => sc.contains(tagEl)) ||
            fItemEls.some(fi => fi.contains(tagEl)) ||
            pipeStepEls.some(ps => ps.contains(tagEl)) ||
            hlBoxEls.some(hb => hb.contains(tagEl)) ||
            tagEl.closest('p, li')
        )) {
            tagEl = null;
        }

        if (!tagEl && cHeading) {
            const hTop = cHeading.getBoundingClientRect().top;
            const eyebrowCandidates = Array.from(c.querySelectorAll('div, span')).filter(el => {
                if (el === cHeading || el.contains(cHeading) || cHeading.contains(el)) return false;
                if (subCardEls.some(sc => sc.contains(el))) return false;
                if (fItemEls.some(fi => fi.contains(el))) return false;
                if (pipeStepEls.some(ps => ps.contains(el))) return false;
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

        if (!cHeading) {
            const boldCandidates = Array.from(c.querySelectorAll('div, span, b, strong')).filter(el => {
                if (el === tagEl || (tagEl && (tagEl.contains(el) || el.contains(tagEl)))) return false;
                if (subCardEls.some(sc => sc.contains(el) || el.contains(sc))) return false;
                if (fItemEls.some(fi => fi.contains(el) || el.contains(fi))) return false;
                if (pipeStepEls.some(ps => ps.contains(el) || el.contains(ps))) return false;
                if (hlBoxEls.some(hb => hb.contains(el) || el.contains(hb))) return false;
                if (el.closest('li, p, table, pre, code')) return false;
                if (el.querySelector('div, p, ul, table, canvas, svg')) return false;
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

        let cardTitle = '';
        if (cHeading) {
            const hClone = cHeading.cloneNode(true);
            hClone.querySelectorAll('.badge, [class*="badge"], [class*="tag"], [class*="pill"]').forEach(b => b.remove());
            cardTitle = (hClone.innerText || hClone.textContent || cHeading.innerText || '').trim();
        }
        const titleRect = cHeading ? toIn(cHeading.getBoundingClientRect()) : null;
        const titleStyles = cHeading ? getStyles(cHeading) : null;

        const cDescEl = c.querySelector('.panel__desc, .card-desc, .runtime__role, p.one');
        const cardDesc = (cDescEl && (!subCardEls.some(sc => sc.contains(cDescEl))) && (!fItemEls.some(fi => fi.contains(cDescEl)))) ? cDescEl.innerText.trim() : '';
        const cardDescObj = (cDescEl && cardDesc) ? {
            text: cardDesc,
            rect: toIn(cDescEl.getBoundingClientRect()),
            styles: getStyles(cDescEl)
        } : null;

        const codeBlocks = [];
        c.querySelectorAll('pre, .file, .code-block').forEach(cb => {
            if (cb.closest('li, p')) return;
            if (subCardEls.some(sc => sc.contains(cb) || cb.contains(sc))) return;
            if (fItemEls.some(fi => fi.contains(cb) || cb.contains(fi))) return;
            const txt = cb.innerText.trim();
            if (txt) {
                codeBlocks.push({
                    text: txt,
                    rect: toIn(cb.getBoundingClientRect()),
                    styles: getStyles(cb)
                });
            }
        });

        const badgeEls = Array.from(c.querySelectorAll('span.badge, [class*="badge"], [class*="tag"], [class*="lang"]')).filter(b => {
            if (b === tagEl || (tagEl && tagEl.contains(b))) return false;
            if (cHeading && (b === cHeading || b.contains(cHeading))) return false;
            if (subCardEls.some(sc => sc.contains(b))) return false;
            if (fItemEls.some(fi => fi.contains(b))) return false;
            if (pipeStepEls.some(ps => ps.contains(b))) return false;
            if (hlBoxEls.some(hb => hb.contains(b))) return false;
            const parentP = b.closest('p, li');
            if (parentP && parentP.innerText.trim() !== b.innerText.trim()) return false;
            if (isAttachCard) return false;
            return b.innerText.trim().length > 0;
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

        if (fItems.length === 0 && pipeSteps.length === 0 && !isBridge && !isAttachCard) {
            const leafBlockEls = [];
            c.querySelectorAll('p, li, div, span').forEach(el => {
                if (el === c || el === cHeading || el === cDescEl || el === tagEl) return;
                if (cHeading && (cHeading.contains(el) || el.contains(cHeading))) return;
                if (cDescEl && (cDescEl.contains(el) || el.contains(cDescEl))) return;
                if (tagEl && (tagEl.contains(el) || el.contains(tagEl))) return;
                if (badgeEls.some(b => b === el || b.contains(el) || el.contains(b))) return;
                if (subCardEls.some(sc => sc === el || sc.contains(el) || el.contains(sc))) return;
                if (flowEls.some(fl => fl.contains(el) || el.contains(fl))) return;
                if (fItemEls.some(fi => fi.contains(el) || el.contains(fi))) return;
                if (pipeStepEls.some(ps => ps.contains(el) || el.contains(ps))) return;
                if (hlBoxEls.some(hb => hb === el || hb.contains(el) || el.contains(hb))) return;
                if (el.closest('pre, code')) return;
                if (codeBlocks.some(cb => cb.text === (el.innerText ? el.innerText.trim() : ''))) return;
                if (el.querySelector('p, li, div, ul, ol, table, pre, canvas, svg')) return;
                if (el.tagName.toLowerCase() === 'span') {
                    if (leafBlockEls.some(lb => lb.contains(el))) return;
                    if (el.querySelector('span')) return;
                }
                const hasBr = !!el.querySelector('br');
                const rawTxt = el.innerText ? (hasBr ? el.innerText.trim().replace(/[ \\t]*\\n[ \\t]*/g, '\\n') : el.innerText.trim().replace(/\\s*\\n\\s*/g, '  ')) : '';
                if (!rawTxt || seenTexts.has(rawTxt)) return;
                if (cardTitle && (cardTitle.includes(rawTxt) || rawTxt.includes(cardTitle)) && rawTxt.length <= cardTitle.length + 4) return;
                seenTexts.add(rawTxt);
                leafBlockEls.push(el);
                paras.push({
                    text: rawTxt,
                    isList: el.tagName.toLowerCase() === 'li',
                    rect: toIn(el.getBoundingClientRect()),
                    styles: getStyles(el)
                });
            });

            if (paras.length === 0 && subCards.length === 0 && codeBlocks.length === 0 && hlBoxEls.length === 0) {
                let directTxt = c.innerText ? c.innerText.trim() : '';
                if (cardTitle && directTxt.startsWith(cardTitle)) {
                    directTxt = directTxt.slice(cardTitle.length).trim();
                }
                if (tagData && directTxt.startsWith(tagData.text)) {
                    directTxt = directTxt.slice(tagData.text.length).trim();
                }
                directTxt = directTxt.replace(/\\s*\\n\\s*/g, '  ').trim();
                if (directTxt && !seenTexts.has(directTxt)) {
                    paras.push({
                        text: directTxt,
                        isList: false,
                        rect: r,
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
        const cTitleEl = parentCard ? Array.from(parentCard.querySelectorAll('h3, h4, [class*="font-bold"], [class*="title"]')).find(el => !tbl.contains(el)) : null;
        const cardTitle = cTitleEl ? cTitleEl.innerText.trim() : '';
        const cardTitleRect = cTitleEl ? toIn(cTitleEl.getBoundingClientRect()) : null;
        const cardTitleStyles = cTitleEl ? getStyles(cTitleEl) : null;
        const cardRect = parentCard ? toIn(parentCard.getBoundingClientRect()) : null;
        const cardStyles = parentCard ? getStyles(parentCard) : null;
        const tRect = toIn(tbl.getBoundingClientRect());
        const headers = [];
        tbl.querySelectorAll('thead th, tr:first-child th').forEach(th => {
            const st = getStyles(th);
            let hBg = st.backgroundColor;
            if (!hBg || hBg === 'transparent' || hBg === 'rgba(0, 0, 0, 0)') {
                const trP = th.parentElement;
                if (trP) hBg = getStyles(trP).backgroundColor;
            }
            if (!hBg || hBg === 'transparent' || hBg === 'rgba(0, 0, 0, 0)') {
                const theadP = th.closest('thead');
                if (theadP) hBg = getStyles(theadP).backgroundColor;
            }
            if (hBg === 'transparent' || hBg === 'rgba(0, 0, 0, 0)') hBg = null;
            headers.push({
                text: th.innerText.trim(),
                color: st.color,
                bgColor: hBg,
                fontSizePt: st.fontSizePt,
                fontWeight: st.fontWeight,
                textAlign: st.textAlign,
                borderColor: st.borderColor
            });
        });
        const rows = [];
        tbl.querySelectorAll('tbody tr, tr:not(:first-child)').forEach(tr => {
            if (tr.querySelector('th') && !tr.querySelector('td') && tr.parentElement.tagName === 'THEAD') return;
            const trSt = getStyles(tr);
            const rowCells = [];
            tr.querySelectorAll('td, th').forEach(td => {
                const mb = td.querySelector('.math-badge, [class*="badge"]');
                const st = getStyles(td);
                let cellBg = st.backgroundColor;
                if (!cellBg || cellBg === 'transparent' || cellBg === 'rgba(0, 0, 0, 0)') {
                    cellBg = trSt.backgroundColor;
                }
                if (!cellBg || cellBg === 'transparent' || cellBg === 'rgba(0, 0, 0, 0)') {
                    const tbodyP = td.closest('tbody');
                    if (tbodyP) cellBg = getStyles(tbodyP).backgroundColor;
                }
                if (cellBg === 'transparent' || cellBg === 'rgba(0, 0, 0, 0)') {
                    cellBg = null;
                }
                const onlyBadge = !!(mb && td.innerText.trim() === mb.innerText.trim());
                const mbSt = mb ? getStyles(mb) : null;
                const runs = [];
                if (!onlyBadge && td.children.length > 0 && td.querySelector('b, strong, span')) {
                    td.childNodes.forEach(n => {
                        if (n.nodeType === 3) {
                            if (n.textContent) runs.push({ text: n.textContent, bold: st.isBold, color: st.color });
                        } else if (n.nodeType === 1) {
                            const nSt = getStyles(n);
                            const nTxt = n.innerText || n.textContent || '';
                            if (nTxt) runs.push({ text: nTxt, bold: nSt.isBold || n.tagName === 'B' || n.tagName === 'STRONG', color: nSt.color });
                        }
                    });
                }
                rowCells.push({
                    text: td.innerText.trim(),
                    color: st.color,
                    bgColor: cellBg,
                    fontSizePt: st.fontSizePt,
                    is_bold: parseInt(st.fontWeight) >= 600 || (td.children.length === 1 && (td.firstElementChild.tagName === 'B' || td.firstElementChild.tagName === 'STRONG') && td.innerText.trim() === td.firstElementChild.innerText.trim()),
                    has_badge: onlyBadge,
                    badge_text: onlyBadge && mb ? mb.innerText.trim() : null,
                    badge_color: onlyBadge && mbSt ? mbSt.color : null,
                    badge_bg: onlyBadge && mbSt ? mbSt.backgroundColor : null,
                    runs: runs.length > 1 ? runs : [],
                    borderColor: st.borderColor || trSt.borderColor
                });
            });
            if (rowCells.length) rows.push(rowCells);
        });
        tables.push({
            rect: tRect,
            cardRect: cardRect,
            cardStyles: cardStyles,
            cardTitle: cardTitle,
            cardTitleRect: cardTitleRect,
            cardTitleStyles: cardTitleStyles,
            inCard: !!parentCard,
            headers: headers,
            rows: rows,
            num_cols: Math.max(headers.length, ...rows.map(r => r.length), 0),
            num_rows: (headers.length ? 1 : 0) + rows.length
        });
    });

    const footnotes = [];
    const footnoteEls = [];
    const seenFn = new Set();
    let footerDividerTop = null;
    let footerDividerColor = null;
    const rawFootnotes = Array.from(s.querySelectorAll('.foot__n, .footnote, [class*="footnote"], [class*="bottom-note"]'));
    const topFootnotes = rawFootnotes.filter(el => !rawFootnotes.some(p => p !== el && p.contains(el)));
    topFootnotes.forEach(el => {
        if (el === num) return;
        const cs = window.getComputedStyle(el);
        const btw = parseFloat(cs.borderTopWidth) || 0;
        const elRectIn = toIn(el.getBoundingClientRect());
        if (btw > 0 && footerDividerTop === null) {
            footerDividerTop = elRectIn.top;
            footerDividerColor = cs.borderTopColor;
        }
        if (cs.display.includes('flex') && el.children.length >= 2) {
            footnoteEls.push(el);
            Array.from(el.children).forEach(ch => {
                const cTxt = ch.innerText ? ch.innerText.trim() : '';
                if (!cTxt || seenFn.has(cTxt)) return;
                seenFn.add(cTxt);
                footnotes.push({
                    rect: toIn(ch.getBoundingClientRect()),
                    styles: getStyles(ch),
                    text: cTxt
                });
            });
            return;
        }
        const txt = el.innerText ? el.innerText.trim() : '';
        if (!txt || seenFn.has(txt)) return;
        seenFn.add(txt);
        footnoteEls.push(el);
        footnotes.push({
            rect: elRectIn,
            styles: getStyles(el),
            text: txt
        });
    });
    if (footnotes.length === 0) {
        s.querySelectorAll('p, div, span, small').forEach(el => {
            if (el === subDesc || el === mainH || el === num) return;
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
    s.querySelectorAll('p.cover-desc, p.cover-description, [class*="cover-desc"], p.intro-desc, div.meta, div.arrow, span.x').forEach(el => {
        if (el === subDesc || el === mainH || el === num) return;
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
        headerDividerTop: headerDividerTop,
        headerDividerColor: headerDividerColor,
        headerBadges: headerBadges,
        tag: headerBadges.length > 0 ? headerBadges[0] : null,
        num: num ? { text: num.innerText.trim(), rect: toIn(num.getBoundingClientRect()), styles: getStyles(num) } : null,
        title: mainH ? { text: mainH.innerText.trim(), rect: toIn(mainH.getBoundingClientRect()), styles: getStyles(mainH), runs: titleRuns } : null,
        sub: subDesc ? { text: subDesc.innerText.trim(), rect: toIn(subDesc.getBoundingClientRect()), styles: getStyles(subDesc) } : null,
        desc: desc,
        images: images,
        cards: cards,
        charts: charts,
        tables: tables,
        highlightBoxes: hlBoxes,
        footerDividerTop: footerDividerTop,
        footerDividerColor: footerDividerColor,
        footnotes: footnotes
    };
})();
"""
