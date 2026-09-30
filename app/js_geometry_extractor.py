JS_SLIDE_GEOMETRY_EXTRACTOR = """
(() => {
    const targetId = '%SLIDE_ID%';
    const slideIdx0 = %SLIDE_IDX_0%;
    const s = document.getElementById(targetId) || document.querySelector('.slide[data-active="true"]') || document.querySelector('.slide.active') || document.querySelector('.slide');
    if (!s) return null;

    // 0. Dynamically invoke any slide switcher function & populate FontAwesome icon symbols for offline/PPTX compatibility
    try {
        for (const fnName of ['updateSlide', 'showSlide', 'goToSlide', 'setSlide', 'renderSlide']) {
            if (typeof window[fnName] === 'function') {
                window[fnName](slideIdx0);
                break;
            }
        }
    } catch (e) {}

    // Dynamically hide any floating/bottom UI control bars or drawers outside the active slide
    try {
        if (s.parentElement) {
            Array.from(s.parentElement.children).forEach(ch => {
                if (ch === s || ch.classList.contains('slide')) return;
                const r = ch.getBoundingClientRect();
                if (ch.querySelector('button') || r.top > 150) {
                    ch.style.display = 'none';
                }
            });
        }
    } catch (e) {}

    const FA_ICON_MAP = {
        'fa-microchip': '⚡',
        'fa-brain': '🧠',
        'fa-gears': '⚙️',
        'fa-cogs': '⚙️',
        'fa-gear': '⚙️',
        'fa-repeat': '🔄',
        'fa-rotate': '🔄',
        'fa-sync': '🔄',
        'fa-check': '✓',
        'fa-circle-check': '✓',
        'fa-triangle-exclamation': '⚠️',
        'fa-exclamation-triangle': '⚠️',
        'fa-shield-halved': '🛡️',
        'fa-shield': '🛡️',
        'fa-flag-usa': '🌐',
        'fa-globe': '🌐',
        'fa-chart-line': '📈',
        'fa-chart-bar': '📊',
        'fa-chart-pie': '📊',
        'fa-compass': '🧭',
        'fa-cube': '◼',
        'fa-heart': '♥',
        'fa-mobile-screen': '📱',
        'fa-mobile': '📱',
        'fa-gem': '◆',
        'fa-bolt': '⚡',
        'fa-rocket': '🚀',
        'fa-star': '★',
        'fa-lock': '🔒',
        'fa-user': '👤',
        'fa-users': '👥',
        'fa-lightbulb': '💡',
        'fa-arrow-right': '→',
        'fa-chevron-right': '›',
        'fa-chevron-left': '‹'
    };

    document.querySelectorAll('i[class*="fa-"]').forEach(iconEl => {
        if (iconEl.innerText && iconEl.innerText.trim()) return;
        const cls = iconEl.className || '';
        for (const [faKey, sym] of Object.entries(FA_ICON_MAP)) {
            if (cls.includes(faKey)) {
                iconEl.textContent = sym;
                iconEl.style.fontStyle = 'normal';
                break;
            }
        }
    });

    try {
        document.querySelectorAll('.slide').forEach(el => {
            if (el !== s) {
                el.classList.remove('active');
                el.style.display = 'none';
            }
        });
        s.classList.add('active');
        s.style.opacity = '1';
        s.style.visibility = 'visible';
        s.style.display = 'flex';
        s.style.transform = 'none';
    } catch (e) {}

    // Ensure any Chart.js canvases inside s are resized and rendered synchronously
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

    const sRect = s.getBoundingClientRect();
    const slideW = sRect.width > 100 ? sRect.width : 1920;
    const slideH = sRect.height > 100 ? sRect.height : 1080;
    const scaleX = 13.333333 / slideW;
    const scaleY = 7.5 / slideH;

    const toIn = (r) => {
        const relLeft = r.left - sRect.left;
        const relTop = r.top - sRect.top;
        const inLeft = Math.max(0, relLeft * scaleX);
        const inTop = Math.max(0, relTop * scaleY);
        const inWidth = Math.min(13.333333 - inLeft, Math.max(0, r.width * scaleX));
        const inHeight = Math.min(7.5 - inTop, Math.max(0, r.height * scaleY));
        return {
            left: Number(inLeft.toFixed(3)),
            top: Number(inTop.toFixed(3)),
            width: Number(inWidth.toFixed(3)),
            height: Number(inHeight.toFixed(3)),
        };
    };

    const parseRgba = (str) => {
        if (!str || str === 'transparent') return null;
        const m = str.match(/rgba?\\(\\s*(\\d+)\\s*,\\s*(\\d+)\\s*,\\s*(\\d+)(?:\\s*,\\s*([0-9.]+))?\\s*\\)/);
        if (!m) return null;
        return {
            r: parseInt(m[1], 10),
            g: parseInt(m[2], 10),
            b: parseInt(m[3], 10),
            a: m[4] !== undefined ? parseFloat(m[4]) : 1.0
        };
    };

    const getStyles = (el) => {
        const cs = window.getComputedStyle(el);
        const radius = parseFloat(cs.borderRadius) || 0;
        const borderWidth = parseFloat(cs.borderWidth) || 0;
        const borderTopWidth = parseFloat(cs.borderTopWidth) || 0;
        const borderLeftWidth = parseFloat(cs.borderLeftWidth) || 0;
        const borderBottomWidth = parseFloat(cs.borderBottomWidth) || 0;
        const borderRightWidth = parseFloat(cs.borderRightWidth) || 0;
        const fontSize = parseFloat(cs.fontSize) || 14;
        const fontWeight = parseInt(cs.fontWeight) || (cs.fontWeight === 'bold' ? 700 : 400);
        let color = cs.color;
        if (cs.webkitTextFillColor && cs.webkitTextFillColor === 'rgba(0, 0, 0, 0)' && cs.backgroundImage && cs.backgroundImage.includes('gradient')) {
            const rgbMatches = cs.backgroundImage.match(/rgb\\(\\s*\\d+\\s*,\\s*\\d+\\s*,\\s*\\d+\\s*\\)/g);
            if (rgbMatches && rgbMatches.length > 0) {
                color = rgbMatches[rgbMatches.length - 1];
            }
        }
        return {
            color: color,
            backgroundColor: cs.backgroundColor,
            fontSizePt: Number((fontSize * 0.75).toFixed(1)),
            fontWeight: fontWeight,
            isBold: fontWeight >= 600,
            borderRadius: radius,
            isRounded: radius >= 4,
            borderTopColor: cs.borderTopColor,
            borderTopWidth: borderTopWidth,
            borderLeftColor: cs.borderLeftColor,
            borderLeftWidth: borderLeftWidth,
            borderBottomColor: cs.borderBottomColor,
            borderBottomWidth: borderBottomWidth,
            borderRightWidth: borderRightWidth,
            borderColor: cs.borderColor,
            borderWidth: borderWidth,
            textAlign: cs.textAlign,
            display: cs.display,
            opacity: parseFloat(cs.opacity) ?? 1.0
        };
    };

    // Walk ancestor chain from documentElement/body down to s to compute accurate alpha-blended canvas background
    const ancestorChain = [];
    let currAnc = s;
    while (currAnc && currAnc !== document) {
        ancestorChain.unshift(currAnc);
        currAnc = currAnc.parentElement;
    }
    let bgRgb = null;
    for (const anc of ancestorChain) {
        const cs = window.getComputedStyle(anc);
        let rgba = parseRgba(cs.backgroundColor);
        if ((!rgba || rgba.a === 0) && cs.backgroundImage && cs.backgroundImage !== 'none') {
            const mBg = cs.backgroundImage.match(/rgb\\(\\s*(\\d+)\\s*,\\s*(\\d+)\\s*,\\s*(\\d+)\\s*\\)/g);
            if (mBg && mBg.length > 0) {
                rgba = parseRgba(mBg[mBg.length - 1]);
            }
        }
        if (rgba && rgba.a > 0) {
            if (!bgRgb || rgba.a >= 0.99) {
                bgRgb = { r: rgba.r, g: rgba.g, b: rgba.b };
            } else {
                bgRgb = {
                    r: Math.round(rgba.r * rgba.a + bgRgb.r * (1 - rgba.a)),
                    g: Math.round(rgba.g * rgba.a + bgRgb.g * (1 - rgba.a)),
                    b: Math.round(rgba.b * rgba.a + bgRgb.b * (1 - rgba.a))
                };
            }
        }
    }
    const slideBg = bgRgb ? `rgb(${bgRgb.r}, ${bgRgb.g}, ${bgRgb.b})` : 'rgb(255, 255, 255)';

    // Extract Universal Top Header if present on s.parentElement
    let topHeader = null;
    if (s.parentElement) {
        const siblings = Array.from(s.parentElement.children).filter(ch => {
            if (ch === s || ch.classList.contains('slide') || ch.querySelector('button')) return false;
            const cs = window.getComputedStyle(ch);
            if (cs.display === 'none' || cs.visibility === 'hidden' || parseFloat(cs.opacity) === 0) return false;
            const r = ch.getBoundingClientRect();
            return r.top < 120 && r.height > 10 && r.height < 90 && r.width > slideW * 0.5;
        });
        if (siblings.length > 0) {
            const thEl = siblings[0];
            const thRect = toIn(thEl.getBoundingClientRect());
            const thSt = getStyles(thEl);
            const items = [];
            thEl.querySelectorAll('span, div, p, strong').forEach(el => {
                if (el.children.length > 0) return;
                const txt = el.innerText ? el.innerText.trim() : '';
                if (!txt) return;
                const r = toIn(el.getBoundingClientRect());
                if (r.width > 0.05 && r.height > 0.05) {
                    items.push({
                        text: txt,
                        rect: r,
                        styles: getStyles(el)
                    });
                }
            });
            topHeader = {
                rect: thRect,
                styles: thSt,
                items: items
            };
        }
    }

    // Universal Heading & Title Discovery
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

    // Extract inline runs for mainH if it contains styled child spans (e.g. .gradient-text-lg)
    let titleRuns = [];
    if (mainH && mainH.childNodes.length > 0) {
        const mainSt = getStyles(mainH);
        const walkTitleNodes = (node, inheritedColor) => {
            if (node.nodeType === Node.TEXT_NODE) {
                let t = node.textContent.replace(/\\s+/g, ' ');
                if (titleRuns.length === 0 || titleRuns[titleRuns.length - 1].text.endsWith('\\n')) {
                    t = t.replace(/^\\s+/, '');
                }
                if (t) titleRuns.push({ text: t, color: inheritedColor });
            } else if (node.nodeType === Node.ELEMENT_NODE) {
                if (node.tagName.toLowerCase() === 'br') {
                    if (titleRuns.length > 0) {
                        titleRuns[titleRuns.length - 1].text = titleRuns[titleRuns.length - 1].text.replace(/\\s+$/, '');
                    }
                    titleRuns.push({ text: '\\n', color: inheritedColor });
                } else {
                    const st = getStyles(node);
                    Array.from(node.childNodes).forEach(ch => walkTitleNodes(ch, st.color || inheritedColor));
                }
            }
        };
        Array.from(mainH.childNodes).forEach(ch => walkTitleNodes(ch, mainSt.color));
    }

    // Subtitle / Premise Discovery
    let subDesc = Array.from(s.querySelectorAll('p.subtitle, p.premise, p.lead, p.cover-subtitle, h2.subtitle, [class*="subtitle"], [class*="premise"], p.text-slate-400')).find(
        el => el !== mainH && (!mainH || (!mainH.contains(el) && !el.contains(mainH)))
    ) || null;
    if (!subDesc && mainH) {
        let next = mainH.closest('.slide-head') ? mainH.closest('.slide-head').nextElementSibling : mainH.nextElementSibling;
        while (next) {
            if (next.tagName.toLowerCase() === 'p') {
                subDesc = next;
                break;
            }
            if (next.querySelector('p')) {
                subDesc = next.querySelector('p');
                break;
            }
            next = next.nextElementSibling;
        }
    }

    // Extract ALL Header Badges / Pills / Eyebrows (above or beside mainH in the header region)
    const headerBadges = [];
    const headerBadgeEls = new Set();
    const seenBadges = new Set();
    const mainHBottom = mainH ? mainH.getBoundingClientRect().bottom : 240;
    const subBottom = subDesc ? subDesc.getBoundingClientRect().bottom : mainHBottom;
    const headerZoneLimit = Math.min(300, Math.max(mainHBottom, subBottom) + 15);

    const addHeaderBadge = (el) => {
        if (!el || el === mainH || el === subDesc) return;
        if (mainH && (el.contains(mainH) || mainH.contains(el))) return;
        if (subDesc && (el.contains(subDesc) || subDesc.contains(el))) return;
        const txt = el.innerText ? el.innerText.trim().replace(/\\s+/g, ' ') : '';
        if (!txt || txt.length > 65 || seenBadges.has(txt)) return;
        const r = el.getBoundingClientRect();
        if (r.width < 20 || r.height < 12 || r.height > 60 || r.top > headerZoneLimit) return;
        seenBadges.add(txt);
        headerBadgeEls.add(el);
        headerBadges.push({
            text: txt,
            rect: toIn(r),
            styles: getStyles(el)
        });
    };

    s.querySelectorAll('.slide-head .eyebrow, .slide-head .badge-pill, .slide-head [class*="tag"], .slide-head [class*="badge"], .slide-head [class*="pill"], .slide-tag, [class*="eyebrow"], [class*="badge-"]').forEach(el => {
        if (el.getBoundingClientRect().top <= headerZoneLimit) addHeaderBadge(el);
    });

    // Also discover eyebrows right before mainH and top-right pills in the header row
    if (mainH) {
        let prev = mainH.previousElementSibling;
        while (prev) {
            addHeaderBadge(prev);
            prev = prev.previousElementSibling;
        }
        const headerRow = mainH.parentElement && mainH.parentElement.parentElement !== s ? mainH.parentElement.parentElement : mainH.parentElement;
        if (headerRow && headerRow !== s && headerRow.getBoundingClientRect().top <= headerZoneLimit) {
            headerRow.querySelectorAll('span, div').forEach(el => {
                if (el === mainH || el === subDesc || el.contains(mainH) || el.contains(subDesc)) return;
                if (Array.from(headerBadgeEls).some(hb => hb.contains(el) || el.contains(hb))) return;
                const cs = window.getComputedStyle(el);
                const hasBg = cs.backgroundColor !== 'rgba(0, 0, 0, 0)' && cs.backgroundColor !== 'transparent';
                const hasBorder = parseFloat(cs.borderWidth) > 0;
                const isLeafText = el.children.length === 0 || (el.children.length <= 2 && el.querySelector('i'));
                if ((hasBg || hasBorder || isLeafText) && el.getBoundingClientRect().top <= headerZoneLimit) {
                    addHeaderBadge(el);
                }
            });
        }
    }

    const num = s.querySelector('.slide-number, [class*="slide-counter"], [class*="foot__n"]');

    // Descriptions: ONLY non-subtitle cover/intro paragraphs to prevent duplicate rendering
    const desc = [];
    s.querySelectorAll('p.cover-desc, p.intro-desc').forEach(el => {
        if (el !== subDesc) {
            const txt = el.innerText.trim();
            if (txt) {
                desc.push({
                    text: txt,
                    rect: toIn(el.getBoundingClientRect()),
                    styles: getStyles(el)
                });
            }
        }
    });

    // Detect Slide Bottom Footer Bar (e.g. border-t row at bottom of slide) so it is NOT misclassified as a card
    const footerEls = new Set();
    const footnotes = [];
    const seenFn = new Set();
    let footerDividerTop = null;
    let footerDividerColor = null;

    Array.from(s.children).forEach(ch => {
        const r = ch.getBoundingClientRect();
        const inR = toIn(r);
        const st = getStyles(ch);
        const hasOnlyTopBorder = st.borderTopWidth > 0 && st.borderBottomWidth === 0 && st.borderLeftWidth === 0 && (st.backgroundColor === 'rgba(0, 0, 0, 0)' || st.backgroundColor === 'transparent');
        if (inR.top >= 6.2 && inR.height <= 0.85 && (hasOnlyTopBorder || ch.classList.contains('foot') || ch.tagName.toLowerCase() === 'footer')) {
            footerEls.add(ch);
            if (st.borderTopWidth > 0) {
                footerDividerTop = inR.top;
                footerDividerColor = st.borderTopColor;
            }
            const leafSpans = Array.from(ch.querySelectorAll('span, p, div')).filter(el => {
                if (el.querySelector('span, p, div')) return false;
                return el.innerText && el.innerText.trim().length > 0;
            });
            if (leafSpans.length > 0) {
                leafSpans.forEach(sp => {
                    const txt = sp.innerText.trim();
                    if (txt && !seenFn.has(txt)) {
                        seenFn.add(txt);
                        footnotes.push({
                            rect: toIn(sp.getBoundingClientRect()),
                            styles: getStyles(sp),
                            text: txt
                        });
                    }
                });
            } else if (ch.innerText && ch.innerText.trim()) {
                const txt = ch.innerText.trim();
                seenFn.add(txt);
                footnotes.push({
                    rect: inR,
                    styles: st,
                    text: txt
                });
            }
        }
    });

    // Extract Charts (<canvas> with Chart.js config + high-DPI transparent PNG dataUrl)
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
                const legendDisplay = opts.plugins?.legend?.display !== false;
                const legendPosition = opts.plugins?.legend?.position || 'top';
                const cutout = opts.cutout || null;
                const xDisplay = opts.scales?.x?.display !== false;
                const yDisplay = opts.scales?.y?.display !== false;
                chartCfg = {
                    type: cType,
                    indexAxis: indexAxis,
                    labels: labels,
                    datasets: datasets,
                    data: { labels: labels, datasets: datasets },
                    legendDisplay: legendDisplay,
                    legendPosition: legendPosition,
                    cutout: cutout,
                    xDisplay: xDisplay,
                    yDisplay: yDisplay
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

    // Universal Container Discovery: Panels, Columns, Cards, Pillars, Bridges
    const hasVisualSurface = (el) => {
        const cs = window.getComputedStyle(el);
        const hasBg = cs.backgroundColor !== 'rgba(0, 0, 0, 0)' && cs.backgroundColor !== 'transparent' && cs.backgroundColor !== slideBg;
        const hasBoxBorder = (parseFloat(cs.borderWidth) > 0) || (parseFloat(cs.borderTopWidth) > 0 && parseFloat(cs.borderBottomWidth) > 0) || (parseFloat(cs.borderLeftWidth) >= 2) || (parseFloat(cs.borderTopWidth) >= 2);
        const hasShadow = cs.boxShadow && cs.boxShadow !== 'none';
        return hasBg || hasBoxBorder || hasShadow;
    };

    const isTopContainer = (el) => {
        if (el === s || el.contains(s) || el.querySelector('table') || el.closest('table')) return false;
        if (el.closest('.slide-head') || el.closest('header') || el.closest('.foot') || el.closest('footer')) return false;
        if (Array.from(footerEls).some(fe => fe === el || fe.contains(el))) return false;
        if (Array.from(headerBadgeEls).some(hb => hb === el || hb.contains(el) || el.contains(hb))) return false;
        if (mainH && (el === mainH || el.contains(mainH) || mainH.contains(el))) return false;
        if (subDesc && (el === subDesc || el.contains(subDesc) || subDesc.contains(el))) return false;

        const cls = el.className || '';
        if (typeof cls !== 'string') return false;

        const r = el.getBoundingClientRect();
        if (r.width < 80 || r.height < 36 || r.width > slideW * 0.98 || r.height > slideH * 0.98) return false;

        const visual = hasVisualSurface(el);
        const hasCardClass = /(?<![\\w-])(panel|card|glass-card|info-card|pillar|runtime|sub-card|bridge-badge|bridge-col|f-item|serving-box|attach-card|oss-card|template-box)(?![\\w-])/i.test(cls);

        if (!visual && !hasCardClass) return false;

        // If this element has NO visual surface of its own and contains child elements with visual surfaces, do not treat wrapper as card
        if (!visual && Array.from(el.querySelectorAll('*')).some(ch => hasVisualSurface(ch) && ch.getBoundingClientRect().width >= 80)) {
            return false;
        }
        return true;
    };

    let allCandidateContainers = Array.from(s.querySelectorAll('*')).filter(isTopContainer);
    let topContainers = allCandidateContainers.filter(c => !allCandidateContainers.some(p => p !== c && p.contains(c) && isTopContainer(p)));
    if (topContainers.length === 0) topContainers = allCandidateContainers;

    // Also capture standalone section headings outside topContainers (e.g. "AI 가전 진화 4단계 로드맵" above the 4 roadmap cards on Slide 2)
    const standaloneTexts = [];
    s.querySelectorAll('div, p, h3, h4, span').forEach(el => {
        if (el === mainH || (mainH && (mainH.contains(el) || el.contains(mainH)))) return;
        if (el === subDesc || (subDesc && (subDesc.contains(el) || el.contains(subDesc)))) return;
        if (Array.from(headerBadgeEls).some(hb => hb === el || hb.contains(el) || el.contains(hb))) return;
        if (Array.from(footerEls).some(fe => fe === el || fe.contains(el) || el.contains(fe))) return;
        if (topContainers.some(tc => tc === el || tc.contains(el) || el.contains(tc))) return;
        if (el.closest('table') || el.querySelector('table')) return;
        if (el.children.length > 0) return;
        const txt = el.innerText ? el.innerText.trim() : '';
        if (!txt) return;
        const r = toIn(el.getBoundingClientRect());
        if (r.width > 0.2 && r.height > 0.08 && r.top > 1.2 && r.top < 6.4) {
            standaloneTexts.push({
                text: txt,
                rect: r,
                styles: getStyles(el)
            });
        }
    });

    const cards = [];
    topContainers.forEach(c => {
        if (c.querySelector('table')) return;
        const r = toIn(c.getBoundingClientRect());
        const st = getStyles(c);

        // Check Bridge connector
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

        // Check Attach Card specialized structure
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

        const tagEl = c.querySelector('.panel__tag, [class*="panel__tag"], span[class*="tag"], .ma-pillar__num');
        const tagData = tagEl ? {
            text: tagEl.innerText.trim(),
            rect: toIn(tagEl.getBoundingClientRect()),
            styles: getStyles(tagEl)
        } : null;

        let titleEl = c.querySelector('.panel__title, h2, h3, h4, [class*="title"], [class*="name"]');
        if (!titleEl && !isAttachCard) {
            const boldLeaves = Array.from(c.querySelectorAll('div, span, p')).filter(el => {
                if (el.children.length > 0) return false;
                const cs = window.getComputedStyle(el);
                const fw = parseInt(cs.fontWeight) || 400;
                const fs = parseFloat(cs.fontSize) || 0;
                return (fw >= 700 || fs >= 16) && el.innerText.trim().length > 0 && el.innerText.trim().length < 60;
            });
            if (boldLeaves.length > 0) titleEl = boldLeaves[0];
        }
        let cardTitle = titleEl ? titleEl.innerText.trim() : null;
        if (isAttachCard && attachData && attachData.target) {
            cardTitle = attachData.target;
        }
        const titleRect = titleEl ? toIn(titleEl.getBoundingClientRect()) : null;
        const titleStyles = titleEl ? getStyles(titleEl) : null;

        const descEl = c.querySelector('.panel__desc, [class*="desc"], .oss-card__pos');
        let cardDesc = descEl && descEl !== titleEl ? {
            text: descEl.innerText.trim(),
            rect: toIn(descEl.getBoundingClientRect()),
            styles: getStyles(descEl)
        } : null;
        if (isAttachCard && attachData && attachData.desc) {
            cardDesc = {
                text: attachData.desc,
                rect: r,
                styles: st
            };
        }

        // Pipeline steps
        const pipeStepEls = Array.from(c.querySelectorAll('.ma-pipe-step, [class*="pipe-step"], [class*="step-item"]'));
        const pipeSteps = pipeStepEls.map(stEl => {
            const h = stEl.querySelector('h3, h4, h5, [class*="title"], strong');
            const p = stEl.querySelector('p, span:last-child');
            return {
                title: h ? h.innerText.trim() : stEl.innerText.trim(),
                desc: p && p !== h ? p.innerText.trim() : '',
                rect: toIn(stEl.getBoundingClientRect()),
                styles: getStyles(stEl)
            };
        });

        // Sub-cards (both class-based and nested visual boxes inside c)
        const subCardEls = Array.from(c.querySelectorAll('.sub-card, .runtime-card, [class*="sub-card"], [class*="runtime-card"]')).filter(sc => {
            const cls = sc.className || '';
            if (typeof cls !== 'string') return false;
            return !cls.includes('__head') && !cls.includes('__badge') && !cls.includes('__mech') && !cls.includes('__target') && !cls.includes('__mode') && !cls.includes('-grid') && !cls.includes('-layout');
        });
        const subCards = subCardEls.map(sc => {
            const scHead = sc.querySelector('.sub-card__head, .runtime-card__head, [class*="head"], h4, h5');
            const scBadge = sc.querySelector('.runtime-card__badge, [class*="badge"], [class*="mode"]');
            const headTxt = scHead ? scHead.innerText.trim() : '';
            const badgeTxt = scBadge ? scBadge.innerText.trim() : '';
            const items = Array.from(sc.querySelectorAll('li, p')).filter(el => {
                if (scHead && (el === scHead || scHead.contains(el))) return false;
                if (scBadge && (el === scBadge || scBadge.contains(el))) return false;
                return true;
            }).map(el => el.innerText.trim()).filter(Boolean);
            return {
                rect: toIn(sc.getBoundingClientRect()),
                head: headTxt,
                badge: badgeTxt,
                items: items,
                styles: getStyles(sc)
            };
        });

        // Extract all inner visual boxes (nested cards, icon boxes, number badges, pill badges, color swatches)
        const innerBoxes = [];
        const innerBoxEls = new Set();
        c.querySelectorAll('div, span').forEach(ib => {
            if (ib === c || subCardEls.includes(ib) || pipeStepEls.includes(ib)) return;
            const cs = window.getComputedStyle(ib);
            const ibSt = getStyles(ib);
            const hasBg = cs.backgroundColor !== 'rgba(0, 0, 0, 0)' && cs.backgroundColor !== 'transparent' && cs.backgroundColor !== st.backgroundColor;
            const hasBorder = parseFloat(cs.borderWidth) > 0 && parseFloat(cs.borderTopWidth) > 0 && parseFloat(cs.borderBottomWidth) > 0;
            if (!hasBg && !hasBorder) return;
            const ibRect = ib.getBoundingClientRect();
            if (ibRect.width < 6 || ibRect.height < 6) return;
            innerBoxEls.add(ib);
            const isLeafBox = ib.children.length === 0 || (ib.children.length === 1 && ib.children[0].tagName.toLowerCase() === 'i');
            innerBoxes.push({
                rect: toIn(ibRect),
                styles: ibSt,
                hasBg: hasBg,
                hasBorder: hasBorder,
                hasLeftAccent: ibSt.borderLeftWidth >= 2 && ibSt.borderRightWidth === 0,
                leftAccentColor: ibSt.borderLeftColor,
                text: isLeafBox ? (ib.innerText ? ib.innerText.trim() : '') : ''
            });
        });

        // Extract internal divider lines (border-t or border-l inside card)
        const dividers = [];
        c.querySelectorAll('div').forEach(dv => {
            if (dv === c || innerBoxEls.has(dv)) return;
            const dvSt = getStyles(dv);
            const dvRect = toIn(dv.getBoundingClientRect());
            if (dvSt.borderTopWidth > 0 && dvSt.borderBottomWidth === 0 && dvRect.width > 0.5) {
                dividers.push({
                    orientation: 'horizontal',
                    rect: { left: dvRect.left, top: dvRect.top, width: dvRect.width, height: 0.01 },
                    color: dvSt.borderTopColor
                });
            }
            if (dvSt.borderLeftWidth > 0 && dvSt.borderRightWidth === 0 && dvRect.height > 0.25) {
                dividers.push({
                    orientation: 'vertical',
                    rect: { left: dvRect.left, top: dvRect.top, width: 0.01, height: dvRect.height },
                    color: dvSt.borderLeftColor
                });
            }
        });

        // Extract exact positioned leaf text blocks inside card for 1:1 coordinate rendering
        const textBlocks = [];
        const consumedLeafEls = new Set();
        innerBoxes.forEach(ib => {
            if (ib.text) {
                c.querySelectorAll('div, span, i').forEach(el => {
                    if (el.innerText && el.innerText.trim() === ib.text && Math.abs(toIn(el.getBoundingClientRect()).top - ib.rect.top) < 0.08) {
                        consumedLeafEls.add(el);
                    }
                });
            }
        });

        c.querySelectorAll('h1, h2, h3, h4, h5, h6, p, li, div, span').forEach(el => {
            if (el === c || consumedLeafEls.has(el)) return;
            if (Array.from(consumedLeafEls).some(ce => ce.contains(el))) return;
            if (innerBoxEls.has(el) && (el.children.length === 0 || (el.children.length === 1 && el.children[0].tagName.toLowerCase() === 'i'))) return;
            if (el.closest('canvas') || el.querySelector('canvas') || el.querySelector('table')) return;

            // Check if el is a leaf text block (no block children; only inline strong/b/span/i allowed if they sit on the same line)
            const hasBlockChildren = Array.from(el.children).some(ch => {
                const tag = ch.tagName.toLowerCase();
                if (['div', 'p', 'ul', 'ol', 'li', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'table', 'canvas'].includes(tag)) return true;
                const chCs = window.getComputedStyle(ch);
                const chHasBg = chCs.backgroundColor !== 'rgba(0, 0, 0, 0)' && chCs.backgroundColor !== 'transparent';
                if (chHasBg) return true;
                return false;
            });
            if (hasBlockChildren) return;

            // If el is a flex container with multiple distinct horizontal child spans (like left label + right value), let the child spans be extracted individually
            const elCs = window.getComputedStyle(el);
            if ((elCs.display === 'flex' || elCs.display === 'inline-flex') && el.children.length >= 2) {
                const childTags = Array.from(el.children).map(ch => ch.tagName.toLowerCase());
                if (!childTags.includes('i')) return;
            }

            const txt = el.innerText ? el.innerText.trim() : '';
            if (!txt) return;

            const elRect = toIn(el.getBoundingClientRect());
            if (elRect.width < 0.05 || elRect.height < 0.05) return;

            consumedLeafEls.add(el);
            el.querySelectorAll('*').forEach(descEl => consumedLeafEls.add(descEl));

            const elSt = getStyles(el);
            const tag = el.tagName.toLowerCase();
            const isListItem = tag === 'li';

            // Extract inline runs if strong/span children have distinct bold or color
            const runs = [];
            if (el.children.length > 0) {
                const walkInline = (node, inhBold, inhColor) => {
                    if (node.nodeType === Node.TEXT_NODE) {
                        const t = node.textContent.replace(/\\s+/g, ' ');
                        if (t) runs.push({ text: t, isBold: inhBold, color: inhColor });
                    } else if (node.nodeType === Node.ELEMENT_NODE) {
                        const nSt = getStyles(node);
                        const b = inhBold || nSt.isBold || ['strong', 'b'].includes(node.tagName.toLowerCase());
                        const col = nSt.color || inhColor;
                        Array.from(node.childNodes).forEach(ch => walkInline(ch, b, col));
                    }
                };
                Array.from(el.childNodes).forEach(ch => walkInline(ch, elSt.isBold, elSt.color));
            }

            // Detect right alignment if inside justify-between / text-right
            let align = elSt.textAlign;
            if (el.parentElement) {
                const pCs = window.getComputedStyle(el.parentElement);
                if (pCs.textAlign === 'right' || el.className.includes('text-right')) align = 'right';
                else if (pCs.textAlign === 'center' || el.className.includes('text-center')) align = 'center';
                else if (pCs.justifyContent === 'space-between' && el === el.parentElement.lastElementChild && el.parentElement.children.length >= 2 && elRect.left > r.left + r.width * 0.55) {
                    align = 'right';
                }
            }

            textBlocks.push({
                text: txt,
                rect: elRect,
                styles: { ...elSt, textAlign: align },
                isListItem: isListItem,
                runs: runs.length > 1 ? runs : null
            });
        });

        // Inline flows
        const flowEls = Array.from(c.querySelectorAll('.inline-flow, .flow, [class*="inline-flow"]'));
        const flows = flowEls.map(fl => {
            const nodes = Array.from(fl.children).map(ch => ({
                text: ch.innerText.trim(),
                isArrow: ch.innerText.includes('──►') || ch.innerText.includes('➔') || ch.innerText.includes('→'),
                isHighlight: (ch.className || '').includes('hl') || (ch.className || '').includes('node--hl'),
                rect: toIn(ch.getBoundingClientRect()),
                styles: getStyles(ch)
            }));
            return { rect: toIn(fl.getBoundingClientRect()), nodes: nodes };
        });

        // Feature items
        const fItemEls = Array.from(c.querySelectorAll('.f-item, [class*="f-item"]')).filter(fi => {
            const cls = fi.className || '';
            return typeof cls === 'string' && !cls.includes('__icon') && !cls.includes('__text') && !cls.includes('-list');
        });
        const fItems = fItemEls.map(fi => {
            const iconEl = fi.querySelector('.f-item__icon, [class*="icon"], span:first-child');
            const textEl = fi.querySelector('div, p') || (iconEl ? Array.from(fi.children).find(ch => ch !== iconEl) : fi);
            return {
                icon: iconEl ? iconEl.innerText.trim() : '',
                iconRect: iconEl ? toIn(iconEl.getBoundingClientRect()) : null,
                iconStyles: iconEl ? getStyles(iconEl) : null,
                text: textEl ? textEl.innerText.trim() : (iconEl ? fi.innerText.replace(iconEl.innerText, '').trim() : fi.innerText.trim()),
                textRect: textEl ? toIn(textEl.getBoundingClientRect()) : toIn(fi.getBoundingClientRect()),
                textStyles: textEl ? getStyles(textEl) : getStyles(fi),
                rect: toIn(fi.getBoundingClientRect())
            };
        });

        const seenTexts = new Set();
        if (cardTitle) seenTexts.add(cardTitle);
        if (tagData) seenTexts.add(tagData.text);
        if (cardDesc) seenTexts.add(cardDesc.text);
        if (attachData) {
            if (attachData.target) seenTexts.add(attachData.target);
            if (attachData.mode) seenTexts.add(attachData.mode);
            if (attachData.mech) seenTexts.add(attachData.mech);
            if (attachData.desc) seenTexts.add(attachData.desc);
            const ah = c.querySelector('.attach-card__head');
            if (ah) seenTexts.add(ah.innerText.trim());
        }
        pipeSteps.forEach(ps => { if (ps.title) seenTexts.add(ps.title); if (ps.desc) seenTexts.add(ps.desc); });
        subCards.forEach(sc => {
            if (sc.head) seenTexts.add(sc.head);
            if (sc.badge) seenTexts.add(sc.badge);
            sc.items.forEach(it => seenTexts.add(it));
        });
        fItems.forEach(fi => {
            if (fi.icon) seenTexts.add(fi.icon);
            if (fi.text) seenTexts.add(fi.text);
        });

        // Block-level code and example blocks
        const isBlockCode = (cb) => {
            if (cb.tagName.toLowerCase() === 'pre') return true;
            const cls = cb.className || '';
            if (typeof cls === 'string' && (cls.includes('__code') || cls.includes('__ex') || cls.includes('code-block') || cls.includes('spec') || cls.includes('template-box__code'))) return true;
            const cs = window.getComputedStyle(cb);
            const isMono = cs.fontFamily.includes('mono') || cls.includes('mono');
            const hasMultipleLines = cb.innerText.trim().includes('\\n') || cb.offsetHeight >= 36;
            return isMono && hasMultipleLines;
        };

        const codeBlocks = Array.from(c.querySelectorAll('pre, [class*="code"], [class*="spec"], [class*="__ex"]')).filter(cb => {
            if (cb === c || (titleEl && titleEl.contains(cb)) || (tagEl && tagEl.contains(cb))) return false;
            return isBlockCode(cb) && cb.innerText.trim().length > 0;
        }).map(cb => ({
            text: cb.innerText.trim(),
            rect: toIn(cb.getBoundingClientRect()),
            styles: getStyles(cb)
        }));
        codeBlocks.forEach(cb => seenTexts.add(cb.text));

        // Paragraphs / list items
        const paras = [];
        if (!isAttachCard) {
            const directLis = Array.from(c.querySelectorAll('li')).filter(li => !subCardEls.some(sc => sc.contains(li)));
            if (directLis.length > 0) {
                directLis.forEach(li => {
                    const txt = li.innerText.trim();
                    if (txt && !seenTexts.has(txt)) {
                        seenTexts.add(txt);
                        paras.push({ text: txt, rect: toIn(li.getBoundingClientRect()), styles: getStyles(li) });
                    }
                });
            }
            c.querySelectorAll('p, div, span').forEach(el => {
                if (el === titleEl || (titleEl && titleEl.contains(el))) return;
                if (descEl && (el === descEl || descEl.contains(el))) return;
                if (tagEl && (el === tagEl || tagEl.contains(el))) return;
                if (subCardEls.some(sc => sc.contains(el))) return;
                if (flowEls.some(fl => fl.contains(el))) return;
                if (fItemEls.some(fi => fi.contains(el))) return;
                if (pipeStepEls.some(ps => ps.contains(el))) return;
                if (codeBlocks.some(cb => cb.text === el.innerText.trim())) return;
                if (el.querySelector('p, div, table, pre, span')) return;
                const txt = el.innerText.trim();
                if (txt && !seenTexts.has(txt)) {
                    seenTexts.add(txt);
                    paras.push({ text: txt, rect: toIn(el.getBoundingClientRect()), styles: getStyles(el) });
                }
            });
        }

        const badges = Array.from(c.querySelectorAll('span.badge, [class*="badge"], [class*="tag"], [class*="lang"]')).filter(b => {
            if (b === tagEl || (tagEl && tagEl.contains(b))) return false;
            if (subCardEls.some(sc => sc.contains(b))) return false;
            if (isAttachCard) return false;
            return true;
        }).map(b => ({
            text: b.innerText.trim(),
            rect: toIn(b.getBoundingClientRect()),
            styles: getStyles(b)
        }));

        cards.push({
            rect: r,
            styles: st,
            isRounded: st.isRounded,
            borderRadius: st.borderRadius,
            hasTopAccent: st.borderTopWidth >= 2,
            topAccentColor: st.borderTopColor,
            hasLeftAccent: st.borderLeftWidth >= 2,
            leftAccentColor: st.borderLeftColor,
            isBridge: isBridge,
            bridge: bridgeData,
            isAttachCard: isAttachCard,
            attachData: attachData,
            tag: tagData,
            title: cardTitle,
            titleRect: titleRect,
            titleStyles: titleStyles,
            desc: cardDesc,
            pipelineSteps: pipeSteps,
            subCards: subCards,
            innerBoxes: innerBoxes,
            dividers: dividers,
            textBlocks: textBlocks,
            flows: flows,
            fItems: fItems,
            badges: badges,
            codeBlocks: codeBlocks,
            paragraphs: paras
        });
    });

    // Tables (with multi-line styled block extraction per cell)
    const tables = [];
    s.querySelectorAll('table').forEach(tbl => {
        const parentCard = tbl.closest('.info-card, .card, .panel, [class*="rounded"]');
        const cardTitle = parentCard && parentCard.querySelector('h3, h4, [class*="font-bold"], [class*="title"]') ? parentCard.querySelector('h3, h4, [class*="font-bold"], [class*="title"]').innerText.trim() : '';
        const tRect = toIn(tbl.getBoundingClientRect());
        const headers = [];
        tbl.querySelectorAll('thead th, tr:first-child th').forEach(th => {
            const st = getStyles(th);
            let hBg = st.backgroundColor;
            if ((!hBg || hBg === 'transparent' || hBg === 'rgba(0, 0, 0, 0)') && th.parentElement) {
                const trBg = getStyles(th.parentElement).backgroundColor;
                if (trBg && trBg !== 'transparent' && trBg !== 'rgba(0, 0, 0, 0)') hBg = trBg;
            }
            headers.push({
                text: th.innerText.trim(),
                color: st.color,
                bgColor: hBg !== 'transparent' && hBg !== 'rgba(0, 0, 0, 0)' ? hBg : null,
                fontWeight: st.fontWeight
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
                if (cellBg === 'transparent' || cellBg === 'rgba(0, 0, 0, 0)') {
                    cellBg = null;
                }
                const blockLines = [];
                if (td.querySelector('div, p, br')) {
                    td.childNodes.forEach(n => {
                        if (n.nodeType === Node.TEXT_NODE) {
                            const tTxt = n.textContent ? n.textContent.trim() : '';
                            if (tTxt) {
                                blockLines.push({
                                    text: tTxt,
                                    color: st.color,
                                    fontSizePt: st.fontSizePt,
                                    isBold: parseInt(st.fontWeight) >= 600
                                });
                            }
                        } else if (n.nodeType === Node.ELEMENT_NODE && n.tagName !== 'BR') {
                            const dTxt = n.innerText ? n.innerText.trim() : '';
                            if (dTxt) {
                                const dSt = getStyles(n);
                                blockLines.push({
                                    text: dTxt,
                                    color: dSt.color,
                                    fontSizePt: dSt.fontSizePt,
                                    isBold: dSt.isBold
                                });
                            }
                        }
                    });
                }
                rowCells.push({
                    text: td.innerText.trim(),
                    color: st.color,
                    bgColor: cellBg,
                    is_bold: parseInt(st.fontWeight) >= 600,
                    has_badge: !!mb && blockLines.length <= 1,
                    badge_text: mb ? mb.innerText.trim() : null,
                    lines: blockLines.length > 0 ? blockLines : null
                });
            });
            if (rowCells.length) rows.push(rowCells);
        });
        tables.push({
            rect: tRect,
            cardTitle: cardTitle,
            inCard: !!parentCard,
            headers: headers,
            rows: rows,
            num_cols: Math.max(headers.length, ...rows.map(r => r.length), 0),
            num_rows: (headers.length ? 1 : 0) + rows.length
        });
    });

    // Highlight / Callout boxes (strictly exclude wrappers like div.foot that contain .foot__pocket)
    const hlBoxes = [];
    const seenHl = new Set();
    s.querySelectorAll('.foot__pocket, .highlight-box, [class*="highlight-box"], [class*="callout"]').forEach(hl => {
        const txt = hl.innerText ? hl.innerText.trim() : '';
        if (!txt || seenHl.has(txt)) return;
        seenHl.add(txt);
        hlBoxes.push({
            rect: toIn(hl.getBoundingClientRect()),
            styles: getStyles(hl),
            text: txt
        });
    });

    // Additional Footnotes / Bottom notes if not already extracted from footer bar
    s.querySelectorAll('.foot__n, .footnote, [class*="footnote"], [class*="bottom-note"]').forEach(el => {
        const txt = el.innerText ? el.innerText.trim() : '';
        if (!txt || seenFn.has(txt)) return;
        seenFn.add(txt);
        footnotes.push({
            rect: toIn(el.getBoundingClientRect()),
            styles: getStyles(el),
            text: txt
        });
    });
    if (footnotes.length === 0) {
        s.querySelectorAll('p, div, span, small').forEach(el => {
            if (el.querySelector('p, div, table')) return;
            const txt = el.innerText ? el.innerText.trim() : '';
            if (!txt || seenFn.has(txt)) return;
            const isFootnote = txt.startsWith('*') || txt.startsWith('※') || txt.startsWith('†') || 
                               txt.startsWith('1)') || txt.startsWith('참조') || txt.startsWith('출처');
            const r = toIn(el.getBoundingClientRect());
            if (isFootnote && r.top > 5.5) {
                seenFn.add(txt);
                footnotes.push({
                    rect: r,
                    styles: getStyles(el),
                    text: txt
                });
            }
        });
    }

    return {
        id: targetId,
        slideBgColor: slideBg,
        topHeader: topHeader,
        headerBadges: headerBadges,
        tag: headerBadges.length > 0 ? headerBadges[0] : null,
        num: num ? { text: num.innerText.trim(), rect: toIn(num.getBoundingClientRect()), styles: getStyles(num) } : null,
        title: mainH ? { text: mainH.innerText.trim(), rect: toIn(mainH.getBoundingClientRect()), styles: getStyles(mainH), runs: titleRuns.length > 1 ? titleRuns : null } : null,
        sub: subDesc ? { text: subDesc.innerText.trim(), rect: toIn(subDesc.getBoundingClientRect()), styles: getStyles(subDesc) } : null,
        desc: desc,
        standaloneTexts: standaloneTexts,
        cards: cards,
        charts: charts,
        tables: tables,
        highlightBoxes: hlBoxes,
        footnotes: footnotes,
        footerDividerTop: footerDividerTop,
        footerDividerColor: footerDividerColor
    };
})()
"""

