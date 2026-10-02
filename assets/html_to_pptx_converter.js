/**
 * HTML to Editable PPTX Converter Engine
 * Đại Nam University ("Đại Nam Style") Presentation Framework
 * 
 * Standardizes HTML/CSS DOM elements into native PowerPoint (OpenXML/PptxGenJS) design primitives:
 * - 16:9 Coordinate Space Mapping (1920x1080px -> 13.333x7.5 inches, 144 DPI)
 * - Proportional Font Sizing matching web projection scale (Title: 26pt, Option text: 18.5pt, Slogan: 24pt, Badges: 15-16.5pt)
 * - Picture Objects: 18px rounded corner clipping with border stroked DIRECTLY ON CANVAS (100% coincident, ZERO gap)
 * - Full Typography Fidelity: Strictly Times New Roman, ample width for single-line fit (NO premature line wraps)
 * - Exact Brand Color Mapping: #003882 (Imperial Blue), #EA580C (Energetic Orange), #FF7A00 (Slogan Accent), #10B981 (Emerald)
 * - Clean Footer: NO timer display, clean single-line clickable dainam.edu.vn link
 */

(function (global) {
  'use strict';

  // Constants
  const CANVAS_WIDTH_PX = 1920;
  const CANVAS_HEIGHT_PX = 1080;
  const PPTX_WIDTH_INCH = 13.333333;
  const PPTX_HEIGHT_INCH = 7.5;
  const DPI_SCALE = CANVAS_WIDTH_PX / PPTX_WIDTH_INCH; // Exactly 144 px = 1 inch

  /**
   * Helper: Convert pixel length to inches
   */
  function pxToInch(px) {
    return (px || 0) / DPI_SCALE;
  }

  /**
   * Helper: Convert CSS font size (px) to proportional PowerPoint font size (pt)
   * Using 0.62 scaling factor to match projector-scale web typography
   */
  function pxToPt(px, defaultPt = 16) {
    if (!px) return defaultPt;
    const val = typeof px === 'number' ? px : parseFloat(px);
    if (isNaN(val)) return defaultPt;
    return Math.round(val * 0.62 * 10) / 10;
  }

  /**
   * Helper: Get clean text from element (works on hidden / transitioning elements)
   */
  function getCleanText(el) {
    if (!el) return '';
    const txt = el.innerText || el.textContent || '';
    return txt.replace(/\s+/g, ' ').trim();
  }

  /**
   * Helper: Parse CSS color into PPTX 6-char hex and transparency
   */
  function parseCssColor(cssColor, defaultHex = '000000') {
    if (!cssColor || cssColor === 'transparent' || cssColor === 'rgba(0, 0, 0, 0)') {
      return null;
    }

    if (cssColor.startsWith('#')) {
      let hex = cssColor.replace('#', '').trim();
      if (hex.length === 3) {
        hex = hex.split('').map(c => c + c).join('');
      }
      return { hex: hex.toUpperCase(), transparency: 0 };
    }

    const match = cssColor.match(/rgba?\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)(?:\s*,\s*([\d.]+))?\s*\)/i);
    if (match) {
      const r = parseInt(match[1], 10).toString(16).padStart(2, '0');
      const g = parseInt(match[2], 10).toString(16).padStart(2, '0');
      const b = parseInt(match[3], 10).toString(16).padStart(2, '0');
      const alpha = match[4] !== undefined ? parseFloat(match[4]) : 1.0;
      const transparency = Math.round((1 - alpha) * 100);
      return {
        hex: (r + g + b).toUpperCase(),
        transparency: transparency
      };
    }

    return { hex: defaultHex, transparency: 0 };
  }

  /**
   * Helper: Render image with exact 18px rounded corner clipping and stroke border DIRECTLY on canvas!
   * This ensures the border is 100% coincident with the image edge with ZERO white gap!
   */
  async function renderCardImageToDataUrl(imgEl, targetWidthPx, targetHeightPx, radiusPx = 18, borderColorHex = null, borderWidthPx = 0) {
    if (!imgEl) return null;

    try {
      const scale = 2; // High-DPI 2x supersampling
      const cw = Math.round((targetWidthPx || imgEl.clientWidth || 610) * scale);
      const ch = Math.round((targetHeightPx || imgEl.clientHeight || 765) * scale);
      const cr = Math.round((radiusPx || 18) * scale);
      const bw = Math.round((borderWidthPx || 0) * scale);

      const canvas = document.createElement('canvas');
      canvas.width = cw;
      canvas.height = ch;
      const ctx = canvas.getContext('2d');
      ctx.imageSmoothingEnabled = true;
      ctx.imageSmoothingQuality = 'high';

      // Step 1: Clip image path (18px rounded rect)
      ctx.save();
      ctx.beginPath();
      ctx.roundRect(0, 0, cw, ch, cr);
      ctx.clip();

      // Exact object-fit: cover math
      const iw = imgEl.naturalWidth || imgEl.width || cw;
      const ih = imgEl.naturalHeight || imgEl.height || ch;
      const ratio = Math.max(cw / iw, ch / ih);
      const nw = iw * ratio;
      const nh = ih * ratio;
      const cx = (cw - nw) / 2;
      const cy = (ch - nh) / 2;

      ctx.drawImage(imgEl, cx, cy, nw, nh);
      ctx.restore();

      // Step 2: Stroke border DIRECTLY onto the canvas image edge!
      // Inset stroke by half border width so it aligns flush with the outer perimeter (ZERO GAP!)
      if (bw > 0 && borderColorHex) {
        ctx.beginPath();
        const halfBw = bw / 2;
        ctx.roundRect(halfBw, halfBw, cw - bw, ch - bw, Math.max(0, cr - halfBw));
        ctx.lineWidth = bw;
        ctx.strokeStyle = '#' + borderColorHex;
        ctx.stroke();
      }

      return canvas.toDataURL('image/png');
    } catch (e) {
      try {
        const resp = await fetch(imgEl.src);
        const blob = await resp.blob();
        return await new Promise((resolve, reject) => {
          const reader = new FileReader();
          reader.onloadend = () => resolve(reader.result);
          reader.onerror = reject;
          reader.readAsDataURL(blob);
        });
      } catch (err) {
        if (imgEl.src && imgEl.src.startsWith('data:')) {
          return imgEl.src;
        }
        console.warn('[PPTX Exporter] Could not extract image dataUrl:', imgEl.src);
        return null;
      }
    }
  }

  /**
   * Helper: Convert SVG logo to high-res PNG Data URL
   */
  const DAI_NAM_LOGO_DATA_URL = "data:image/png;base64,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";

  async function rasterizeSvgToPng(svgImgEl, scale = 2.5) {
    if (!svgImgEl) return DAI_NAM_LOGO_DATA_URL;
    if (svgImgEl.src && svgImgEl.src.includes('dai-nam-logo')) {
      return DAI_NAM_LOGO_DATA_URL;
    }
    try {
      const canvas = document.createElement('canvas');
      const w = (svgImgEl.naturalWidth || svgImgEl.width || 223) * scale;
      const h = (svgImgEl.naturalHeight || svgImgEl.height || 56) * scale;
      canvas.width = w;
      canvas.height = h;
      const ctx = canvas.getContext('2d');
      ctx.imageSmoothingEnabled = true;
      ctx.imageSmoothingQuality = 'high';
      ctx.drawImage(svgImgEl, 0, 0, w, h);
      return canvas.toDataURL('image/png');
    } catch (e) {
      return DAI_NAM_LOGO_DATA_URL;
    }
  }

  /**
   * Helper: Generate smooth warm ambient radial glow for PPTX background
   */
  let _cachedAmbientGlowData = null;
  function getAmbientGlowDataUrl() {
    if (_cachedAmbientGlowData) return _cachedAmbientGlowData;
    try {
      const canvas = document.createElement('canvas');
      canvas.width = 600;
      canvas.height = 600;
      const ctx = canvas.getContext('2d');
      const grad = ctx.createRadialGradient(300, 300, 0, 300, 300, 300);
      grad.addColorStop(0, 'rgba(234, 88, 12, 0.08)');
      grad.addColorStop(0.4, 'rgba(0, 56, 130, 0.035)');
      grad.addColorStop(1, 'rgba(255, 255, 255, 0)');
      ctx.fillStyle = grad;
      ctx.fillRect(0, 0, 600, 600);
      _cachedAmbientGlowData = canvas.toDataURL('image/png');
      return _cachedAmbientGlowData;
    } catch (e) {
      return null;
    }
  }

  /**
   * Core Engine: Convert single Slide DOM node into native PPTX Slide
   */
  async function convertSlideToPptx(slideEl, pptxInstance, slideIndex, totalSlides) {
    const slide = pptxInstance.addSlide();
    
    // Ensure slide has active geometry for precise getBoundingClientRect
    const origOpacity = slideEl.style.opacity;
    const origVisibility = slideEl.style.visibility;
    const origTransition = slideEl.style.transition;
    const origTransform = slideEl.style.transform;

    slideEl.style.opacity = '1';
    slideEl.style.visibility = 'visible';
    slideEl.style.transition = 'none';
    slideEl.style.transform = 'none';

    const sRect = slideEl.getBoundingClientRect();
    const slideComp = window.getComputedStyle(slideEl);

    // 1. Slide Background
    const isDarkTheme = slideEl.classList.contains('theme-dark') || 
                        slideComp.backgroundColor.includes('rgb(5,') ||
                        slideComp.backgroundColor.includes('rgb(11,');
    slide.background = { color: isDarkTheme ? '050B14' : 'F8FAFC' };

    // Ambient soft glow in background (warm top-right light matching Web UI)
    const ambientGlow = getAmbientGlowDataUrl();
    if (ambientGlow && !isDarkTheme) {
      slide.addImage({
        data: ambientGlow,
        x: PPTX_WIDTH_INCH - 6.5,
        y: 0.8,
        w: 6.5,
        h: 6.5
      });
    }

    // Helper for relative bounding box in inches
    function getBox(el) {
      if (!el) return null;
      const r = el.getBoundingClientRect();
      return {
        x: pxToInch(r.left - sRect.left),
        y: pxToInch(r.top - sRect.top),
        w: pxToInch(r.width),
        h: pxToInch(r.height),
        rawX: r.left - sRect.left,
        rawY: r.top - sRect.top,
        rawW: r.width,
        rawH: r.height
      };
    }

    try {
      // =========================================================================
      // LAYER 1: MANDATORY HEADER (Logo on Left, Divider, Slogan, and Right Badge)
      // =========================================================================
      const headerEl = slideEl.querySelector('.slide-header');
      if (headerEl) {
        const hBox = getBox(headerEl);
        const headerH = hBox ? hBox.h : pxToInch(86);

        // Header Blue Banner (#003882) with subtle bottom elevation shadow
        slide.addShape(pptxInstance.shapes.RECTANGLE, {
          x: 0,
          y: 0,
          w: PPTX_WIDTH_INCH,
          h: headerH,
          fill: { color: '003882' },
          line: { type: 'none' },
          shadow: { type: 'outer', color: '000000', blur: 6, offset: 2.5, angle: 90, opacity: 0.18 }
        });

        // Accent Orange Line at bottom of header (#EA580C)
        slide.addShape(pptxInstance.shapes.LINE, {
          x: 0,
          y: headerH,
          w: PPTX_WIDTH_INCH,
          h: 0,
          line: { color: 'EA580C', width: 2.25 }
        });

        // 1. Official Logo Image on the far left of Header-Left
        const logoEl = headerEl.querySelector('.official-logo-img');
        let logoEndX = pxToInch(54);
        if (logoEl) {
          const lBox = getBox(logoEl);
          const logoData = await rasterizeSvgToPng(logoEl, 2.5);
          if (logoData) {
            const logoW = lBox ? lBox.w : pxToInch(205);
            const logoH = lBox ? lBox.h : pxToInch(52);
            const logoX = pxToInch(54);
            const logoY = pxToInch(17);
            slide.addImage({
              data: logoData,
              x: logoX,
              y: logoY,
              w: logoW,
              h: logoH
            });
            logoEndX = logoX + logoW;
          }
        }

        // 2. Vertical Divider between Logo and Slogan
        const dividerX = logoEndX + pxToInch(18);
        slide.addShape(pptxInstance.shapes.LINE, {
          x: dividerX,
          y: pxToInch(24),
          w: 0,
          h: pxToInch(38),
          line: { color: '4B72A4', width: 1.5 }
        });

        // 3. Slogan Text ("HỌC ĐỂ THAY ĐỔI") immediately following Logo & Divider
        const sloganX = dividerX + pxToInch(18);
        slide.addText('HỌC ĐỂ THAY ĐỔI', {
          x: sloganX,
          y: pxToInch(16),
          w: pxToInch(620),
          h: pxToInch(54),
          fontFace: 'Times New Roman',
          fontSize: 22,
          bold: true,
          color: 'FF7A00',
          charSpacing: 2,
          valign: 'middle',
          wrap: false,
          margin: 0
        });

        // 4. Header-Right Badge ("HỆ THỐNG ÔN TẬP GDQP-AN")
        const headerBadgeEl = headerEl.querySelector('.header-badge');
        const badgeText = headerBadgeEl ? getCleanText(headerBadgeEl) : 'HỆ THỐNG ÔN TẬP GDQP-AN';
        const hBadgeW = pxToInch(330);
        const hBadgeH = pxToInch(42);
        const hBadgeX = PPTX_WIDTH_INCH - hBadgeW - pxToInch(54);
        const hBadgeY = pxToInch(22);

        slide.addShape(pptxInstance.shapes.ROUNDED_RECTANGLE, {
          x: hBadgeX,
          y: hBadgeY,
          w: hBadgeW,
          h: hBadgeH,
          fill: { color: '144482' },
          line: { color: '2E60A3', width: 1.2 },
          rectRadius: 0.5,
          shadow: { type: 'outer', color: '000000', blur: 4, offset: 1.5, angle: 90, opacity: 0.15 }
        });
        slide.addText(badgeText, {
          x: hBadgeX,
          y: hBadgeY,
          w: hBadgeW,
          h: hBadgeH,
          fontFace: 'Times New Roman',
          fontSize: 11,
          bold: true,
          color: 'FFFFFF',
          align: 'center',
          valign: 'middle',
          margin: 0
        });
      }

      // =========================================================================
      // LAYER 2: SLIDE BODY & QUIZ CONTENT
      // =========================================================================
      const quizContainer = slideEl.querySelector('.quiz-container');
      if (quizContainer) {
        // 2.1 Badges
        const qBadge = slideEl.querySelector('.quiz-badge-question');
        const aBadge = slideEl.querySelector('.quiz-badge-answer');
        const tBadge = slideEl.querySelector('.quiz-topic-badge');

        if (qBadge) {
          const bBox = getBox(qBadge);
          const cs = window.getComputedStyle(qBadge);
          const text = getCleanText(qBadge) || `CÂU HỎI ${String(slideIndex).padStart(2, '0')}`;
          const bgCol = parseCssColor(cs.backgroundColor, '003882')?.hex || '003882';

          // Standardized 10px rounded rectangle (rectRadius: 0.05) with elevation shadow
          slide.addShape(pptxInstance.shapes.ROUNDED_RECTANGLE, {
            x: bBox.x,
            y: bBox.y,
            w: bBox.w,
            h: bBox.h,
            fill: { color: bgCol },
            line: { type: 'none' },
            rectRadius: 0.05,
            shadow: { type: 'outer', color: '003882', blur: 5, offset: 2, angle: 90, opacity: 0.20 }
          });
          slide.addText(text, {
            x: bBox.x,
            y: bBox.y,
            w: bBox.w,
            h: bBox.h,
            fontFace: 'Times New Roman',
            fontSize: 13.5,
            bold: true,
            color: 'FFFFFF',
            align: 'center',
            valign: 'middle',
            margin: 0
          });
        }

        if (aBadge) {
          const bBox = getBox(aBadge);
          const cs = window.getComputedStyle(aBadge);
          const text = getCleanText(aBadge) || `ĐÁP ÁN CÂU ${String(Math.ceil(slideIndex / 2)).padStart(2, '0')}`;
          const bgCol = parseCssColor(cs.backgroundColor, 'EA580C')?.hex || 'EA580C';

          // Standardized 10px rounded rectangle (rectRadius: 0.05) with glowing orange shadow
          slide.addShape(pptxInstance.shapes.ROUNDED_RECTANGLE, {
            x: bBox.x,
            y: bBox.y,
            w: bBox.w,
            h: bBox.h,
            fill: { color: bgCol },
            line: { type: 'none' },
            rectRadius: 0.05,
            shadow: { type: 'outer', color: 'EA580C', blur: 6, offset: 2, angle: 90, opacity: 0.28 }
          });
          slide.addText(text, {
            x: bBox.x,
            y: bBox.y,
            w: bBox.w,
            h: bBox.h,
            fontFace: 'Times New Roman',
            fontSize: 13.5,
            bold: true,
            color: 'FFFFFF',
            align: 'center',
            valign: 'middle',
            margin: 0
          });
        }

        if (tBadge) {
          const bBox = getBox(tBadge);
          const cs = window.getComputedStyle(tBadge);
          const text = getCleanText(tBadge);
          const bgCol = parseCssColor(cs.backgroundColor, 'FFF7ED')?.hex || 'FFF7ED';
          const lineCol = parseCssColor(cs.borderColor, 'FFEDD5')?.hex || 'FFEDD5';
          const textCol = parseCssColor(cs.color, 'C2410C')?.hex || 'C2410C';

          const badgeExtraW = pxToInch(45);
          const badgeW = bBox.w + badgeExtraW;
          const badgeX = bBox.x - badgeExtraW;

          slide.addShape(pptxInstance.shapes.ROUNDED_RECTANGLE, {
            x: badgeX,
            y: bBox.y,
            w: badgeW,
            h: bBox.h,
            fill: { color: bgCol },
            line: { color: lineCol, width: 1.5 },
            rectRadius: 0.05,
            shadow: { type: 'outer', color: 'EA580C', blur: 4, offset: 1.5, angle: 90, opacity: 0.12 }
          });
          slide.addText(text, {
            x: badgeX,
            y: bBox.y,
            w: badgeW,
            h: bBox.h,
            fontFace: 'Times New Roman',
            fontSize: 12.5,
            bold: true,
            color: textCol,
            align: 'center',
            valign: 'middle',
            wrap: false,
            margin: 0
          });
        }

        // 2.2 Question Title (Full Width, Auto-wrap, Times New Roman 20pt bold - perfectly fitted)
        const titleEl = slideEl.querySelector('.quiz-question-title');
        if (titleEl) {
          const tBox = getBox(titleEl);
          const titleText = getCleanText(titleEl);
          const titleW = Math.max(tBox.w, pxToInch(1155));
          const titleH = Math.max(tBox.h, pxToInch(245));
          slide.addText(titleText, {
            x: tBox.x,
            y: tBox.y,
            w: titleW,
            h: titleH,
            fontFace: 'Times New Roman',
            fontSize: 20,
            bold: true,
            color: '0F172A',
            lineSpacingMultiple: 1.18,
            valign: 'top',
            wrap: true,
            margin: 0
          });
        }

        // 2.3 Option Cards List
        const optionCards = slideEl.querySelectorAll('.quiz-option-card');
        for (let optEl of optionCards) {
          const oBox = getBox(optEl);
          const cs = window.getComputedStyle(optEl);
          const isCorrect = optEl.classList.contains('correct-answer');
          const isDimmed = optEl.classList.contains('dimmed');

          // Extract exact computed colors from DOM
          const cardFill = parseCssColor(cs.backgroundColor, isCorrect ? 'FFF7ED' : (isDimmed ? 'F8FAFC' : 'FFFFFF'))?.hex || 'FFFFFF';
          const cardBorderColor = parseCssColor(cs.borderColor, isCorrect ? 'EA580C' : 'CBD5E1')?.hex || 'CBD5E1';
          const cardBorderWidth = parseFloat(cs.borderWidth) || (isCorrect ? 3.5 : 2);

          slide.addShape(pptxInstance.shapes.ROUNDED_RECTANGLE, {
            x: oBox.x,
            y: oBox.y,
            w: oBox.w,
            h: oBox.h,
            fill: { color: cardFill },
            line: { color: cardBorderColor, width: cardBorderWidth },
            rectRadius: 0.07,
            shadow: isCorrect ? {
              type: 'outer',
              color: 'EA580C',
              blur: 14,
              offset: 0,
              angle: 0,
              opacity: 0.42
            } : (isDimmed ? undefined : {
              type: 'outer',
              color: '003882',
              blur: 6,
              offset: 2.5,
              angle: 90,
              opacity: 0.09
            })
          });

          // Check if option has circle letter or plain letter
          const letterCircleEl = optEl.querySelector('.opt-letter-circle');
          const letterPlainEl = optEl.querySelector('.opt-letter');
          const letterText = getCleanText(letterCircleEl || letterPlainEl);
          const correctTagEl = optEl.querySelector('.quiz-correct-tag');

          if (letterCircleEl) {
            // Circle letter badge: exact 50px DOM size, 15pt bold centered
            const circleCs = window.getComputedStyle(letterCircleEl);
            const circleBg = parseCssColor(circleCs.backgroundColor, 'EA580C')?.hex || 'EA580C';
            const circleColor = parseCssColor(circleCs.color, 'FFFFFF')?.hex || 'FFFFFF';
            const circleSize = pxToInch(50);
            const circleX = oBox.x + pxToInch(20);
            const circleY = oBox.y + (oBox.h - circleSize) / 2;

            slide.addShape(pptxInstance.shapes.OVAL, {
              x: circleX,
              y: circleY,
              w: circleSize,
              h: circleSize,
              fill: { color: circleBg },
              line: { type: 'none' },
              shadow: { type: 'outer', color: 'EA580C', blur: 4, offset: 1.5, angle: 90, opacity: 0.25 }
            });
            slide.addText(letterText, {
              x: circleX,
              y: circleY,
              w: circleSize,
              h: circleSize,
              fontFace: 'Times New Roman',
              fontSize: 15,
              bold: true,
              color: circleColor,
              align: 'center',
              valign: 'middle',
              margin: 0
            });
          } else if (letterPlainEl) {
            // Plain letter (A., B., C., D.) - 18pt bold
            const lcs = window.getComputedStyle(letterPlainEl);
            const lColor = parseCssColor(lcs.color, isDimmed ? '94A3B8' : '003882')?.hex || (isDimmed ? '94A3B8' : '003882');
            slide.addText(letterText, {
              x: oBox.x + pxToInch(20),
              y: oBox.y,
              w: pxToInch(48),
              h: oBox.h,
              fontFace: 'Times New Roman',
              fontSize: 18,
              bold: true,
              color: lColor,
              valign: 'middle',
              margin: 0
            });
          }

          // Option Text Body - Proportional 16.5pt perfectly contained inside card box
          const optTextEl = optEl.querySelector('.opt-text');
          if (optTextEl) {
            const tcs = window.getComputedStyle(optTextEl);
            const textStartX = oBox.x + pxToInch(76);
            const textWidth = isCorrect 
              ? (oBox.w - pxToInch(310))
              : (oBox.w - pxToInch(90));

            let textColor = parseCssColor(tcs.color, isCorrect ? '9A3412' : (isDimmed ? '64748B' : '1E293B'))?.hex || (isDimmed ? '64748B' : '1E293B');
            let textBold = isCorrect || tcs.fontWeight === '700' || tcs.fontWeight === 'bold';

            slide.addText(getCleanText(optTextEl), {
              x: textStartX,
              y: oBox.y,
              w: textWidth,
              h: oBox.h,
              fontFace: 'Times New Roman',
              fontSize: 16.5,
              bold: textBold,
              color: textColor,
              lineSpacingMultiple: 1.15,
              valign: 'middle',
              wrap: true,
              margin: 0
            });
          }

          // Correct Tag Badge ("ĐÁP ÁN ĐÚNG") - Standardized 8px rounded rectangle (rectRadius: 0.05), exact DOM metrics
          if (correctTagEl) {
            const tagCs = window.getComputedStyle(correctTagEl);
            const tagBg = parseCssColor(tagCs.backgroundColor, 'EA580C')?.hex || 'EA580C';
            const tagTextColor = parseCssColor(tagCs.color, 'FFFFFF')?.hex || 'FFFFFF';
            const tagBox = getBox(correctTagEl);
            const tagW = tagBox ? Math.max(tagBox.w, pxToInch(197)) : pxToInch(197);
            const tagH = tagBox ? tagBox.h : pxToInch(39.4);
            const tagX = oBox.x + oBox.w - tagW - pxToInch(22);
            const tagY = oBox.y + (oBox.h - tagH) / 2;

            slide.addShape(pptxInstance.shapes.ROUNDED_RECTANGLE, {
              x: tagX,
              y: tagY,
              w: tagW,
              h: tagH,
              fill: { color: tagBg },
              line: { type: 'none' },
              rectRadius: 0.05,
              shadow: { type: 'outer', color: 'EA580C', blur: 5, offset: 1.5, angle: 90, opacity: 0.30 }
            });
            slide.addText(getCleanText(correctTagEl) || 'ĐÁP ÁN ĐÚNG', {
              x: tagX,
              y: tagY,
              w: tagW,
              h: tagH,
              fontFace: 'Times New Roman',
              fontSize: 11.5,
              bold: true,
              color: tagTextColor,
              align: 'center',
              valign: 'middle',
              wrap: false,
              margin: 0
            });
          }
        }

        // 2.4 Visual Media Card & Image (BORDER STROKED DIRECTLY ON CANVAS: 100% COINCIDENT, ZERO GAP!)
        const visualCardEl = slideEl.querySelector('.quiz-visual-card');
        const visualImgEl = slideEl.querySelector('.quiz-visual-img');
        if (visualCardEl && visualImgEl) {
          const vBox = getBox(visualCardEl);
          const vcs = window.getComputedStyle(visualCardEl);
          const borderColorHex = parseCssColor(vcs.borderColor, 'CBD5E1')?.hex || 'CBD5E1';
          const borderWidthPx = parseFloat(vcs.borderWidth) || 2;
          const borderRadiusPx = parseFloat(vcs.borderRadius) || 18;
          const isCorrectVisual = visualCardEl.classList.contains('correct-visual');

          // Render image onto offscreen canvas with exact 18px rounded corner clipping AND border stroked flush on edge!
          const imgData = await renderCardImageToDataUrl(
            visualImgEl,
            vBox.rawW,
            vBox.rawH,
            borderRadiusPx,
            borderColorHex,
            borderWidthPx
          );

          if (imgData) {
            // Place image at EXACT visual card bounds with realistic shadow/glow!
            slide.addImage({
              data: imgData,
              x: vBox.x,
              y: vBox.y,
              w: vBox.w,
              h: vBox.h,
              shadow: isCorrectVisual ? {
                type: 'outer',
                color: 'EA580C',
                blur: 16,
                offset: 0,
                angle: 0,
                opacity: 0.45
              } : {
                type: 'outer',
                color: '003882',
                blur: 10,
                offset: 3.5,
                angle: 90,
                opacity: 0.18
              }
            });
          }
        }
      } else {
        // --- GENERAL PRESENTATION LAYOUT FALLBACK ---
        const contentContainers = slideEl.querySelectorAll('.hero-content, .slide-body, .definition-box, .pillar-card, .comparison-card, .matrix-card');
        const targetContainer = contentContainers.length > 0 ? contentContainers : [slideEl.querySelector('.slide-body') || slideEl];

        for (let container of targetContainer) {
          if (!container) continue;

          // Cards & containers with background/border
          const cards = container.querySelectorAll('.card, .hero-visual-card, .hero-meta-grid > div, .pillar-card, .characteristic-card, .takeaway-card');
          for (let card of cards) {
            const cBox = getBox(card);
            const cs = window.getComputedStyle(card);
            const bg = parseCssColor(cs.backgroundColor, 'FFFFFF');
            const border = parseCssColor(cs.borderColor, 'E2E8F0');
            const bWidth = parseFloat(cs.borderWidth) || 1;

            if (bg) {
              slide.addShape(pptxInstance.shapes.ROUNDED_RECTANGLE, {
                x: cBox.x,
                y: cBox.y,
                w: cBox.w,
                h: cBox.h,
                fill: { color: bg.hex, transparency: bg.transparency },
                line: border ? { color: border.hex, width: bWidth } : { type: 'none' },
                rectRadius: 0.12,
                shadow: { type: 'outer', color: '003882', blur: 6, offset: 2.5, angle: 90, opacity: 0.10 }
              });
            }
          }

          // Headings & Titles
          const headings = container.querySelectorAll('h1, h2, h3, .hero-main-title, .slide-title');
          for (let h of headings) {
            const hBox = getBox(h);
            const hs = window.getComputedStyle(h);
            const hColor = parseCssColor(hs.color, '003882');
            slide.addText(getCleanText(h), {
              x: hBox.x,
              y: hBox.y,
              w: Math.max(hBox.w, pxToInch(1140)),
              h: hBox.h,
              fontFace: 'Times New Roman',
              fontSize: pxToPt(hs.fontSize, 26),
              bold: true,
              color: hColor ? hColor.hex : '003882',
              wrap: true,
              margin: 0
            });
          }

          // Paragraphs & Lists
          const paragraphs = container.querySelectorAll('p, .hero-sub-text, .slide-subtitle, li');
          for (let p of paragraphs) {
            if (p.closest('.quiz-option-card')) continue;
            const pBox = getBox(p);
            const ps = window.getComputedStyle(p);
            const pColor = parseCssColor(ps.color, '0F172A');
            slide.addText(getCleanText(p), {
              x: pBox.x,
              y: pBox.y,
              w: Math.max(pBox.w, pxToInch(900)),
              h: pBox.h,
              fontFace: 'Times New Roman',
              fontSize: pxToPt(ps.fontSize, 18.5),
              color: pColor ? pColor.hex : '0F172A',
              wrap: true,
              margin: 0
            });
          }

          // Images
          const images = container.querySelectorAll('img');
          for (let img of images) {
            if (img.classList.contains('official-logo-img') || img.classList.contains('quiz-visual-img')) continue;
            const imgBox = getBox(img);
            const imgData = await renderCardImageToDataUrl(img, imgBox.rawW, imgBox.rawH, 16);
            if (imgData) {
              slide.addImage({
                data: imgData,
                x: imgBox.x,
                y: imgBox.y,
                w: imgBox.w,
                h: imgBox.h
              });
            }
          }
        }
      }

      // =========================================================================
      // LAYER 3: MANDATORY FOOTER (NO TIMER, PROPER LINK WIDTH)
      // =========================================================================
      const footerEl = slideEl.querySelector('.slide-footer');
      if (footerEl) {
        const fBox = getBox(footerEl);
        const footerY = fBox ? fBox.y : pxToInch(1018);
        const footerH = fBox ? fBox.h : pxToInch(62);

        // Footer Blue Banner (#003882)
        slide.addShape(pptxInstance.shapes.RECTANGLE, {
          x: 0,
          y: footerY,
          w: PPTX_WIDTH_INCH,
          h: footerH,
          fill: { color: '003882' },
          line: { type: 'none' }
        });

        // Accent Orange Line at top of footer (#EA580C)
        slide.addShape(pptxInstance.shapes.LINE, {
          x: 0,
          y: footerY,
          w: PPTX_WIDTH_INCH,
          h: 0,
          line: { color: 'EA580C', width: 2.25 }
        });

        // NOTE: TIMER IS INTENTIONALLY OMITTED PER USER MANDATE ("ko có đồng hồ đếm")

        // Footer Center Metadata (Multi-run text) - 13pt bold
        slide.addText([
          { text: 'Trường Đại học Đại Nam', options: { color: 'FB923C', bold: true } },
          { text: ' • ', options: { color: 'FFFFFF' } },
          { text: 'GV: Đào Bá Công', options: { color: 'FFFFFF', bold: true } }
        ], {
          x: 3.5,
          y: footerY,
          w: 6.333,
          h: footerH,
          fontFace: 'Times New Roman',
          fontSize: 13,
          align: 'center',
          valign: 'middle',
          margin: 0
        });

        // Footer Link (Clean single line, adequate width, no arrow wrapping) - 12pt bold
        slide.addText('dainam.edu.vn', {
          x: 11.2,
          y: footerY,
          w: 1.8,
          h: footerH,
          fontFace: 'Times New Roman',
          fontSize: 12,
          bold: true,
          color: 'FB923C',
          align: 'right',
          valign: 'middle',
          margin: 0,
          hyperlink: { url: 'https://dainam.edu.vn' }
        });
      }
    } finally {
      // Restore slide properties
      slideEl.style.opacity = origOpacity;
      slideEl.style.visibility = origVisibility;
      slideEl.style.transition = origTransition;
      slideEl.style.transform = origTransform;
    }
  }

  /**
   * Main Public API: Export All Slides to Editable PPTX
   */
  async function exportPresentationToEditablePPTX(options = {}) {
    const {
      fileName = 'Presentation_DaiNam_Editable.pptx',
      title = 'Bài Giảng Điện Tử GDQP-AN | Đại Học Đại Nam',
      author = 'GV: Đào Bá Công - Đại Học Đại Nam',
      company = 'Trường Đại học Đại Nam (DNU)',
      onProgress = null
    } = options;

    if (typeof PptxGenJS === 'undefined') {
      throw new Error('Thư viện PptxGenJS chưa sẵn sàng! Vui lòng kiểm tra kết nối CDN.');
    }

    const pptx = new PptxGenJS();
    
    // Set 16:9 Standard Presentation Dimensions (13.333 x 7.5 inches)
    pptx.defineLayout({ name: 'LAYOUT_16X9_DNU', width: PPTX_WIDTH_INCH, height: PPTX_HEIGHT_INCH });
    pptx.layout = 'LAYOUT_16X9_DNU';
    pptx.title = title;
    pptx.author = author;
    pptx.company = company;

    // Save current slide state
    const allSlides = document.querySelectorAll('.slide');
    const total = allSlides.length;
    let savedIndex = 1;

    allSlides.forEach((s, idx) => {
      if (s.classList.contains('active')) {
        savedIndex = idx + 1;
      }
    });

    const stage = document.getElementById('slide-stage');
    const origTransform = stage ? stage.style.transform : '';

    try {
      // Normalize stage for 1:1 unscaled measurement
      if (stage) {
        stage.style.transform = 'none';
        stage.style.left = '0px';
        stage.style.top = '0px';
      }

      for (let i = 1; i <= total; i++) {
        if (onProgress) {
          onProgress(i, total, `Đang chuẩn hoá Slide ${i}/${total} sang PowerPoint Shapes & Text...`);
        }

        allSlides.forEach(s => s.classList.remove('active'));
        const targetSlide = document.getElementById(`slide-${i}`) || allSlides[i - 1];
        if (!targetSlide) continue;

        targetSlide.classList.add('active');
        if (window.lucide) window.lucide.createIcons();

        // Allow micro-layout repaint
        await new Promise(r => setTimeout(r, 60));

        // Convert slide
        await convertSlideToPptx(targetSlide, pptx, i, total);
      }

      if (onProgress) {
        onProgress(total, total, 'Đang đóng gói và hoàn tất file PowerPoint (.pptx)...');
      }

      await pptx.writeFile({ fileName });
      return { success: true, totalSlides: total, fileName };
    } finally {
      // Restore slide state
      if (stage) {
        stage.style.transform = origTransform;
        stage.style.left = '50%';
        stage.style.top = '50%';
      }
      allSlides.forEach(s => s.classList.remove('active'));
      const restoreSlide = document.getElementById(`slide-${savedIndex}`) || allSlides[savedIndex - 1];
      if (restoreSlide) restoreSlide.classList.add('active');
      if (typeof window.updateStageScale === 'function') {
        window.updateStageScale();
      }
    }
  }

  // Export to global scope
  global.HtmlToPptxConverter = {
    exportPresentationToEditablePPTX,
    convertSlideToPptx,
    pxToInch,
    pxToPt,
    getCleanText,
    parseCssColor
  };

})(window);
