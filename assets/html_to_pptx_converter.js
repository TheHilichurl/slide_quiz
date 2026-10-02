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
  async function rasterizeSvgToPng(svgImgEl, scale = 2.5) {
    if (!svgImgEl) return null;
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
      return await renderCardImageToDataUrl(svgImgEl, 223, 56, 0);
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
      // LAYER 1: MANDATORY HEADER
      // =========================================================================
      const headerEl = slideEl.querySelector('.slide-header');
      if (headerEl) {
        const hBox = getBox(headerEl);
        const headerH = hBox ? hBox.h : pxToInch(86);

        // Header Blue Banner (#003882)
        slide.addShape(pptxInstance.shapes.RECTANGLE, {
          x: 0,
          y: 0,
          w: PPTX_WIDTH_INCH,
          h: headerH,
          fill: { color: '003882' },
          line: { type: 'none' }
        });

        // Accent Orange Line at bottom of header (#EA580C)
        slide.addShape(pptxInstance.shapes.LINE, {
          x: 0,
          y: headerH,
          w: PPTX_WIDTH_INCH,
          h: 0,
          line: { color: 'EA580C', width: 2.25 }
        });

        // Slogan Accent Bar (Vertical Orange Pill #FF7A00)
        slide.addShape(pptxInstance.shapes.ROUNDED_RECTANGLE, {
          x: pxToInch(54),
          y: pxToInch(22),
          w: pxToInch(8),
          h: pxToInch(42),
          fill: { color: 'FF7A00' },
          line: { type: 'none' },
          rectRadius: 0.5
        });

        // Slogan Text ("HỌC ĐỂ THAY ĐỔI") - Proportional 24pt bold
        slide.addText('HỌC ĐỂ THAY ĐỔI', {
          x: pxToInch(78),
          y: pxToInch(16),
          w: pxToInch(550),
          h: pxToInch(54),
          fontFace: 'Times New Roman',
          fontSize: 24,
          bold: true,
          color: 'FF7A00',
          charSpacing: 2,
          valign: 'middle',
          margin: 0
        });

        // Official Logo Image (dai-nam-logo-ngang.svg)
        const logoEl = headerEl.querySelector('.official-logo-img');
        if (logoEl) {
          const lBox = getBox(logoEl);
          const logoData = await rasterizeSvgToPng(logoEl, 2.5);
          if (logoData) {
            slide.addImage({
              data: logoData,
              x: lBox ? lBox.x : pxToInch(1643),
              y: lBox ? lBox.y : pxToInch(14),
              w: lBox ? lBox.w : pxToInch(223),
              h: lBox ? lBox.h : pxToInch(56)
            });
          }
        }
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

          slide.addShape(pptxInstance.shapes.ROUNDED_RECTANGLE, {
            x: bBox.x,
            y: bBox.y,
            w: bBox.w,
            h: bBox.h,
            fill: { color: bgCol },
            line: { type: 'none' },
            rectRadius: 0.35
          });
          slide.addText(text, {
            x: bBox.x,
            y: bBox.y,
            w: bBox.w,
            h: bBox.h,
            fontFace: 'Times New Roman',
            fontSize: 16.5,
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

          slide.addShape(pptxInstance.shapes.ROUNDED_RECTANGLE, {
            x: bBox.x,
            y: bBox.y,
            w: bBox.w,
            h: bBox.h,
            fill: { color: bgCol },
            line: { type: 'none' },
            rectRadius: 0.35
          });
          slide.addText(text, {
            x: bBox.x,
            y: bBox.y,
            w: bBox.w,
            h: bBox.h,
            fontFace: 'Times New Roman',
            fontSize: 16.5,
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
            rectRadius: 0.35
          });
          slide.addText(text, {
            x: badgeX,
            y: bBox.y,
            w: badgeW,
            h: bBox.h,
            fontFace: 'Times New Roman',
            fontSize: 14,
            bold: true,
            color: textCol,
            align: 'center',
            valign: 'middle',
            wrap: false,
            margin: 0
          });
        }

        // 2.2 Question Title (Full Width, Auto-wrap, Times New Roman 22.5pt bold - exact web proportion)
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
            fontSize: 22.5,
            bold: true,
            color: '0F172A',
            lineSpacingMultiple: 1.15,
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
            rectRadius: 0.15
          });

          // Check if option has circle letter or plain letter
          const letterCircleEl = optEl.querySelector('.opt-letter-circle');
          const letterPlainEl = optEl.querySelector('.opt-letter');
          const letterText = getCleanText(letterCircleEl || letterPlainEl);
          const correctTagEl = optEl.querySelector('.quiz-correct-tag');

          if (letterCircleEl) {
            // Circle letter badge (computed style matches theme: Orange #EA580C or Emerald #10B981)
            const circleCs = window.getComputedStyle(letterCircleEl);
            const circleBg = parseCssColor(circleCs.backgroundColor, 'EA580C')?.hex || 'EA580C';
            const circleColor = parseCssColor(circleCs.color, 'FFFFFF')?.hex || 'FFFFFF';
            const circleSize = pxToInch(46);
            const circleX = oBox.x + pxToInch(18);
            const circleY = oBox.y + (oBox.h - circleSize) / 2;

            slide.addShape(pptxInstance.shapes.OVAL, {
              x: circleX,
              y: circleY,
              w: circleSize,
              h: circleSize,
              fill: { color: circleBg },
              line: { type: 'none' }
            });
            slide.addText(letterText, {
              x: circleX,
              y: circleY,
              w: circleSize,
              h: circleSize,
              fontFace: 'Times New Roman',
              fontSize: 18,
              bold: true,
              color: circleColor,
              align: 'center',
              valign: 'middle',
              margin: 0
            });
          } else if (letterPlainEl) {
            // Plain letter (A., B., C., D.) - 20pt bold matching web 35px
            const lcs = window.getComputedStyle(letterPlainEl);
            const lColor = parseCssColor(lcs.color, isDimmed ? '94A3B8' : '003882')?.hex || (isDimmed ? '94A3B8' : '003882');
            slide.addText(letterText, {
              x: oBox.x + pxToInch(20),
              y: oBox.y,
              w: pxToInch(48),
              h: oBox.h,
              fontFace: 'Times New Roman',
              fontSize: 20,
              bold: true,
              color: lColor,
              valign: 'middle',
              margin: 0
            });
          }

          // Option Text Body - Proportional 18.5pt with generous width
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
              fontSize: 18.5,
              bold: textBold,
              color: textColor,
              valign: 'middle',
              wrap: true,
              margin: 0
            });
          }

          // Correct Tag Badge ("ĐÁP ÁN ĐÚNG") - single-line width, exact DOM color
          if (correctTagEl) {
            const tagCs = window.getComputedStyle(correctTagEl);
            const tagBg = parseCssColor(tagCs.backgroundColor, 'EA580C')?.hex || 'EA580C';
            const tagTextColor = parseCssColor(tagCs.color, 'FFFFFF')?.hex || 'FFFFFF';
            const tagW = Math.max(pxToInch(190), getBox(correctTagEl)?.w || pxToInch(190));
            const tagH = pxToInch(36);
            const tagX = oBox.x + oBox.w - tagW - pxToInch(18);
            const tagY = oBox.y + (oBox.h - tagH) / 2;

            slide.addShape(pptxInstance.shapes.ROUNDED_RECTANGLE, {
              x: tagX,
              y: tagY,
              w: tagW,
              h: tagH,
              fill: { color: tagBg },
              line: { type: 'none' },
              rectRadius: 0.4
            });
            slide.addText(getCleanText(correctTagEl) || 'ĐÁP ÁN ĐÚNG', {
              x: tagX,
              y: tagY,
              w: tagW,
              h: tagH,
              fontFace: 'Times New Roman',
              fontSize: 12,
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
            // Place image at EXACT visual card bounds - NO background shape with white gap, NO inset!
            slide.addImage({
              data: imgData,
              x: vBox.x,
              y: vBox.y,
              w: vBox.w,
              h: vBox.h
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
                rectRadius: 0.12
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
