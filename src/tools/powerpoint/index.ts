import type { PptToolConfig } from '../codegen';
import { createPptTools } from '../codegen';
import pptxgen from 'pptxgenjs';

/* global PowerPoint */

// ─── Internal helper (within a PowerPoint.run context — no inner runs) ────────
async function loadSlideTexts(
  slides: PowerPoint.SlideCollection,
  start: number,
  end: number,
  context: PowerPoint.RequestContext
): Promise<string[]> {
  const slideRefs = slides.items.slice(start, end + 1);
  for (const slide of slideRefs) {
    slide.shapes.load('items');
  }
  await context.sync();
  for (const slide of slideRefs) {
    for (const shape of slide.shapes.items) {
      try {
        shape.textFrame.textRange.load('text');
      } catch {
        // shape may not have a textFrame
      }
    }
  }
  await context.sync();
  const results: string[] = [];
  for (let i = 0; i < slideRefs.length; i++) {
    const slide = slideRefs[i];
    const texts = slide.shapes.items
      .map(s => {
        try {
          return s.textFrame.textRange.text?.trim() ?? '';
        } catch {
          return '';
        }
      })
      .filter(t => t.length > 0);
    results.push(`Slide ${start + i + 1}: ${texts.length > 0 ? texts.join(' | ') : '(no text)'}`);
  }
  return results;
}

// ─── Tool Configs ──────────────────────────────────────────────────────────────

export const powerPointConfigs: readonly PptToolConfig[] = [
  {
    name: 'get_presentation_overview',
    description:
      "Get a full overview of the PowerPoint presentation: total slide count, a text preview of each slide's shapes, AND a PNG thumbnail image of every slide. " +
      'Call this FIRST before making any changes. The thumbnail images let you see the exact visual layout, design, and positioning of each slide — ' +
      'without them you cannot know the slide layout. Requires PowerPoint on Windows (16.0.17628+), Mac (16.85+), or PowerPoint on the web for images.',
    params: {
      thumbnailWidth: {
        type: 'number',
        required: false,
        description: 'Width in pixels for slide thumbnails. Default: 600.',
      },
    },
    execute: async (context, args) => {
      const { thumbnailWidth = 600 } = args as { thumbnailWidth?: number };

      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const slideCount = slides.items.length;
      if (slideCount === 0) return 'Presentation has no slides.';

      const textLines = await loadSlideTexts(slides, 0, slideCount - 1, context);

      // Capture PNG thumbnail for every slide
      interface SlideWithImage {
        getImageAsBase64(width: number): { value: string };
      }
      const imageResults: { value: string }[] = [];
      let imagesSupported = true;
      try {
        for (const slide of slides.items) {
          imageResults.push((slide as unknown as SlideWithImage).getImageAsBase64(thumbnailWidth));
        }
        await context.sync();
      } catch {
        imagesSupported = false;
      }

      const overview = [
        `Presentation Overview`,
        `${'='.repeat(40)}`,
        `Total slides: ${String(slideCount)}`,
        ``,
        ...textLines,
      ].join('\n');

      if (!imagesSupported || imageResults.length === 0) {
        return overview;
      }

      return {
        text: overview,
        slides: imageResults.map((r, i) => ({
          slideNumber: i + 1,
          image: `data:image/png;base64,${r.value}`,
        })),
      };
    },
  },

  {
    name: 'get_presentation_content',
    description:
      'Get the text content of one or more slides. Specify slideIndex for a single slide, startIndex/endIndex for a range, or omit all to read every slide.',
    params: {
      slideIndex: {
        type: 'number',
        required: false,
        description: '0-based index for reading a single slide. Omit to use range or read all.',
      },
      startIndex: {
        type: 'number',
        required: false,
        description: '0-based start index for a range of slides.',
      },
      endIndex: {
        type: 'number',
        required: false,
        description: '0-based end index (inclusive) for a range of slides.',
      },
    },
    execute: async (context, args) => {
      const { slideIndex, startIndex, endIndex } = args as {
        slideIndex?: number;
        startIndex?: number;
        endIndex?: number;
      };

      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const slideCount = slides.items.length;
      if (slideCount === 0) return 'Presentation has no slides.';

      let start: number;
      let end: number;

      if (slideIndex !== undefined) {
        if (slideIndex < 0 || slideIndex >= slideCount) {
          throw new Error(
            `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`
          );
        }
        start = slideIndex;
        end = slideIndex;
      } else if (startIndex !== undefined && endIndex !== undefined) {
        start = Math.max(0, startIndex);
        end = Math.min(slideCount - 1, endIndex);
        if (start > end) {
          throw new Error(
            `Invalid range: startIndex (${String(startIndex)}) must be <= endIndex (${String(endIndex)}).`
          );
        }
      } else {
        start = 0;
        end = slideCount - 1;
      }

      const lines = await loadSlideTexts(slides, start, end, context);
      return lines.join('\n');
    },
  },

  {
    name: 'get_slide_image',
    description:
      'Capture a slide as a PNG image to see its visual design, layout, colors, and styling. ' +
      'Use region="full" for an overview, or "bottom-left"/"bottom-right" for 2× zoom into the bottom corners to check text overflow. ' +
      'Requires PowerPoint on Windows (16.0.17628+), Mac (16.85+), or PowerPoint on the web.',
    params: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
      width: {
        type: 'number',
        required: false,
        description: 'Image width in pixels. Aspect ratio is preserved. Default: 800.',
      },
      region: {
        type: 'string',
        required: false,
        description:
          'Which part of the slide to return. "full" (default) = entire slide. ' +
          '"top-left", "top-right", "bottom-left", "bottom-right" = that quadrant zoomed 2×. ' +
          'Use bottom quadrants to inspect text overflow at the bottom of slides.',
        enum: ['full', 'top-left', 'top-right', 'bottom-left', 'bottom-right'],
        default: 'full',
      },
    },
    execute: async (context, args) => {
      const {
        slideIndex,
        width = 800,
        region = 'full',
      } = args as {
        slideIndex: number;
        width?: number;
        region?: 'full' | 'top-left' | 'top-right' | 'bottom-left' | 'bottom-right';
      };

      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const slideCount = slides.items.length;
      if (slideIndex < 0 || slideIndex >= slideCount) {
        throw new Error(
          `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`
        );
      }

      const slide = slides.items[slideIndex];
      // getImageAsBase64 is available in PowerPoint requirement set 1.5+
      interface SlideWithImage {
        getImageAsBase64(width: number): { value: string };
      }

      let imageResult: { value: string };
      try {
        imageResult = (slide as unknown as SlideWithImage).getImageAsBase64(width);
        await context.sync();
      } catch {
        return (
          'Slide image capture is not available in this version of PowerPoint. ' +
          'Image capture requires PowerPoint on Windows (16.0.17628+), Mac (16.85+), or PowerPoint on the web. ' +
          'Use get_slide_shapes and get_presentation_content to inspect the slide via text instead.'
        );
      }

      const fullDataUrl = `data:image/png;base64,${imageResult.value}`;

      if (region === 'full') {
        return fullDataUrl;
      }

      // Crop to the requested quadrant and scale up 2× using an offscreen canvas.
      // This gives the model a zoomed view to check for text overflow / layout issues.
      return new Promise<string>((resolve, reject) => {
        const img = new Image();
        img.onload = () => {
          const halfW = Math.floor(img.width / 2);
          const halfH = Math.floor(img.height / 2);

          let sx: number;
          let sy: number;
          switch (region) {
            case 'top-left':
              sx = 0;
              sy = 0;
              break;
            case 'top-right':
              sx = halfW;
              sy = 0;
              break;
            case 'bottom-left':
              sx = 0;
              sy = halfH;
              break;
            case 'bottom-right':
              sx = halfW;
              sy = halfH;
              break;
            default:
              sx = 0;
              sy = 0;
          }

          const canvas = document.createElement('canvas');
          canvas.width = img.width; // output is full-size — the quadrant is stretched 2×
          canvas.height = img.height;
          const ctx = canvas.getContext('2d');
          if (!ctx) {
            resolve(fullDataUrl); // canvas not available — fall back to full image
            return;
          }
          ctx.drawImage(img, sx, sy, halfW, halfH, 0, 0, img.width, img.height);
          resolve(canvas.toDataURL('image/png'));
        };
        img.onerror = () => {
          reject(new Error('Failed to load slide image for region crop'));
        };
        img.src = fullDataUrl;
      });
    },
  },

  {
    name: 'get_slide_notes',
    description:
      'Get speaker notes from a PowerPoint slide. Note: The notes API has limited support in web add-ins; notes may not be available in all environments.',
    params: {
      slideIndex: {
        type: 'number',
        required: false,
        description: '0-based slide index. Omit to get notes from all slides.',
      },
    },
    execute: async (context, args) => {
      const { slideIndex } = args as { slideIndex?: number };

      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const slideCount = slides.items.length;
      if (slideCount === 0) return 'Presentation has no slides.';

      if (slideIndex !== undefined && (slideIndex < 0 || slideIndex >= slideCount)) {
        throw new Error(
          `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`
        );
      }

      const startIdx = slideIndex ?? 0;
      const endIdx = slideIndex !== undefined ? slideIndex + 1 : slideCount;

      const results: string[] = [];
      for (let i = startIdx; i < endIdx; i++) {
        const slide = slides.items[i];
        let notesText = '(no notes)';
        try {
          /* eslint-disable @typescript-eslint/no-explicit-any, @typescript-eslint/no-unsafe-member-access, @typescript-eslint/no-unsafe-call, @typescript-eslint/no-unsafe-assignment */
          const notesObj = (slide as any).notes;
          if (notesObj?.body) {
            notesObj.body.load('text');
            await context.sync();
            notesText = (notesObj.body.text as string) || '(no notes)';
          }
          /* eslint-enable @typescript-eslint/no-explicit-any, @typescript-eslint/no-unsafe-member-access, @typescript-eslint/no-unsafe-call, @typescript-eslint/no-unsafe-assignment */
        } catch {
          notesText = '(notes unavailable in this environment)';
        }
        results.push(`Slide ${i + 1}: ${notesText}`);
      }

      return slideIndex !== undefined
        ? results[0]
        : `Speaker Notes\n${'='.repeat(40)}\n${results.join('\n')}`;
    },
  },

  {
    name: 'set_presentation_content',
    description:
      'Add a text box to a slide. Pass slideIndex equal to the current total slide count to add a new slide first.',
    params: {
      slideIndex: {
        type: 'number',
        description:
          '0-based slide index. Pass the total slide count to append a new slide before adding.',
      },
      text: { type: 'string', description: 'The text content to add.' },
    },
    execute: async (context, args) => {
      const { slideIndex, text } = args as { slideIndex: number; text: string };

      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      let slideCount = slides.items.length;
      if (slideIndex < 0 || slideIndex > slideCount) {
        throw new Error(
          `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount)}.`
        );
      }

      if (slideIndex === slideCount) {
        context.presentation.slides.add();
        await context.sync();
        slides.load('items');
        await context.sync();
        slideCount = slides.items.length;
      }

      const slide = slides.items[slideIndex];
      slide.shapes.addTextBox(text, { left: 50, top: 100, width: 600, height: 400 });
      await context.sync();

      return `Added text box to slide ${String(slideIndex + 1)}.`;
    },
  },

  {
    name: 'add_slide_from_code',
    description: `Add a richly formatted slide to the presentation using PptxGenJS code.
Provide a JavaScript function body that receives a 'slide' parameter (PptxGenJS Slide object).

PptxGenJS API reference:
- Text:   slide.addText("Hello", { x:1, y:1, w:8, h:1, fontSize:24, bold:true, color:"363636" })
- Bullets: slide.addText([{text:"Point 1",options:{bullet:true}},{text:"Point 2",options:{bullet:true}}], { x:0.5, y:1.5, w:9, h:3, fontSize:18 })
- Image (base64): slide.addImage({ data:"data:image/png;base64,...", x:1, y:1, w:4, h:3 })
- Table:  slide.addTable([["H1","H2"],["R1","R2"]], { x:0.5, y:2, w:9, fontSize:14 })
- Shape:  slide.addShape("rect", { x:1, y:1, w:3, h:1, fill:{ color:"FF0000" } })
- All positions (x, y, w, h) are in inches.`,
    params: {
      code: {
        type: 'string',
        description:
          "JavaScript code (function body) receiving a 'slide' parameter. Call PptxGenJS methods on it to build slide content.",
      },
      replaceSlideIndex: {
        type: 'number',
        required: false,
        description:
          'Optional 0-based index of an existing slide to replace. If omitted, the new slide is appended.',
      },
    },
    execute: async (context, args) => {
      const { code, replaceSlideIndex } = args as {
        code: string;
        replaceSlideIndex?: number;
      };

      // Build the pptxgenjs slide (pure JS, no Office context needed)
      const pptx = new pptxgen();
      const slide = pptx.addSlide();

      /* eslint-disable @typescript-eslint/no-implied-eval, @typescript-eslint/no-unsafe-call */
      const buildSlide = new Function('slide', code);
      buildSlide(slide);
      /* eslint-enable @typescript-eslint/no-implied-eval, @typescript-eslint/no-unsafe-call */

      const base64 = (await pptx.write({ outputType: 'base64' })) as string;

      // Insert into presentation using the PowerPoint context
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const slideCount = slides.items.length;

      const insertOptions: PowerPoint.InsertSlideOptions = {
        formatting: PowerPoint.InsertSlideFormatting.useDestinationTheme,
      };

      if (replaceSlideIndex !== undefined) {
        if (replaceSlideIndex < 0 || replaceSlideIndex >= slideCount) {
          throw new Error(
            `Invalid replaceSlideIndex ${String(replaceSlideIndex)}. Must be 0-${String(slideCount - 1)}.`
          );
        }
        if (replaceSlideIndex > 0) {
          const prevSlide = slides.items[replaceSlideIndex - 1];
          prevSlide.load('id');
          await context.sync();
          insertOptions.targetSlideId = prevSlide.id;
        }
      } else if (slideCount > 0) {
        const lastSlide = slides.items[slideCount - 1];
        lastSlide.load('id');
        await context.sync();
        insertOptions.targetSlideId = lastSlide.id;
      }

      context.presentation.insertSlidesFromBase64(base64, insertOptions);
      await context.sync();

      if (replaceSlideIndex !== undefined) {
        // After insert the old slide has shifted by 1
        slides.load('items');
        await context.sync();
        const oldSlide = slides.items[replaceSlideIndex + 1];
        oldSlide.delete();
        await context.sync();
      }

      return replaceSlideIndex !== undefined
        ? `Successfully replaced slide ${String(replaceSlideIndex + 1)}.`
        : 'Successfully added new slide to the presentation.';
    },
  },

  {
    name: 'clear_slide',
    description: 'Remove all shapes from a specific slide, leaving it blank.',
    params: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
    },
    execute: async (context, args) => {
      const { slideIndex } = args as { slideIndex: number };

      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const slideCount = slides.items.length;
      if (slideIndex < 0 || slideIndex >= slideCount) {
        throw new Error(
          `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`
        );
      }

      const slide = slides.items[slideIndex];
      const shapes = slide.shapes;
      shapes.load('items');
      await context.sync();

      for (const shape of shapes.items) {
        shape.delete();
      }
      await context.sync();

      return `Cleared all shapes from slide ${String(slideIndex + 1)}.`;
    },
  },

  {
    name: 'update_slide_shape',
    description: 'Update the text content of an existing shape on a PowerPoint slide.',
    params: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
      shapeIndex: { type: 'number', description: '0-based shape index within the slide.' },
      text: { type: 'string', description: 'The new text content for the shape.' },
    },
    execute: async (context, args) => {
      const { slideIndex, shapeIndex, text } = args as {
        slideIndex: number;
        shapeIndex: number;
        text: string;
      };

      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const slideCount = slides.items.length;
      if (slideIndex < 0 || slideIndex >= slideCount) {
        throw new Error(
          `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`
        );
      }

      const slide = slides.items[slideIndex];
      const shapes = slide.shapes;
      shapes.load('items');
      await context.sync();

      const shapeCount = shapes.items.length;
      if (shapeIndex < 0 || shapeIndex >= shapeCount) {
        throw new Error(
          `Invalid shapeIndex ${String(shapeIndex)}. Slide ${String(slideIndex + 1)} has ${String(shapeCount)} shape(s).`
        );
      }

      const shape = shapes.items[shapeIndex];
      shape.textFrame.textRange.load('text');
      await context.sync();

      shape.textFrame.textRange.text = text;
      await context.sync();

      return `Updated shape ${String(shapeIndex + 1)} on slide ${String(slideIndex + 1)}.`;
    },
  },

  {
    name: 'set_slide_notes',
    description:
      'Add or update speaker notes for a slide. Due to API limitations, this provides guidance rather than directly modifying notes in all environments.',
    params: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
      notes: { type: 'string', description: 'The speaker notes text.' },
    },
    execute: async (context, args) => {
      const { slideIndex, notes } = args as { slideIndex: number; notes: string };

      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const slideCount = slides.items.length;
      if (slideIndex < 0 || slideIndex >= slideCount) {
        throw new Error(
          `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`
        );
      }

      const slide = slides.items[slideIndex];
      try {
        /* eslint-disable @typescript-eslint/no-explicit-any, @typescript-eslint/no-unsafe-member-access, @typescript-eslint/no-unsafe-assignment */
        const notesObj = (slide as any).notes;
        if (!notesObj?.body) throw new Error('Notes API not available');
        notesObj.body.text = notes;
        /* eslint-enable @typescript-eslint/no-explicit-any, @typescript-eslint/no-unsafe-member-access, @typescript-eslint/no-unsafe-assignment */
        await context.sync();
        return `Set speaker notes for slide ${String(slideIndex + 1)}.`;
      } catch {
        const preview = notes.length > 100 ? `${notes.substring(0, 100)}...` : notes;
        return `Notes API unavailable in this environment. For slide ${String(slideIndex + 1)}, please use the Notes pane in PowerPoint to add: "${preview}"`;
      }
    },
  },

  {
    name: 'duplicate_slide',
    description:
      'Duplicate an existing slide by copying its text content into a new slide. Note: Only text shapes are copied; complex graphics may not be preserved.',
    params: {
      sourceIndex: { type: 'number', description: '0-based index of the slide to duplicate.' },
    },
    execute: async (context, args) => {
      const { sourceIndex } = args as { sourceIndex: number };

      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const slideCount = slides.items.length;
      if (slideCount === 0) return 'Presentation has no slides.';
      if (sourceIndex < 0 || sourceIndex >= slideCount) {
        throw new Error(
          `Invalid sourceIndex ${String(sourceIndex)}. Must be 0-${String(slideCount - 1)}.`
        );
      }

      // Load source slide shapes and text
      const sourceSlide = slides.items[sourceIndex];
      sourceSlide.shapes.load('items');
      await context.sync();

      for (const shape of sourceSlide.shapes.items) {
        try {
          shape.textFrame.textRange.load('text');
        } catch {
          // not all shapes have a textFrame
        }
      }
      await context.sync();

      // Add new blank slide and copy text shapes
      slides.add();
      await context.sync();
      slides.load('items');
      await context.sync();
      const newSlide = slides.items[slides.items.length - 1];

      for (const shape of sourceSlide.shapes.items) {
        try {
          const text = shape.textFrame.textRange.text ?? '';
          if (text) {
            newSlide.shapes.addTextBox(text, { left: 50, top: 100, width: 600, height: 100 });
          }
        } catch {
          // skip non-text shapes
        }
      }
      await context.sync();

      return `Duplicated slide ${String(sourceIndex + 1)} (text content only).`;
    },
  },

  {
    name: 'get_selected_slides',
    description:
      'Get the currently selected slide(s). Call this first to know which slide the user is working on.',
    params: {},
    execute: async context => {
      const allSlides = context.presentation.slides;
      allSlides.load('items');
      await context.sync();
      for (const s of allSlides.items) {
        s.load('id');
      }
      await context.sync();

      const selected = context.presentation.getSelectedSlides();
      selected.load('items');
      await context.sync();
      for (const s of selected.items) {
        s.load('id');
      }
      await context.sync();

      const selectedIds = new Set(selected.items.map(s => s.id));
      const matches = allSlides.items
        .map((s, i) => ({ index: i, id: s.id }))
        .filter(s => selectedIds.has(s.id));

      if (matches.length === 0) return 'No slides currently selected.';
      return `Selected slide(s): ${matches.map(r => `Slide ${r.index + 1} (index ${r.index})`).join(', ')}`;
    },
  },

  {
    name: 'get_slide_shapes',
    description:
      'List all shapes on a slide with their index, name, type, position (inches), size (inches), and text. Use this before modifying, moving, or deleting individual shapes.',
    params: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
    },
    execute: async (context, args) => {
      const { slideIndex } = args as { slideIndex: number };

      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const slideCount = slides.items.length;
      if (slideIndex < 0 || slideIndex >= slideCount) {
        throw new Error(
          `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`
        );
      }

      const slide = slides.items[slideIndex];
      slide.shapes.load('items');
      await context.sync();

      if (slide.shapes.items.length === 0) {
        return `Slide ${String(slideIndex + 1)} has no shapes.`;
      }

      for (const shape of slide.shapes.items) {
        shape.load('name,id,left,top,width,height,type');
        try {
          shape.textFrame.textRange.load('text');
        } catch {
          // shape may not have a textFrame
        }
      }
      await context.sync();

      const PTS_PER_INCH = 72;
      const lines = slide.shapes.items.map((shape, i) => {
        const x = (shape.left / PTS_PER_INCH).toFixed(2);
        const y = (shape.top / PTS_PER_INCH).toFixed(2);
        const w = (shape.width / PTS_PER_INCH).toFixed(2);
        const h = (shape.height / PTS_PER_INCH).toFixed(2);
        let text = '';
        try {
          text = shape.textFrame.textRange.text?.trim() ?? '';
        } catch {
          /* no text frame */
        }
        const textPart = text
          ? ` | text: "${text.length > 60 ? `${text.substring(0, 60)}\u2026` : text}"`
          : '';
        return `[${i}] "${shape.name}" type:${String(shape.type)} — x:${x}" y:${y}" w:${w}" h:${h}"${textPart}`;
      });

      return `Slide ${String(slideIndex + 1)} — ${String(slide.shapes.items.length)} shape(s):\n${lines.join('\n')}`;
    },
  },

  {
    name: 'get_slide_layouts',
    description:
      'List all available slide layouts from the first slide master. Use before apply_slide_layout.',
    params: {},
    execute: async context => {
      const masters = context.presentation.slideMasters;
      masters.load('items');
      await context.sync();

      if (masters.items.length === 0) return 'No slide masters found.';

      const master = masters.items[0];
      master.layouts.load('items');
      await context.sync();

      for (const l of master.layouts.items) {
        l.load('name');
      }
      await context.sync();

      if (master.layouts.items.length === 0) return 'No slide layouts found.';
      const lines = master.layouts.items.map((l, i) => `[${i}] ${l.name}`);
      return `Available layouts (${String(master.layouts.items.length)}):\n${lines.join('\n')}`;
    },
  },

  {
    name: 'delete_slide',
    description: 'Delete a slide from the presentation by its 0-based index.',
    params: {
      slideIndex: { type: 'number', description: '0-based index of the slide to delete.' },
    },
    execute: async (context, args) => {
      const { slideIndex } = args as { slideIndex: number };

      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const slideCount = slides.items.length;
      if (slideCount === 0) throw new Error('Presentation has no slides.');
      if (slideIndex < 0 || slideIndex >= slideCount) {
        throw new Error(
          `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`
        );
      }

      slides.items[slideIndex].delete();
      await context.sync();

      return `Deleted slide ${String(slideIndex + 1)}.`;
    },
  },

  {
    name: 'move_slide',
    description:
      'Reorder a slide by moving it to a new position. Requires PowerPoint 16.0.14326+ (requirement set 1.8).',
    params: {
      fromIndex: { type: 'number', description: '0-based index of the slide to move.' },
      toIndex: {
        type: 'number',
        description: '0-based destination index the slide should occupy after the move.',
      },
    },
    execute: async (context, args) => {
      const { fromIndex, toIndex } = args as { fromIndex: number; toIndex: number };

      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const slideCount = slides.items.length;
      if (fromIndex < 0 || fromIndex >= slideCount) {
        throw new Error(
          `Invalid fromIndex ${String(fromIndex)}. Must be 0-${String(slideCount - 1)}.`
        );
      }
      if (toIndex < 0 || toIndex >= slideCount) {
        throw new Error(`Invalid toIndex ${String(toIndex)}. Must be 0-${String(slideCount - 1)}.`);
      }
      if (fromIndex === toIndex)
        return `Slide ${String(fromIndex + 1)} is already at that position.`;

      try {
        /* eslint-disable @typescript-eslint/no-explicit-any, @typescript-eslint/no-unsafe-call, @typescript-eslint/no-unsafe-member-access */
        (slides.items[fromIndex] as any).moveTo(toIndex);
        /* eslint-enable @typescript-eslint/no-explicit-any, @typescript-eslint/no-unsafe-call, @typescript-eslint/no-unsafe-member-access */
        await context.sync();
        return `Moved slide ${String(fromIndex + 1)} to position ${String(toIndex + 1)}.`;
      } catch {
        throw new Error(
          'move_slide requires PowerPoint 16.0.14326+ (requirement set 1.8). Alternatively use delete_slide + add_slide_from_code to recreate the slide at the desired position.'
        );
      }
    },
  },

  {
    name: 'delete_shape',
    description: 'Delete a specific shape from a slide by its index. Use get_slide_shapes first.',
    params: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
      shapeIndex: { type: 'number', description: '0-based shape index (from get_slide_shapes).' },
    },
    execute: async (context, args) => {
      const { slideIndex, shapeIndex } = args as { slideIndex: number; shapeIndex: number };

      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const slideCount = slides.items.length;
      if (slideIndex < 0 || slideIndex >= slideCount) {
        throw new Error(
          `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`
        );
      }

      const slide = slides.items[slideIndex];
      slide.shapes.load('items');
      await context.sync();

      const shapeCount = slide.shapes.items.length;
      if (shapeIndex < 0 || shapeIndex >= shapeCount) {
        throw new Error(
          `Invalid shapeIndex ${String(shapeIndex)}. Slide ${String(slideIndex + 1)} has ${String(shapeCount)} shape(s).`
        );
      }

      const shape = slide.shapes.items[shapeIndex];
      shape.load('name');
      await context.sync();
      const shapeName = shape.name;
      shape.delete();
      await context.sync();

      return `Deleted shape [${String(shapeIndex)}] "${shapeName}" from slide ${String(slideIndex + 1)}.`;
    },
  },

  {
    name: 'set_shape_text',
    description:
      'Set the text content of a shape by name (preferred) or by index. Use get_slide_shapes first to identify names and indices.',
    params: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
      text: { type: 'string', description: 'New text content for the shape.' },
      shapeName: {
        type: 'string',
        required: false,
        description: 'Name of the shape to update (preferred over shapeIndex).',
      },
      shapeIndex: {
        type: 'number',
        required: false,
        description: '0-based shape index. Used when shapeName is not provided.',
      },
    },
    execute: async (context, args) => {
      const { slideIndex, text, shapeName, shapeIndex } = args as {
        slideIndex: number;
        text: string;
        shapeName?: string;
        shapeIndex?: number;
      };

      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const slideCount = slides.items.length;
      if (slideIndex < 0 || slideIndex >= slideCount) {
        throw new Error(
          `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`
        );
      }

      const slide = slides.items[slideIndex];
      slide.shapes.load('items');
      await context.sync();
      for (const s of slide.shapes.items) {
        s.load('name');
      }
      await context.sync();

      let targetIndex = -1;
      if (shapeName) {
        targetIndex = slide.shapes.items.findIndex(
          s => s.name.toLowerCase() === shapeName.toLowerCase()
        );
        if (targetIndex === -1) {
          const available = slide.shapes.items.map(s => `"${s.name}"`).join(', ');
          throw new Error(
            `Shape named "${shapeName}" not found on slide ${String(slideIndex + 1)}. Available: ${available}`
          );
        }
      } else if (shapeIndex !== undefined) {
        if (shapeIndex < 0 || shapeIndex >= slide.shapes.items.length) {
          throw new Error(
            `Invalid shapeIndex ${String(shapeIndex)}. Slide has ${String(slide.shapes.items.length)} shape(s).`
          );
        }
        targetIndex = shapeIndex;
      } else {
        throw new Error('Provide either shapeName or shapeIndex.');
      }

      const shape = slide.shapes.items[targetIndex];
      shape.textFrame.textRange.text = text;
      await context.sync();

      return `Updated text of shape "${shape.name}" on slide ${String(slideIndex + 1)}.`;
    },
  },

  {
    name: 'move_resize_shape',
    description:
      'Move and/or resize a shape on a slide. All values are in inches (matching add_slide_from_code). Omit any property to leave it unchanged.',
    params: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
      shapeIndex: { type: 'number', description: '0-based shape index (from get_slide_shapes).' },
      left: {
        type: 'number',
        required: false,
        description: 'New left position in inches from the left edge of the slide.',
      },
      top: {
        type: 'number',
        required: false,
        description: 'New top position in inches from the top edge of the slide.',
      },
      width: { type: 'number', required: false, description: 'New width in inches.' },
      height: { type: 'number', required: false, description: 'New height in inches.' },
    },
    execute: async (context, args) => {
      const { slideIndex, shapeIndex, left, top, width, height } = args as {
        slideIndex: number;
        shapeIndex: number;
        left?: number;
        top?: number;
        width?: number;
        height?: number;
      };

      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const slideCount = slides.items.length;
      if (slideIndex < 0 || slideIndex >= slideCount) {
        throw new Error(
          `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`
        );
      }

      const slide = slides.items[slideIndex];
      slide.shapes.load('items');
      await context.sync();

      const shapeCount = slide.shapes.items.length;
      if (shapeIndex < 0 || shapeIndex >= shapeCount) {
        throw new Error(
          `Invalid shapeIndex ${String(shapeIndex)}. Slide ${String(slideIndex + 1)} has ${String(shapeCount)} shape(s).`
        );
      }

      const PTS_PER_INCH = 72;
      const shape = slide.shapes.items[shapeIndex];
      if (left !== undefined) shape.left = left * PTS_PER_INCH;
      if (top !== undefined) shape.top = top * PTS_PER_INCH;
      if (width !== undefined) shape.width = width * PTS_PER_INCH;
      if (height !== undefined) shape.height = height * PTS_PER_INCH;
      await context.sync();

      const changes: string[] = [];
      if (left !== undefined) changes.push(`left:${String(left)}"`);
      if (top !== undefined) changes.push(`top:${String(top)}"`);
      if (width !== undefined) changes.push(`width:${String(width)}"`);
      if (height !== undefined) changes.push(`height:${String(height)}"`);
      return `Updated shape [${String(shapeIndex)}] on slide ${String(slideIndex + 1)}: ${changes.join(', ')}.`;
    },
  },

  {
    name: 'update_shape_style',
    description:
      'Update the visual style of a shape: fill color, font color, font size, bold. Use get_slide_shapes to get the shapeIndex first.',
    params: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
      shapeIndex: { type: 'number', description: '0-based shape index (from get_slide_shapes).' },
      fillColor: {
        type: 'string',
        required: false,
        description: '6-digit hex fill color without # (e.g. "4472C4"). Use "none" to remove fill.',
      },
      fontColor: {
        type: 'string',
        required: false,
        description: '6-digit hex font color without # (e.g. "FFFFFF").',
      },
      fontSize: { type: 'number', required: false, description: 'Font size in points.' },
      bold: { type: 'boolean', required: false, description: 'Set text bold (true/false).' },
    },
    execute: async (context, args) => {
      const { slideIndex, shapeIndex, fillColor, fontColor, fontSize, bold } = args as {
        slideIndex: number;
        shapeIndex: number;
        fillColor?: string;
        fontColor?: string;
        fontSize?: number;
        bold?: boolean;
      };

      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const slideCount = slides.items.length;
      if (slideIndex < 0 || slideIndex >= slideCount) {
        throw new Error(
          `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`
        );
      }

      const slide = slides.items[slideIndex];
      slide.shapes.load('items');
      await context.sync();

      const shapeCount = slide.shapes.items.length;
      if (shapeIndex < 0 || shapeIndex >= shapeCount) {
        throw new Error(
          `Invalid shapeIndex ${String(shapeIndex)}. Slide ${String(slideIndex + 1)} has ${String(shapeCount)} shape(s).`
        );
      }

      const shape = slide.shapes.items[shapeIndex];
      const applied: string[] = [];

      if (fillColor !== undefined) {
        const hex = fillColor.startsWith('#') ? fillColor.slice(1) : fillColor;
        if (hex.toLowerCase() === 'none') {
          shape.fill.clear();
        } else {
          shape.fill.setSolidColor(hex);
        }
        applied.push(`fill:${hex.toLowerCase() === 'none' ? 'none' : `#${hex}`}`);
      }

      if (fontColor !== undefined || fontSize !== undefined || bold !== undefined) {
        const font = shape.textFrame.textRange.font;
        if (fontColor !== undefined) {
          const hex = fontColor.startsWith('#') ? fontColor.slice(1) : fontColor;
          font.color = hex;
          applied.push(`fontColor:#${hex}`);
        }
        if (fontSize !== undefined) {
          font.size = fontSize;
          applied.push(`fontSize:${String(fontSize)}pt`);
        }
        if (bold !== undefined) {
          font.bold = bold;
          applied.push(`bold:${String(bold)}`);
        }
      }

      if (applied.length === 0) throw new Error('Provide at least one style property to update.');

      await context.sync();
      return `Updated style of shape [${String(shapeIndex)}] on slide ${String(slideIndex + 1)}: ${applied.join(', ')}.`;
    },
  },

  {
    name: 'set_slide_background',
    description: 'Set the solid background color of a slide.',
    params: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
      color: {
        type: 'string',
        description:
          '6-digit hex color without # (e.g. "1F2937" for dark charcoal). Use "none" to reset to theme default.',
      },
    },
    execute: async (context, args) => {
      const { slideIndex, color } = args as { slideIndex: number; color: string };

      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const slideCount = slides.items.length;
      if (slideIndex < 0 || slideIndex >= slideCount) {
        throw new Error(
          `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`
        );
      }

      const slide = slides.items[slideIndex];
      const hex = color.startsWith('#') ? color.slice(1) : color;
      if (hex.toLowerCase() === 'none') {
        slide.background.reset();
      } else {
        slide.background.fill.setSolidFill({ color: hex });
      }
      await context.sync();

      return `Set background of slide ${String(slideIndex + 1)} to ${hex.toLowerCase() === 'none' ? 'theme default' : `#${hex}`}.`;
    },
  },

  {
    name: 'apply_slide_layout',
    description:
      'Apply a slide layout to a slide by name or index. Use get_slide_layouts to see available layouts first.',
    params: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
      layoutName: {
        type: 'string',
        required: false,
        description: 'Name of the layout to apply (from get_slide_layouts). Preferred.',
      },
      layoutIndex: {
        type: 'number',
        required: false,
        description: '0-based layout index. Used when layoutName is not provided.',
      },
    },
    execute: async (context, args) => {
      const { slideIndex, layoutName, layoutIndex } = args as {
        slideIndex: number;
        layoutName?: string;
        layoutIndex?: number;
      };

      if (layoutName === undefined && layoutIndex === undefined) {
        throw new Error('Provide either layoutName or layoutIndex.');
      }

      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const slideCount = slides.items.length;
      if (slideIndex < 0 || slideIndex >= slideCount) {
        throw new Error(
          `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`
        );
      }

      const masters = context.presentation.slideMasters;
      masters.load('items');
      await context.sync();
      if (masters.items.length === 0) throw new Error('No slide masters found.');

      const master = masters.items[0];
      master.layouts.load('items');
      await context.sync();
      for (const l of master.layouts.items) {
        l.load('name');
      }
      await context.sync();

      let targetLayout: PowerPoint.SlideLayout | undefined;
      if (layoutName) {
        targetLayout = master.layouts.items.find(
          l => l.name.toLowerCase() === layoutName.toLowerCase()
        );
      } else if (layoutIndex !== undefined) {
        targetLayout = master.layouts.items[layoutIndex];
      }

      if (!targetLayout) {
        const available = master.layouts.items.map((l, i) => `[${i}] ${l.name}`).join(', ');
        throw new Error(`Layout not found. Available: ${available}`);
      }

      const layoutFoundName = targetLayout.name;
      try {
        /* eslint-disable @typescript-eslint/no-explicit-any, @typescript-eslint/no-unsafe-member-access */
        (slides.items[slideIndex] as any).layout = targetLayout;
        /* eslint-enable @typescript-eslint/no-explicit-any, @typescript-eslint/no-unsafe-member-access */
        await context.sync();
        return `Applied layout "${layoutFoundName}" to slide ${String(slideIndex + 1)}.`;
      } catch {
        throw new Error(
          `Failed to apply layout "${layoutFoundName}". This API may require a newer PowerPoint version.`
        );
      }
    },
  },

  {
    name: 'add_geometric_shape',
    description:
      'Add a geometric shape to a slide. Position and size are in inches, consistent with add_slide_from_code.',
    params: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
      shapeType: {
        type: 'string',
        description:
          'Shape type — e.g. "rectangle", "roundedRectangle", "ellipse", "triangle", "rightTriangle", "diamond", "pentagon", "hexagon", "star5", "heart", "cloud", "arrowRight", "arrowLeft". See PowerPoint.GeometricShapeType for full list.',
      },
      left: { type: 'number', description: 'Left position in inches.' },
      top: { type: 'number', description: 'Top position in inches.' },
      width: { type: 'number', description: 'Width in inches.' },
      height: { type: 'number', description: 'Height in inches.' },
      fillColor: {
        type: 'string',
        required: false,
        description: '6-digit hex fill color without # (e.g. "4472C4"). Optional.',
      },
      name: { type: 'string', required: false, description: 'Optional name for the shape.' },
    },
    execute: async (context, args) => {
      const { slideIndex, shapeType, left, top, width, height, fillColor, name } = args as {
        slideIndex: number;
        shapeType: string;
        left: number;
        top: number;
        width: number;
        height: number;
        fillColor?: string;
        name?: string;
      };

      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const slideCount = slides.items.length;
      if (slideIndex < 0 || slideIndex >= slideCount) {
        throw new Error(
          `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`
        );
      }

      const slide = slides.items[slideIndex];
      const PTS_PER_INCH = 72;

      const options: PowerPoint.ShapeAddOptions = {
        left: left * PTS_PER_INCH,
        top: top * PTS_PER_INCH,
        width: width * PTS_PER_INCH,
        height: height * PTS_PER_INCH,
      };

      const shape = slide.shapes.addGeometricShape(
        shapeType as PowerPoint.GeometricShapeType,
        options
      );

      if (name) shape.name = name;

      if (fillColor) {
        const hex = fillColor.startsWith('#') ? fillColor.slice(1) : fillColor;
        shape.fill.setSolidColor(hex);
      }

      await context.sync();
      return `Added ${shapeType} shape${name ? ` "${name}"` : ''} to slide ${String(slideIndex + 1)} at (${String(left)}", ${String(top)}") ${String(width)}"\u00d7${String(height)}".`;
    },
  },

  {
    name: 'add_line',
    description:
      'Add a straight line (connector) to a slide. Coordinates are in inches, consistent with add_slide_from_code.',
    params: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
      startX: { type: 'number', description: 'Start X position in inches.' },
      startY: { type: 'number', description: 'Start Y position in inches.' },
      endX: { type: 'number', description: 'End X position in inches.' },
      endY: { type: 'number', description: 'End Y position in inches.' },
      connectorType: {
        type: 'string',
        required: false,
        description: 'Connector type: "straight" (default), "elbow", or "curve".',
        enum: ['straight', 'elbow', 'curve'],
        default: 'straight',
      },
      color: {
        type: 'string',
        required: false,
        description: '6-digit hex line color without # (e.g. "363636"). Optional.',
      },
    },
    execute: async (context, args) => {
      const {
        slideIndex,
        startX,
        startY,
        endX,
        endY,
        connectorType = 'straight',
        color,
      } = args as {
        slideIndex: number;
        startX: number;
        startY: number;
        endX: number;
        endY: number;
        connectorType?: string;
        color?: string;
      };

      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const slideCount = slides.items.length;
      if (slideIndex < 0 || slideIndex >= slideCount) {
        throw new Error(
          `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`
        );
      }

      const slide = slides.items[slideIndex];
      const PTS_PER_INCH = 72;

      // Bounding box derived from start/end points
      const options: PowerPoint.ShapeAddOptions = {
        left: Math.min(startX, endX) * PTS_PER_INCH,
        top: Math.min(startY, endY) * PTS_PER_INCH,
        width: Math.abs(endX - startX) * PTS_PER_INCH,
        height: Math.abs(endY - startY) * PTS_PER_INCH,
      };

      const shape = slide.shapes.addLine(connectorType as PowerPoint.ConnectorType, options);

      if (color) {
        try {
          const hex = color.startsWith('#') ? color.slice(1) : color;
          /* eslint-disable @typescript-eslint/no-explicit-any, @typescript-eslint/no-unsafe-member-access */
          (shape as any).lineFormat.color = `#${hex}`;
          /* eslint-enable @typescript-eslint/no-explicit-any, @typescript-eslint/no-unsafe-member-access */
        } catch {
          // lineFormat not available in this environment — continue
        }
      }

      await context.sync();
      return `Added ${connectorType} line from (${String(startX)}", ${String(startY)}") to (${String(endX)}", ${String(endY)}") on slide ${String(slideIndex + 1)}.`;
    },
  },

  {
    name: 'group_shapes',
    description:
      'Group multiple shapes on a slide into a single group shape. Use get_slide_shapes first to identify the shape indices. ' +
      'Requires PowerPoint 16.0.17531+ (requirement set 1.8).',
    params: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
      shapeIndices: {
        type: 'number[]',
        description:
          'Array of 0-based shape indices (from get_slide_shapes) to group together. Must contain at least 2 indices.',
      },
      groupName: {
        type: 'string',
        required: false,
        description: 'Optional name for the resulting group shape.',
      },
    },
    execute: async (context, args) => {
      const { slideIndex, shapeIndices, groupName } = args as {
        slideIndex: number;
        shapeIndices: number[];
        groupName?: string;
      };

      if (!Array.isArray(shapeIndices) || shapeIndices.length < 2) {
        throw new Error('shapeIndices must be an array of at least 2 shape indices.');
      }

      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const slideCount = slides.items.length;
      if (slideIndex < 0 || slideIndex >= slideCount) {
        throw new Error(
          `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`
        );
      }

      const slide = slides.items[slideIndex];
      slide.shapes.load('items');
      await context.sync();

      const shapeCount = slide.shapes.items.length;
      for (const idx of shapeIndices) {
        if (idx < 0 || idx >= shapeCount) {
          throw new Error(
            `Invalid shapeIndex ${String(idx)}. Slide ${String(slideIndex + 1)} has ${String(shapeCount)} shape(s).`
          );
        }
      }

      // Load IDs for selected shapes
      const selectedShapes = shapeIndices.map(i => slide.shapes.items[i]);
      for (const s of selectedShapes) {
        s.load('id,name');
      }
      await context.sync();

      try {
        const shapeIds = selectedShapes.map(s => s.id);
        const groupShape = slide.shapes.addGroup(shapeIds);
        if (groupName) groupShape.name = groupName;
        groupShape.load('name,id');
        await context.sync();

        const names = selectedShapes.map(s => `"${s.name}"`).join(', ');
        return `Grouped ${String(shapeIndices.length)} shapes (${names}) into group "${groupShape.name}" on slide ${String(slideIndex + 1)}.`;
      } catch {
        throw new Error(
          'group_shapes requires PowerPoint 16.0.17531+ (requirement set 1.8). Ensure shapes are not already in a group and belong to the same slide.'
        );
      }
    },
  },

  {
    name: 'ungroup_shapes',
    description:
      'Ungroup a grouped shape on a slide, releasing its child shapes back to the slide. ' +
      'Use get_slide_shapes first to identify the group shape index (type will be "Group"). ' +
      'Requires PowerPoint 16.0.17531+ (requirement set 1.8).',
    params: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
      shapeIndex: {
        type: 'number',
        description:
          '0-based index of the group shape to ungroup (from get_slide_shapes, type should be "Group").',
      },
    },
    execute: async (context, args) => {
      const { slideIndex, shapeIndex } = args as { slideIndex: number; shapeIndex: number };

      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const slideCount = slides.items.length;
      if (slideIndex < 0 || slideIndex >= slideCount) {
        throw new Error(
          `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`
        );
      }

      const slide = slides.items[slideIndex];
      slide.shapes.load('items');
      await context.sync();

      const shapeCount = slide.shapes.items.length;
      if (shapeIndex < 0 || shapeIndex >= shapeCount) {
        throw new Error(
          `Invalid shapeIndex ${String(shapeIndex)}. Slide ${String(slideIndex + 1)} has ${String(shapeCount)} shape(s).`
        );
      }

      const shape = slide.shapes.items[shapeIndex];
      shape.load('name,type');
      await context.sync();

      if (shape.type !== PowerPoint.ShapeType.group) {
        throw new Error(
          `Shape [${String(shapeIndex)}] "${shape.name}" is not a group (type: ${String(shape.type)}). Use get_slide_shapes to find a shape with type "Group".`
        );
      }

      const groupName = shape.name;

      try {
        const shapeGroup = shape.group;

        // Load child shapes before ungrouping to report count
        shapeGroup.shapes.load('items');
        await context.sync();
        const childCount = shapeGroup.shapes.items.length;

        shapeGroup.ungroup();
        await context.sync();

        return `Ungrouped "${groupName}" on slide ${String(slideIndex + 1)}, releasing ${String(childCount)} shape(s).`;
      } catch {
        throw new Error(
          `Failed to ungroup shape "${groupName}". Ensure it is a valid group shape. group_shapes requires PowerPoint 16.0.17531+ (requirement set 1.8).`
        );
      }
    },
  },

  {
    name: 'get_smartart_info',
    description:
      'List all SmartArt and diagram shapes on a slide with their index, name, position, and size. ' +
      'Use this to inspect existing SmartArt graphics. Note: SmartArt content cannot be modified via the Office.js API — ' +
      'use add_slide_from_code with PptxGenJS to create SmartArt-like visuals programmatically.',
    params: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
    },
    execute: async (context, args) => {
      const { slideIndex } = args as { slideIndex: number };

      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const slideCount = slides.items.length;
      if (slideIndex < 0 || slideIndex >= slideCount) {
        throw new Error(
          `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`
        );
      }

      const slide = slides.items[slideIndex];
      slide.shapes.load('items');
      await context.sync();

      for (const shape of slide.shapes.items) {
        shape.load('name,type,left,top,width,height');
      }
      await context.sync();

      const PTS_PER_INCH = 72;
      const smartArtShapes = slide.shapes.items
        .map((shape, i) => ({ shape, i }))
        .filter(
          ({ shape }) =>
            shape.type === PowerPoint.ShapeType.smartArt ||
            shape.type === PowerPoint.ShapeType.diagram
        );

      if (smartArtShapes.length === 0) {
        return `Slide ${String(slideIndex + 1)} has no SmartArt or diagram shapes.\n\nTip: To create SmartArt-like visuals, use add_slide_from_code with PptxGenJS shapes and connectors.`;
      }

      const lines = smartArtShapes.map(({ shape, i }) => {
        const x = (shape.left / PTS_PER_INCH).toFixed(2);
        const y = (shape.top / PTS_PER_INCH).toFixed(2);
        const w = (shape.width / PTS_PER_INCH).toFixed(2);
        const h = (shape.height / PTS_PER_INCH).toFixed(2);
        return `[${i}] "${shape.name}" type:${String(shape.type)} — x:${x}" y:${y}" w:${w}" h:${h}"`;
      });

      return (
        `Slide ${String(slideIndex + 1)} — ${String(smartArtShapes.length)} SmartArt/diagram shape(s):\n${lines.join('\n')}\n\n` +
        `Note: SmartArt content cannot be modified via the Office.js API. To replace with editable content, delete the SmartArt shape and use add_slide_from_code to create a similar visual layout.`
      );
    },
  },
];

export const powerPointTools = createPptTools(powerPointConfigs);
