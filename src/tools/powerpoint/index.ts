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
      'Capture a slide as a PNG image to see its visual design, layout, colors, and styling. Requires PowerPoint on Windows (16.0.17628+), Mac (16.85+), or PowerPoint on the web.',
    params: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
      width: {
        type: 'number',
        required: false,
        description: 'Image width in pixels. Aspect ratio is preserved. Default: 800.',
      },
    },
    execute: async (context, args) => {
      const { slideIndex, width = 800 } = args as { slideIndex: number; width?: number };

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
      const imageResult = (slide as unknown as SlideWithImage).getImageAsBase64(width);
      await context.sync();

      return `data:image/png;base64,${imageResult.value}`;
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
        results.push(
          `Slide ${i + 1}: (Speaker notes require PowerPoint desktop — API limitation in web add-ins)`
        );
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

      const preview = notes.length > 100 ? `${notes.substring(0, 100)}...` : notes;
      return `Note: Direct speaker notes editing has limited API support in web add-ins. For slide ${String(slideIndex + 1)}, please use the Notes pane in PowerPoint to add: "${preview}"`;
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
];

export const powerPointTools = createPptTools(powerPointConfigs);
