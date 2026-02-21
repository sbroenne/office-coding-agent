import type { Tool, ToolResultObject } from '@github/copilot-sdk';
import pptxgen from 'pptxgenjs';

const CHUNK_SIZE = 10;

async function getSlideCount(): Promise<number> {
  return PowerPoint.run(async context => {
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();
    return slides.items.length;
  });
}

async function getSlideTextContent(startIdx: number, endIdx: number): Promise<string[]> {
  return PowerPoint.run(async context => {
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();

    const slideRefs = slides.items.slice(startIdx, endIdx + 1);
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
      results.push(
        `Slide ${startIdx + i + 1}: ${texts.length > 0 ? texts.join(' | ') : '(no text)'}`
      );
    }
    return results;
  });
}

const getPresentationOverview: Tool = {
  name: 'get_presentation_overview',
  description:
    "Get an overview of the entire PowerPoint presentation including total slide count and a text preview of each slide's shapes. Use this first to understand the presentation structure.",
  parameters: { type: 'object', properties: {}, required: [] },
  handler: async (): Promise<ToolResultObject | string> => {
    try {
      const slideCount = await getSlideCount();
      if (slideCount === 0) return 'Presentation has no slides.';

      const slideOverviews: string[] = [];
      for (let i = 0; i < slideCount; i += CHUNK_SIZE) {
        const chunkEnd = Math.min(i + CHUNK_SIZE - 1, slideCount - 1);
        const chunk = await getSlideTextContent(i, chunkEnd);
        slideOverviews.push(...chunk);
      }

      return `Presentation Overview\n${'='.repeat(40)}\nTotal slides: ${String(slideCount)}\n\n${slideOverviews.join('\n')}`;
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const getPresentationContent: Tool = {
  name: 'get_presentation_content',
  description:
    'Get the text content of one or more slides. Specify slideIndex for a single slide, startIndex/endIndex for a range, or omit all to read every slide.',
  parameters: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '0-based index for reading a single slide. Omit to use range or read all.',
      },
      startIndex: { type: 'number', description: '0-based start index for a range of slides.' },
      endIndex: {
        type: 'number',
        description: '0-based end index (inclusive) for a range of slides.',
      },
    },
    required: [],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { slideIndex, startIndex, endIndex } = (args ?? {}) as {
      slideIndex?: number;
      startIndex?: number;
      endIndex?: number;
    };
    try {
      const slideCount = await getSlideCount();
      if (slideCount === 0) return 'Presentation has no slides.';

      let start: number;
      let end: number;

      if (slideIndex !== undefined) {
        if (slideIndex < 0 || slideIndex >= slideCount) {
          return {
            textResultForLlm: `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`,
            resultType: 'failure',
            error: 'Invalid slideIndex',
            toolTelemetry: {},
          };
        }
        start = slideIndex;
        end = slideIndex;
      } else if (startIndex !== undefined && endIndex !== undefined) {
        start = Math.max(0, startIndex);
        end = Math.min(slideCount - 1, endIndex);
        if (start > end) {
          return {
            textResultForLlm: `Invalid range: startIndex (${String(startIndex)}) must be <= endIndex (${String(endIndex)}).`,
            resultType: 'failure',
            error: 'Invalid range',
            toolTelemetry: {},
          };
        }
      } else {
        start = 0;
        end = slideCount - 1;
      }

      const results: string[] = [];
      for (let i = start; i <= end; i += CHUNK_SIZE) {
        const chunkEnd = Math.min(i + CHUNK_SIZE - 1, end);
        const chunk = await getSlideTextContent(i, chunkEnd);
        results.push(...chunk);
      }
      return results.join('\n');
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const getSlideImage: Tool = {
  name: 'get_slide_image',
  description:
    'Capture a slide as a PNG image to see its visual design, layout, colors, and styling. Requires PowerPoint on Windows (16.0.17628+), Mac (16.85+), or PowerPoint on the web.',
  parameters: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
      width: {
        type: 'number',
        description: 'Image width in pixels. Aspect ratio is preserved. Default: 800.',
      },
    },
    required: ['slideIndex'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { slideIndex, width = 800 } = (args ?? {}) as { slideIndex: number; width?: number };
    try {
      return await PowerPoint.run(async context => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();

        const slideCount = slides.items.length;
        if (slideIndex < 0 || slideIndex >= slideCount) {
          return {
            textResultForLlm: `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`,
            resultType: 'failure',
            error: 'Invalid slideIndex',
            toolTelemetry: {},
          };
        }

        const slide = slides.items[slideIndex];
        // getImageAsBase64 is available in PowerPoint requirement set 1.5+
        interface SlideWithImage {
          getImageAsBase64(width: number): { value: string };
        }
        const imageResult = (slide as unknown as SlideWithImage).getImageAsBase64(width);
        await context.sync();

        return {
          textResultForLlm: `data:image/png;base64,${imageResult.value}`,
          resultType: 'success',
          toolTelemetry: {},
        };
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return {
        textResultForLlm: `Slide image capture failed: ${msg}. Ensure you are using PowerPoint on Windows (16.0.17628+), Mac (16.85+), or PowerPoint on the web.`,
        resultType: 'failure',
        error: msg,
        toolTelemetry: {},
      };
    }
  },
};

const getSlideNotes: Tool = {
  name: 'get_slide_notes',
  description:
    'Get speaker notes from a PowerPoint slide. Note: The notes API has limited support in web add-ins; notes may not be available in all environments.',
  parameters: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '0-based slide index. Omit to get notes from all slides.',
      },
    },
    required: [],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { slideIndex } = (args ?? {}) as { slideIndex?: number };
    try {
      return await PowerPoint.run(async context => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();

        const slideCount = slides.items.length;
        if (slideCount === 0) return 'Presentation has no slides.';

        if (slideIndex !== undefined && (slideIndex < 0 || slideIndex >= slideCount)) {
          return `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`;
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
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const setPresentationContent: Tool = {
  name: 'set_presentation_content',
  description:
    'Add a text box to a slide. Pass slideIndex equal to the current total slide count to add a new slide first.',
  parameters: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description:
          '0-based slide index. Pass the total slide count to append a new slide before adding.',
      },
      text: { type: 'string', description: 'The text content to add.' },
    },
    required: ['slideIndex', 'text'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { slideIndex, text } = (args ?? {}) as { slideIndex: number; text: string };
    try {
      return await PowerPoint.run(async context => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();

        let slideCount = slides.items.length;
        if (slideIndex < 0 || slideIndex > slideCount) {
          return {
            textResultForLlm: `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount)}.`,
            resultType: 'failure',
            error: 'Invalid slideIndex',
            toolTelemetry: {},
          };
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
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const addSlideFromCode: Tool = {
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
  parameters: {
    type: 'object',
    properties: {
      code: {
        type: 'string',
        description:
          "JavaScript code (function body) receiving a 'slide' parameter. Call PptxGenJS methods on it to build slide content.",
      },
      replaceSlideIndex: {
        type: 'number',
        description:
          'Optional 0-based index of an existing slide to replace. If omitted, the new slide is appended.',
      },
    },
    required: ['code'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { code, replaceSlideIndex } = (args ?? {}) as {
      code: string;
      replaceSlideIndex?: number;
    };

    const pptx = new pptxgen();
    const slide = pptx.addSlide();

    try {
      /* eslint-disable @typescript-eslint/no-implied-eval, @typescript-eslint/no-unsafe-call */
      const buildSlide = new Function('slide', code);
      buildSlide(slide);
      /* eslint-enable @typescript-eslint/no-implied-eval, @typescript-eslint/no-unsafe-call */
    } catch (codeError) {
      const msg = codeError instanceof Error ? codeError.message : String(codeError);
      return {
        textResultForLlm: `Code execution error: ${msg}`,
        resultType: 'failure',
        error: msg,
        toolTelemetry: {},
      };
    }

    let base64: string;
    try {
      base64 = (await pptx.write({ outputType: 'base64' })) as string;
    } catch (writeError) {
      const msg = writeError instanceof Error ? writeError.message : String(writeError);
      return {
        textResultForLlm: `Failed to generate slide: ${msg}`,
        resultType: 'failure',
        error: msg,
        toolTelemetry: {},
      };
    }

    try {
      await PowerPoint.run(async context => {
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
      });

      return replaceSlideIndex !== undefined
        ? `Successfully replaced slide ${String(replaceSlideIndex + 1)}.`
        : 'Successfully added new slide to the presentation.';
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const clearSlide: Tool = {
  name: 'clear_slide',
  description: 'Remove all shapes from a specific slide, leaving it blank.',
  parameters: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
    },
    required: ['slideIndex'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { slideIndex } = (args ?? {}) as { slideIndex: number };
    try {
      return await PowerPoint.run(async context => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();

        const slideCount = slides.items.length;
        if (slideIndex < 0 || slideIndex >= slideCount) {
          return {
            textResultForLlm: `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`,
            resultType: 'failure',
            error: 'Invalid slideIndex',
            toolTelemetry: {},
          };
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
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const updateSlideShape: Tool = {
  name: 'update_slide_shape',
  description: 'Update the text content of an existing shape on a PowerPoint slide.',
  parameters: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
      shapeIndex: { type: 'number', description: '0-based shape index within the slide.' },
      text: { type: 'string', description: 'The new text content for the shape.' },
    },
    required: ['slideIndex', 'shapeIndex', 'text'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { slideIndex, shapeIndex, text } = (args ?? {}) as {
      slideIndex: number;
      shapeIndex: number;
      text: string;
    };
    try {
      return await PowerPoint.run(async context => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();

        const slideCount = slides.items.length;
        if (slideIndex < 0 || slideIndex >= slideCount) {
          return {
            textResultForLlm: `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`,
            resultType: 'failure',
            error: 'Invalid slideIndex',
            toolTelemetry: {},
          };
        }

        const slide = slides.items[slideIndex];
        const shapes = slide.shapes;
        shapes.load('items');
        await context.sync();

        const shapeCount = shapes.items.length;
        if (shapeIndex < 0 || shapeIndex >= shapeCount) {
          return {
            textResultForLlm: `Invalid shapeIndex ${String(shapeIndex)}. Slide ${String(slideIndex + 1)} has ${String(shapeCount)} shape(s).`,
            resultType: 'failure',
            error: 'Invalid shapeIndex',
            toolTelemetry: {},
          };
        }

        const shape = shapes.items[shapeIndex];
        shape.textFrame.textRange.load('text');
        await context.sync();

        shape.textFrame.textRange.text = text;
        await context.sync();

        return `Updated shape ${String(shapeIndex + 1)} on slide ${String(slideIndex + 1)}.`;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const setSlideNotes: Tool = {
  name: 'set_slide_notes',
  description:
    'Add or update speaker notes for a slide. Due to API limitations, this provides guidance rather than directly modifying notes in all environments.',
  parameters: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '0-based slide index.' },
      notes: { type: 'string', description: 'The speaker notes text.' },
    },
    required: ['slideIndex', 'notes'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { slideIndex, notes } = (args ?? {}) as { slideIndex: number; notes: string };
    try {
      return await PowerPoint.run(async context => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();

        const slideCount = slides.items.length;
        if (slideIndex < 0 || slideIndex >= slideCount) {
          return `Invalid slideIndex ${String(slideIndex)}. Must be 0-${String(slideCount - 1)}.`;
        }

        const preview = notes.length > 100 ? `${notes.substring(0, 100)}...` : notes;
        return `Note: Direct speaker notes editing has limited API support in web add-ins. For slide ${String(slideIndex + 1)}, please use the Notes pane in PowerPoint to add: "${preview}"`;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const duplicateSlide: Tool = {
  name: 'duplicate_slide',
  description:
    'Duplicate an existing slide by copying its text content into a new slide. Note: Only text shapes are copied; complex graphics may not be preserved.',
  parameters: {
    type: 'object',
    properties: {
      sourceIndex: { type: 'number', description: '0-based index of the slide to duplicate.' },
    },
    required: ['sourceIndex'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { sourceIndex } = (args ?? {}) as { sourceIndex: number };
    try {
      return await PowerPoint.run(async context => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();

        const slideCount = slides.items.length;
        if (slideCount === 0) return 'Presentation has no slides.';
        if (sourceIndex < 0 || sourceIndex >= slideCount) {
          return `Invalid sourceIndex ${String(sourceIndex)}. Must be 0-${String(slideCount - 1)}.`;
        }

        // Collect shapes from source slide
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

        // Add a new blank slide and copy text shapes
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
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

export const powerPointTools: Tool[] = [
  getPresentationOverview,
  getPresentationContent,
  getSlideImage,
  getSlideNotes,
  setPresentationContent,
  addSlideFromCode,
  clearSlide,
  updateSlideShape,
  setSlideNotes,
  duplicateSlide,
];
