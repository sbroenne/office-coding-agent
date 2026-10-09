import PptxGenJS from 'pptxgenjs';

type ShapeKind =
  | 'rect'
  | 'roundRect'
  | 'ellipse'
  | 'triangle'
  | 'diamond'
  | 'hexagon'
  | 'star5'
  | 'chevron'
  | 'arrowRight'
  | 'line';
type Alignment = 'left' | 'center' | 'right';
type VerticalAlignment = 'top' | 'mid' | 'bottom';

interface Position {
  x: number;
  y: number;
  w: number;
  h: number;
}

interface TextElement extends Position {
  type: 'text';
  text: string | string[];
  fontSize?: number;
  fontFace?: string;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  align?: Alignment;
  valign?: VerticalAlignment;
  fillColor?: string;
}

interface ShapeElement extends Position {
  type: 'shape';
  shape: ShapeKind;
  fillColor?: string;
  lineColor?: string;
  lineWidth?: number;
}

interface ImageElement extends Position {
  type: 'image';
  data: string;
  altText?: string;
}

interface TableElement extends Position {
  type: 'table';
  rows: string[][];
  fontSize?: number;
  color?: string;
  borderColor?: string;
}

type ChartType = 'bar' | 'line' | 'pie' | 'doughnut';

interface ChartSeries {
  name: string;
  labels: string[];
  values: number[];
}

interface ChartElement extends Position {
  type: 'chart';
  chartType: ChartType;
  title?: string;
  series: ChartSeries[];
  colors?: string[];
}

type SlideElement = TextElement | ShapeElement | ImageElement | TableElement | ChartElement;

interface SlideSpec {
  backgroundColor?: string;
  elements: SlideElement[];
}

const MAX_SPEC_LENGTH = 8_500_000;
const MAX_ELEMENTS = 100;
const MAX_TEXT_LENGTH = 10_000;
const MAX_TABLE_ROWS = 50;
const MAX_TABLE_COLUMNS = 20;
const MAX_IMAGE_LENGTH = 8_000_000;
const COLOR_RE = /^[\da-f]{6}$/i;
const IMAGE_DATA_RE = /^data:image\/(?:png|jpeg);base64,[\da-z+/]+={0,2}$/i;
const CHART_KINDS = new Set<ChartType>(['bar', 'line', 'pie', 'doughnut']);
const SHAPE_KINDS = new Set<ShapeKind>([
  'rect',
  'roundRect',
  'ellipse',
  'triangle',
  'diamond',
  'hexagon',
  'star5',
  'chevron',
  'arrowRight',
  'line',
]);

function asRecord(value: unknown, label: string): Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new Error(`${label} must be an object.`);
  }
  return value as Record<string, unknown>;
}

function checkKeys(record: Record<string, unknown>, allowedKeys: string[], label: string): void {
  const unknownKey = Object.keys(record).find(key => !allowedKeys.includes(key));
  if (unknownKey) throw new Error(`${label} contains unsupported property '${unknownKey}'.`);
}

function readString(value: unknown, label: string, maxLength = MAX_TEXT_LENGTH): string {
  if (typeof value !== 'string' || value.length > maxLength) {
    throw new Error(`${label} must be a string no longer than ${String(maxLength)} characters.`);
  }
  return value;
}

function readOptionalString(
  record: Record<string, unknown>,
  key: string,
  label: string
): string | undefined {
  const value = record[key];
  if (value === undefined) return undefined;
  return readString(value, label);
}

function readColor(value: unknown, label: string): string {
  if (typeof value !== 'string' || !COLOR_RE.test(value)) {
    throw new Error(`${label} must be a 6-digit hex color without '#'.`);
  }
  return value.toUpperCase();
}

function readOptionalColor(
  record: Record<string, unknown>,
  key: string,
  label: string
): string | undefined {
  const value = record[key];
  if (value === undefined) return undefined;
  return readColor(value, label);
}

function readOptionalNumber(
  record: Record<string, unknown>,
  key: string,
  label: string,
  min: number,
  max: number
): number | undefined {
  const value = record[key];
  if (value === undefined) return undefined;
  if (typeof value !== 'number' || !Number.isFinite(value) || value < min || value > max) {
    throw new Error(`${label} must be a number between ${String(min)} and ${String(max)}.`);
  }
  return value;
}

function readOptionalBoolean(
  record: Record<string, unknown>,
  key: string,
  label: string
): boolean | undefined {
  const value = record[key];
  if (value === undefined) return undefined;
  if (typeof value !== 'boolean') throw new Error(`${label} must be true or false.`);
  return value;
}

function isSupportedImageData(data: string): boolean {
  const match = /^data:image\/(png|jpeg);base64,([\da-z+/]+={0,2})$/i.exec(data);
  if (!match) return false;

  try {
    const bytes = atob(match[2]);
    if (match[1].toLowerCase() === 'png') {
      return (
        bytes.length >= 8 &&
        bytes.charCodeAt(0) === 0x89 &&
        bytes.slice(1, 4) === 'PNG' &&
        bytes.charCodeAt(4) === 0x0d &&
        bytes.charCodeAt(5) === 0x0a &&
        bytes.charCodeAt(6) === 0x1a &&
        bytes.charCodeAt(7) === 0x0a
      );
    }
    return bytes.length >= 3 && bytes.charCodeAt(0) === 0xff && bytes.charCodeAt(1) === 0xd8;
  } catch {
    return false;
  }
}

function readPosition(
  record: Record<string, unknown>,
  label: string,
  width: number,
  height: number,
  allowZeroHeight = false
): Position {
  const values = {} as Position;
  for (const key of ['x', 'y', 'w', 'h'] as const) {
    const value = record[key];
    if (typeof value !== 'number' || !Number.isFinite(value)) {
      throw new Error(`${label}.${key} must be a finite number.`);
    }
    values[key] = value;
  }

  if (
    values.x < 0 ||
    values.y < 0 ||
    values.w <= 0 ||
    values.h < 0 ||
    (!allowZeroHeight && values.h === 0) ||
    values.x + values.w > width ||
    values.y + values.h > height
  ) {
    throw new Error(`${label} must fit within the ${width}" × ${height}" slide.`);
  }
  return values;
}

function readElement(value: unknown, index: number, width: number, height: number): SlideElement {
  const label = `elements[${String(index)}]`;
  const record = asRecord(value, label);
  const type = record.type;

  if (type === 'text') {
    checkKeys(
      record,
      [
        'type',
        'text',
        'x',
        'y',
        'w',
        'h',
        'fontSize',
        'fontFace',
        'color',
        'bold',
        'italic',
        'align',
        'valign',
        'fillColor',
      ],
      label
    );
    const align = record.align;
    if (align !== undefined && align !== 'left' && align !== 'center' && align !== 'right') {
      throw new Error(`${label}.align must be left, center, or right.`);
    }
    const valign = record.valign;
    if (valign !== undefined && valign !== 'top' && valign !== 'mid' && valign !== 'bottom') {
      throw new Error(`${label}.valign must be top, mid, or bottom.`);
    }
    const text = record.text;
    if (Array.isArray(text)) {
      if (
        text.length === 0 ||
        text.length > 100 ||
        text.some(line => typeof line !== 'string' || line.length > MAX_TEXT_LENGTH)
      ) {
        throw new Error(`${label}.text must contain 1-100 short text lines.`);
      }
    } else if (typeof text !== 'string' || text.length > MAX_TEXT_LENGTH) {
      throw new Error(`${label}.text must be a short string or a list of short text lines.`);
    }
    return {
      type,
      text: Array.isArray(text) ? (text as string[]) : (text as string),
      ...readPosition(record, label, width, height),
      fontSize: readOptionalNumber(record, 'fontSize', `${label}.fontSize`, 8, 96),
      fontFace: readOptionalString(record, 'fontFace', `${label}.fontFace`),
      color: readOptionalColor(record, 'color', `${label}.color`),
      bold: readOptionalBoolean(record, 'bold', `${label}.bold`),
      italic: readOptionalBoolean(record, 'italic', `${label}.italic`),
      align,
      valign,
      fillColor: readOptionalColor(record, 'fillColor', `${label}.fillColor`),
    };
  }

  if (type === 'shape') {
    checkKeys(
      record,
      ['type', 'shape', 'x', 'y', 'w', 'h', 'fillColor', 'lineColor', 'lineWidth'],
      label
    );
    const shape = record.shape;
    if (typeof shape !== 'string' || !SHAPE_KINDS.has(shape as ShapeKind)) {
      throw new Error(`${label}.shape is not a supported shape.`);
    }
    return {
      type,
      shape: shape as ShapeKind,
      ...readPosition(record, label, width, height, shape === 'line'),
      fillColor: readOptionalColor(record, 'fillColor', `${label}.fillColor`),
      lineColor: readOptionalColor(record, 'lineColor', `${label}.lineColor`),
      lineWidth: readOptionalNumber(record, 'lineWidth', `${label}.lineWidth`, 0.1, 20),
    };
  }

  if (type === 'image') {
    checkKeys(record, ['type', 'data', 'altText', 'x', 'y', 'w', 'h'], label);
    const data = readString(record.data, `${label}.data`, MAX_IMAGE_LENGTH);
    if (!IMAGE_DATA_RE.test(data) || !isSupportedImageData(data)) {
      throw new Error(`${label}.data must contain a valid base64 PNG or JPEG image.`);
    }
    return {
      type,
      data,
      altText: readOptionalString(record, 'altText', `${label}.altText`),
      ...readPosition(record, label, width, height),
    };
  }

  if (type === 'table') {
    checkKeys(
      record,
      ['type', 'rows', 'x', 'y', 'w', 'h', 'fontSize', 'color', 'borderColor'],
      label
    );
    if (
      !Array.isArray(record.rows) ||
      record.rows.length === 0 ||
      record.rows.length > MAX_TABLE_ROWS
    ) {
      throw new Error(`${label}.rows must contain between 1 and ${String(MAX_TABLE_ROWS)} rows.`);
    }
    const rows = record.rows.map((row, rowIndex) => {
      if (
        !Array.isArray(row) ||
        row.length === 0 ||
        row.length > MAX_TABLE_COLUMNS ||
        row.some(cell => typeof cell !== 'string' || cell.length > MAX_TEXT_LENGTH)
      ) {
        throw new Error(
          `${label}.rows[${String(rowIndex)}] must contain 1-${String(MAX_TABLE_COLUMNS)} short text cells.`
        );
      }
      return row as string[];
    });
    if (new Set(rows.map(row => row.length)).size !== 1) {
      throw new Error(`${label}.rows must all have the same number of cells.`);
    }
    return {
      type,
      rows,
      ...readPosition(record, label, width, height),
      fontSize: readOptionalNumber(record, 'fontSize', `${label}.fontSize`, 8, 48),
      color: readOptionalColor(record, 'color', `${label}.color`),
      borderColor: readOptionalColor(record, 'borderColor', `${label}.borderColor`),
    };
  }

  if (type === 'chart') {
    checkKeys(
      record,
      ['type', 'chartType', 'title', 'series', 'colors', 'x', 'y', 'w', 'h'],
      label
    );
    const chartType = record.chartType;
    if (typeof chartType !== 'string' || !CHART_KINDS.has(chartType as ChartType)) {
      throw new Error(`${label}.chartType must be bar, line, pie, or doughnut.`);
    }
    if (!Array.isArray(record.series) || record.series.length === 0 || record.series.length > 10) {
      throw new Error(`${label}.series must contain between 1 and 10 data series.`);
    }
    const series = record.series.map((item, seriesIndex): ChartSeries => {
      const seriesLabel = `${label}.series[${String(seriesIndex)}]`;
      const row = asRecord(item, seriesLabel);
      checkKeys(row, ['name', 'labels', 'values'], seriesLabel);
      const name = readString(row.name, `${seriesLabel}.name`);
      if (
        !Array.isArray(row.labels) ||
        row.labels.length === 0 ||
        row.labels.length > 100 ||
        row.labels.some(itemLabel => typeof itemLabel !== 'string' || itemLabel.length > 200)
      ) {
        throw new Error(`${seriesLabel}.labels must contain 1-100 short strings.`);
      }
      if (
        !Array.isArray(row.values) ||
        row.values.length !== row.labels.length ||
        row.values.some(number => typeof number !== 'number' || !Number.isFinite(number))
      ) {
        throw new Error(`${seriesLabel}.values must contain a finite number for each label.`);
      }
      return { name, labels: row.labels as string[], values: row.values as number[] };
    });
    if (new Set(series.map(item => item.labels.length)).size !== 1) {
      throw new Error(`${label}.series must all have the same number of labels.`);
    }
    const colors = record.colors;
    if (
      colors !== undefined &&
      (!Array.isArray(colors) ||
        colors.length > 10 ||
        colors.some(color => typeof color !== 'string' || !COLOR_RE.test(color)))
    ) {
      throw new Error(`${label}.colors must contain up to 10 six-digit hex colors.`);
    }
    return {
      type,
      chartType: chartType as ChartType,
      title: readOptionalString(record, 'title', `${label}.title`),
      series,
      colors:
        colors === undefined ? undefined : (colors as string[]).map(color => color.toUpperCase()),
      ...readPosition(record, label, width, height),
    };
  }

  throw new Error(`${label}.type must be text, shape, image, table, or chart.`);
}

function parseSlideSpec(json: string, width: number, height: number): SlideSpec {
  if (json.length > MAX_SPEC_LENGTH) {
    throw new Error(`Slide description exceeds the ${String(MAX_SPEC_LENGTH)} character limit.`);
  }
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) {
    throw new Error('Slide dimensions must be positive finite numbers.');
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(json);
  } catch {
    throw new Error('Slide description must be valid JSON.');
  }

  const record = asRecord(parsed, 'Slide description');
  checkKeys(record, ['backgroundColor', 'elements'], 'Slide description');
  if (!Array.isArray(record.elements) || record.elements.length > MAX_ELEMENTS) {
    throw new Error(
      `Slide description must contain no more than ${String(MAX_ELEMENTS)} elements.`
    );
  }
  return {
    backgroundColor:
      record.backgroundColor === undefined
        ? undefined
        : readColor(record.backgroundColor, 'backgroundColor'),
    elements: record.elements.map((element, index) => readElement(element, index, width, height)),
  };
}

export async function renderSlideSpecToBase64(
  json: string,
  width: number,
  height: number
): Promise<string> {
  const spec = parseSlideSpec(json, width, height);
  const pptx = new PptxGenJS();
  const chartTypes: Record<ChartType, PptxGenJS.CHART_NAME> = {
    bar: pptx.ChartType.bar,
    line: pptx.ChartType.line,
    pie: pptx.ChartType.pie,
    doughnut: pptx.ChartType.doughnut,
  };
  const shapeTypes: Record<ShapeKind, PptxGenJS.ShapeType> = {
    rect: pptx.ShapeType.rect,
    roundRect: pptx.ShapeType.roundRect,
    ellipse: pptx.ShapeType.ellipse,
    triangle: pptx.ShapeType.triangle,
    diamond: pptx.ShapeType.diamond,
    hexagon: pptx.ShapeType.hexagon,
    star5: pptx.ShapeType.star5,
    chevron: pptx.ShapeType.chevron,
    arrowRight: pptx.ShapeType.rightArrow,
    line: pptx.ShapeType.line,
  };
  pptx.defineLayout({ name: 'CUSTOM', width, height });
  pptx.layout = 'CUSTOM';
  const slide = pptx.addSlide();

  if (spec.backgroundColor) slide.background = { color: spec.backgroundColor };

  for (const element of spec.elements) {
    if (element.type === 'text') {
      const text = Array.isArray(element.text)
        ? element.text.map((line, index) => ({
            text: line,
            options: { bullet: true, breakLine: index < element.text.length - 1 },
          }))
        : element.text;
      slide.addText(text, {
        x: element.x,
        y: element.y,
        w: element.w,
        h: element.h,
        fontSize: element.fontSize,
        fontFace: element.fontFace,
        color: element.color,
        bold: element.bold,
        italic: element.italic,
        align: element.align,
        valign: element.valign === 'mid' ? 'middle' : element.valign,
        fill: element.fillColor ? { color: element.fillColor } : undefined,
        fit: 'shrink',
      });
    } else if (element.type === 'shape') {
      slide.addShape(shapeTypes[element.shape], {
        x: element.x,
        y: element.y,
        w: element.w,
        h: element.h,
        fill: element.fillColor ? { color: element.fillColor } : undefined,
        line: element.lineColor
          ? { color: element.lineColor, width: element.lineWidth ?? 1 }
          : undefined,
      });
    } else if (element.type === 'image') {
      slide.addImage({
        data: element.data,
        x: element.x,
        y: element.y,
        w: element.w,
        h: element.h,
        altText: element.altText,
      });
    } else if (element.type === 'table') {
      slide.addTable(
        element.rows.map(row => row.map(text => ({ text }))),
        {
          x: element.x,
          y: element.y,
          w: element.w,
          h: element.h,
          fontSize: element.fontSize,
          color: element.color,
          border: {
            type: 'solid',
            color: element.borderColor ?? 'CFCFCF',
            pt: 0.5,
          },
        }
      );
    } else {
      slide.addChart(chartTypes[element.chartType], element.series, {
        x: element.x,
        y: element.y,
        w: element.w,
        h: element.h,
        showTitle: Boolean(element.title),
        title: element.title,
        showLegend: element.series.length > 1,
        chartColors: element.colors,
      });
    }
  }

  const output = await pptx.write({ outputType: 'base64' });
  if (typeof output !== 'string') throw new Error('Failed to generate slide data.');
  return output;
}
