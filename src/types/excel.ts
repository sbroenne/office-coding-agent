/** Result of an Excel tool call (returned by tool execute functions) */
export interface ToolCallResult {
  /** Whether the tool execution succeeded */
  success: boolean;
  /** Result data */
  data?: unknown;
  /** Error message if failed */
  error?: string;
}

/** Result of an Excel operation */
export interface ExcelOperationResult {
  /** Whether the operation succeeded */
  success: boolean;
  /** Result data */
  data?: unknown;
  /** Error message if failed */
  errorMessage?: string;
}

/** Represents cell values from a range read */
export interface RangeData {
  /** Address of the range (e.g., "Sheet1!A1:C3") */
  address: string;
  /** 2D array of cell values */
  values: unknown[][];
  /** Number of rows */
  rowCount: number;
  /** Number of columns */
  columnCount: number;
}

/** Represents an Excel table */
export interface TableInfo {
  /** Table name */
  name: string;
  /** Worksheet the table is on */
  sheetName: string;
  /** Range address of the table */
  address: string;
  /** Column headers */
  headers: string[];
  /** Number of data rows (excluding header) */
  rowCount: number;
}

/** Represents a worksheet */
export interface SheetInfo {
  /** Sheet name */
  name: string;
  /** Sheet position (0-based) */
  position: number;
  /** Whether this sheet is currently active */
  isActive: boolean;
  /** Visibility state */
  visibility: 'Visible' | 'Hidden' | 'VeryHidden';
}

/** Represents a chart */
export interface ChartInfo {
  /** Chart name */
  name: string;
  /** Chart type */
  chartType: string;
  /** Worksheet the chart is on */
  sheetName: string;
  /** Chart title */
  title?: string;
}

/** Represents a PivotTable */
export interface PivotTableInfo {
  /** PivotTable name */
  name: string;
  /** Worksheet the PivotTable is on */
  sheetName: string;
}
