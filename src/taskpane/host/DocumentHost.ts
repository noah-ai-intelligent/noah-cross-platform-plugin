import type { XlsxTable, Citation } from "../../addonClient";
import type { Selection } from "../../document/selection";

export interface DocumentHost {
  insertTable(table: XlsxTable, isUpdate?: boolean): Promise<{ address: string } | void>;
  insertChart(chartType: string, rangeAddress: string): Promise<void>;
  insertProse(text: string): Promise<void>;
  insertImageBase64(base64: string): Promise<void>;
  handleCitation(citation: Citation): Promise<void>;
  captureSelection(): Promise<Selection | null>;
}
