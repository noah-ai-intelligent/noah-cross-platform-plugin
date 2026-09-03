import type { XlsxTable, Citation, EditOperationOut, EditOperationReport } from "../../addonClient";
import type { Selection } from "../../document/selection";

export interface DocumentHost {
  insertTable(table: XlsxTable, isUpdate?: boolean): Promise<{ address: string } | void>;
  insertChart(chartType: string, rangeAddress: string): Promise<void>;
  insertProse(text: string): Promise<void>;
  insertImageBase64(base64: string): Promise<void>;
  handleCitation(citation: Citation): Promise<void>;
  captureSelection(): Promise<Selection | null>;
  getDocumentId(): Promise<string>;
  getDocumentTitle(): Promise<string>;
  applyEditOperation(op: EditOperationOut, index: number): Promise<EditOperationReport>;
  insertReport(answer: any): Promise<void>;
}
