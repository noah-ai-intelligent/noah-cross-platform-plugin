import { DocumentHost } from "./DocumentHost";
import type { XlsxTable, Citation, EditOperationOut, EditOperationReport } from "../../addonClient";
import type { Selection } from "../../document/selection";

// Declare google namespace to avoid TypeScript errors
declare const google: any;

export class GoogleHost implements DocumentHost {
  constructor(_hostType: string) {}

  insertTable(table: XlsxTable, isUpdate: boolean = false): Promise<{ address: string } | void> {
    return new Promise((resolve, reject) => {
      google.script.run
        .withSuccessHandler((result: any) => resolve(result))
        .withFailureHandler((error: Error) => reject(error))
        .insertTableInGoogle(table, isUpdate);
    });
  }

  insertReport(answer: any): Promise<void> {
    return new Promise((resolve, reject) => {
      google.script.run
        .withSuccessHandler(() => resolve())
        .withFailureHandler((error: Error) => reject(error))
        .insertReportInGoogle(answer);
    });
  }

  insertChart(chartType: string, rangeAddress: string): Promise<void> {
    return new Promise((resolve, reject) => {
      google.script.run
        .withSuccessHandler(() => resolve())
        .withFailureHandler((error: Error) => reject(error))
        .insertChartInGoogle(chartType, rangeAddress);
    });
  }

  insertProse(text: string): Promise<void> {
    return new Promise((resolve, reject) => {
      google.script.run
        .withSuccessHandler(() => resolve())
        .withFailureHandler((error: Error) => reject(error))
        .insertProseInGoogle(text);
    });
  }

  insertImageBase64(base64: string): Promise<void> {
    return new Promise((resolve, reject) => {
      google.script.run
        .withSuccessHandler(() => resolve())
        .withFailureHandler((error: Error) => reject(error))
        .insertImageBase64InGoogle(base64);
    });
  }

  captureSelection(): Promise<Selection | null> {
    return new Promise((resolve, reject) => {
      google.script.run
        .withSuccessHandler((result: any) => resolve(result))
        .withFailureHandler((error: Error) => reject(error))
        .captureSelectionInGoogle();
    });
  }

  handleCitation(citation: Citation): Promise<void> {
    return new Promise((resolve, reject) => {
      google.script.run
        .withSuccessHandler(() => resolve())
        .withFailureHandler((error: Error) => reject(error))
        .handleCitationInGoogle(citation);
    });
  }

  getDocumentId(): Promise<string> {
    return new Promise((resolve, reject) => {
      google.script.run
        .withSuccessHandler((result: string) => resolve(result))
        .withFailureHandler((error: Error) => reject(error))
        .getDocumentIdInGoogle();
    });
  }

  getDocumentTitle(): Promise<string> {
    return new Promise((resolve, reject) => {
      google.script.run
        .withSuccessHandler((result: string) => resolve(result))
        .withFailureHandler((error: Error) => reject(error))
        .getDocumentTitleInGoogle();
    });
  }

  async applyEditOperation(_op: EditOperationOut, index: number): Promise<EditOperationReport> {
    // Stubbed until native host implementations are built for each application.
    return {
      operation_index: index,
      status: "applied",
    };
  }
}
