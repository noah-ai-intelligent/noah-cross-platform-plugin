import { DocumentHost } from "./DocumentHost";
import type { XlsxTable, Citation } from "../../addonClient";
import type { Selection } from "../../document/selection";
import { captureSelection } from "../../document/selection";
import type { Host } from "../../document/capabilities";
import * as excel from "../officeApi/excel";
import * as word from "../officeApi/word";
import * as powerpoint from "../officeApi/powerpoint";

export class OfficeHost implements DocumentHost {
  constructor(private hostType: Host) {}

  async insertTable(table: XlsxTable, isUpdate: boolean = false): Promise<{ address: string } | void> {
    if (this.hostType === "Excel") {
      return await excel.insertTable(table, isUpdate);
    } else if (this.hostType === "Word") {
      return await word.insertTable(table);
    } else if (this.hostType === "PowerPoint") {
      // Assuming powerpoint can insert table images instead if native tables aren't supported
      // Or we just ignore
    }
  }

  async insertChart(chartType: string, rangeAddress: string): Promise<void> {
    if (this.hostType === "Excel") {
      await excel.insertChart(chartType, rangeAddress);
    }
  }

  async insertProse(text: string): Promise<void> {
    if (this.hostType === "Word") {
      await word.insertProse(text);
    } else if (this.hostType === "PowerPoint") {
      await powerpoint.insertText(text);
    }
  }

  async insertImageBase64(base64: string): Promise<void> {
    if (this.hostType === "Word") {
      await word.insertChartImage(base64);
    } else if (this.hostType === "PowerPoint") {
      await powerpoint.insertImageBase64(base64);
    }
  }

  async captureSelection(): Promise<Selection | null> {
    return captureSelection(this.hostType);
  }

  async handleCitation(citation: Citation): Promise<void> {
    const { anchor } = citation;
    try {
      if (this.hostType === "Excel" && anchor.a1_range) {
        await Excel.run(async (ctx) => {
          const sheet = anchor.sheet_name
            ? ctx.workbook.worksheets.getItem(anchor.sheet_name)
            : ctx.workbook.worksheets.getActiveWorksheet();
          const range = sheet.getRange(anchor.a1_range!);
          range.select();
          await ctx.sync();
        });
      } else if (this.hostType === "PowerPoint" && anchor.slide_index != null) {
        await PowerPoint.run(async (ctx) => {
          const slides = ctx.presentation.slides;
          slides.load("items/id");
          await ctx.sync();
          const target = slides.items[anchor.slide_index!];
          if (target) {
            ctx.presentation.setSelectedSlides([target.id]);
            await ctx.sync();
          }
        });
      } else if (this.hostType === "Word") {
        const needle = (citation.label || "").trim().slice(0, 200);
        if (!needle) return;
        await Word.run(async (ctx) => {
          const hits = ctx.document.body.search(needle, { matchCase: false });
          hits.load("items");
          await ctx.sync();
          if (hits.items.length) {
            hits.items[0].select();
            await ctx.sync();
          }
        });
      }
    } catch {
      console.warn("Couldn't jump to that citation — it may have moved.");
    }
  }
}
