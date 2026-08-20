/** PowerPoint insertion helpers — the weakest host: no rich in-place text-run
 * or table-object insertion, and no native chart-object creation. Both text
 * and tabular data go through `setSelectedDataAsync`, tables/charts as a
 * rendered image. This is a real, disclosed capability gap versus Word/Excel
 * (see the implementation plan), not a bug. */

export async function insertText(text: string): Promise<void> {
  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(
      text,
      { coercionType: Office.CoercionType.Text },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(result.error);
        } else {
          resolve();
        }
      }
    );
  });
}

/** Drops a table or chart in as a static image — PowerPoint's JS API has no
 * native table/chart insertion, so a client- or server-rendered PNG is the
 * practical fallback for tabular/chart data on a slide. */
export async function insertImageBase64(base64: string): Promise<void> {
  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(
      base64,
      { coercionType: Office.CoercionType.Image },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(result.error);
        } else {
          resolve();
        }
      }
    );
  });
}
