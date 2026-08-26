function getActiveUi() {
  try { return SpreadsheetApp.getUi(); } catch(e) {}
  try { return DocumentApp.getUi(); } catch(e) {}
  try { return SlidesApp.getUi(); } catch(e) {}
  throw new Error("Cannot determine active Google Workspace App. Please ensure you are testing this Add-on from inside a Document, Spreadsheet, or Presentation, rather than clicking 'Run' directly in the Apps Script editor.");
}

function getHostType() {
  try { if (SpreadsheetApp.getActiveSpreadsheet()) return 'GoogleSheets'; } catch(e) {}
  try { if (DocumentApp.getActiveDocument()) return 'GoogleDocs'; } catch(e) {}
  try { if (SlidesApp.getActivePresentation()) return 'GoogleSlides'; } catch(e) {}
  return 'Unknown';
}

function onOpen(e) {
  var ui = getActiveUi();
  ui.createMenu('NoahAI')
    .addItem('Open Sidebar', 'showSidebar')
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function showSidebar() {
  var html = HtmlService.createTemplateFromFile('index');
  html.hostType = getHostType();
  getActiveUi().showSidebar(html.evaluate().setTitle('NoahAI').setWidth(350));
}

function insertTableInGoogle(table, isUpdate) {
  var numCols = table.columns.length;
  var tableData = [table.columns];
  
  // Fix row lengths
  if (table.rows && table.rows.length > 0) {
    for (var i = 0; i < table.rows.length; i++) {
      var row = table.rows[i];
      if (row.length < numCols) {
        while (row.length < numCols) row.push("");
      } else if (row.length > numCols) {
        row = row.slice(0, numCols);
      }
      tableData.push(row);
    }
    table.rows = tableData.slice(1);
  } else {
    table.rows = [];
  }

  try {
    if (SpreadsheetApp.getActiveSpreadsheet()) {
      var sheet = SpreadsheetApp.getActiveSheet();
      var startRow = sheet.getActiveCell().getRow();
      var startCol = sheet.getActiveCell().getColumn();
      
      var maxRows = sheet.getMaxRows();
      var maxCols = sheet.getMaxColumns();
      var neededRows = startRow + table.rows.length;
      var neededCols = startCol + numCols - 1;
      
      if (neededRows > maxRows) {
        sheet.insertRowsAfter(maxRows, neededRows - maxRows);
      }
      if (neededCols > maxCols) {
        sheet.insertColumnsAfter(maxCols, neededCols - maxCols);
      }

      var headerRange = sheet.getRange(startRow, startCol, 1, numCols);
      headerRange.setValues([table.columns]);
      
      // Apply premium styling to the header
      headerRange.setBackground("#0d5c63");
      headerRange.setFontColor("#ffffff");
      headerRange.setFontWeight("bold");
      headerRange.setHorizontalAlignment("center");
      
      var totalRows = 1;
      if (table.rows.length > 0) {
        var dataRange = sheet.getRange(startRow + 1, startCol, table.rows.length, numCols);
        dataRange.setValues(table.rows);
        totalRows += table.rows.length;
      }
      
      // Apply borders to the entire table and auto-resize columns
      var fullRange = sheet.getRange(startRow, startCol, totalRows, numCols);
      fullRange.setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
      sheet.autoResizeColumns(startCol, numCols);
      
      return { address: fullRange.getA1Notation() };
    }
  } catch (e) {
    throw new Error("Sheets error: " + e.message);
  }

  try {
    if (DocumentApp.getActiveDocument()) {
      var doc = DocumentApp.getActiveDocument();
      var cursor = doc.getCursor();
      
      var uiTable;
      if (cursor) {
        uiTable = cursor.insertTable(tableData);
      } else {
        uiTable = doc.getBody().appendTable(tableData);
      }
      
      // Basic styling
      var headerRow = uiTable.getRow(0);
      for (var c = 0; c < numCols; c++) {
        var cell = headerRow.getCell(c);
        cell.setBackgroundColor("#0d5c63");
        cell.editAsText().setForegroundColor("#ffffff");
      }
      
      return { address: "" };
    }
  } catch (e) {
    throw new Error("Docs error: " + e.message);
  }

  try {
    if (SlidesApp.getActivePresentation()) {
      var selection = SlidesApp.getActivePresentation().getSelection();
      var slide = selection.getCurrentPage();
      if (!slide) throw new Error("No slide selected");
      
      var uiTable = slide.insertTable(tableData.length, numCols);
      for (var r = 0; r < tableData.length; r++) {
        for (var c = 0; c < numCols; c++) {
          uiTable.getCell(r, c).getText().setText(String(tableData[r][c]));
          if (r === 0) {
             uiTable.getCell(r, c).getFill().setSolidFill("#0d5c63");
             uiTable.getCell(r, c).getText().getTextStyle().setForegroundColor("#ffffff");
          }
        }
      }
      return { address: "" };
    }
  } catch (e) {
    throw new Error("Slides error: " + e.message);
  }

  throw new Error("Cannot determine active document to insert table.");
}

function insertChartInGoogle(chartType, rangeAddress) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange(rangeAddress);
  var chart = sheet.newChart()
    .asColumnChart()
    .addRange(range)
    .setPosition(range.getRow(), range.getColumn() + range.getNumColumns() + 1, 0, 0)
    .build();
  sheet.insertChart(chart);
}

function insertProseInGoogle(text) {
  try {
    if (DocumentApp.getActiveDocument()) {
      var doc = DocumentApp.getActiveDocument();
      var cursor = doc.getCursor();
      if (cursor) {
        cursor.insertText(text);
      } else {
        doc.getBody().appendParagraph(text);
      }
    }
  } catch(e) {}
  
  try {
    if (SlidesApp.getActivePresentation()) {
      var selection = SlidesApp.getActivePresentation().getSelection();
      if (selection.getSelectionType() == SlidesApp.SelectionType.TEXT) {
        selection.getTextRange().insertText(0, text);
      }
    }
  } catch(e) {}
}

function insertImageBase64InGoogle(base64) {
  var blob = Utilities.newBlob(Utilities.base64Decode(base64), 'image/png', 'image.png');
  
  try {
    if (DocumentApp.getActiveDocument()) {
      var doc = DocumentApp.getActiveDocument();
      var cursor = doc.getCursor();
      if (cursor) {
        cursor.insertInlineImage(blob);
      } else {
        doc.getBody().appendImage(blob);
      }
      return;
    }
  } catch(e) {}
  
  try {
    if (SlidesApp.getActivePresentation()) {
      var slide = SlidesApp.getActivePresentation().getSelection().getCurrentPage();
      if (slide) {
        slide.insertImage(blob);
      }
    }
  } catch(e) {}
}

function captureSelectionInGoogle() {
  try {
    if (SpreadsheetApp.getActiveSpreadsheet()) {
      var range = SpreadsheetApp.getActiveSheet().getActiveRange();
      var values = range.getValues();
      return {
        anchor: { kind: "grid", a1_range: range.getA1Notation(), sheet_name: range.getSheet().getName() },
        grid: { columns: [], rows: values }
      };
    }
  } catch(e) {}

  try {
    if (DocumentApp.getActiveDocument()) {
      var selection = DocumentApp.getActiveDocument().getSelection();
      var text = "";
      if (selection) {
        var elements = selection.getRangeElements();
        for (var i = 0; i < elements.length; i++) {
          if (elements[i].getElement().editAsText) {
            text += elements[i].getElement().asText().getText() + "\n";
          }
        }
      }
      return {
        anchor: { kind: "text", content_sha: text },
        text: { blocks: [{ type: "paragraph", text: text }], char_count: text.length }
      };
    }
  } catch(e) {}

  return null;
}

function handleCitationInGoogle(citation) {
  var anchor = citation.anchor;
  try {
    if (SpreadsheetApp.getActiveSpreadsheet() && anchor.a1_range) {
      var sheet = anchor.sheet_name ? SpreadsheetApp.getActiveSpreadsheet().getSheetByName(anchor.sheet_name) : SpreadsheetApp.getActiveSheet();
      sheet.getRange(anchor.a1_range).activate();
    }
  } catch(e) {}

  try {
    if (DocumentApp.getActiveDocument()) {
      var needle = (citation.label || "").trim().substring(0, 200);
      var body = DocumentApp.getActiveDocument().getBody();
      var found = body.findText(needle);
      if (found) {
        DocumentApp.getActiveDocument().setSelection(found);
      }
    }
  } catch(e) {}
}

function createSidebarCard() {
  // Automatically open the custom HTML sidebar when the add-on icon is clicked.
  try {
    showSidebar();
  } catch (e) {
    console.error("Failed to auto-open sidebar:", e);
  }

  var card = CardService.newCardBuilder();
  card.setHeader(
    CardService.newCardHeader()
      .setTitle("NoahAI")
      .setImageUrl("https://noah.enpointe.io/brand/logo-green.png")
  );
  var section = CardService.newCardSection();
  
  section.addWidget(
    CardService.newTextParagraph().setText(
      "NoahAI is ready!\n\nIf the interactive sidebar didn't open automatically, or if you closed it previously, click the button below to launch it."
    )
  );
  
  var action = CardService.newAction().setFunctionName('showSidebar');
  section.addWidget(
    CardService.newTextButton()
      .setText('Launch NoahAI')
      .setOnClickAction(action)
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor("#0d5c63")
  );
  
  card.addSection(section);
  return card.build();
}

function onHomepage(e) { return createSidebarCard(); }
function onDocsHomepage(e) { return createSidebarCard(); }
function onSheetsHomepage(e) { return createSidebarCard(); }
function onSlidesHomepage(e) { return createSidebarCard(); }
