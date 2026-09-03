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

function computeSha256(text) {
  if (!text) return "";
  var signature = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  var hexStr = '';
  for (var i = 0; i < signature.length; i++) {
    var byte = signature[i];
    if (byte < 0) byte += 256;
    var byteStr = byte.toString(16);
    if (byteStr.length == 1) byteStr = '0' + byteStr;
    hexStr += byteStr;
  }
  return hexStr;
}

function captureSelectionInGoogle() {
  try {
    if (SpreadsheetApp.getActiveSpreadsheet()) {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getActiveSheet();
      var range = sheet.getActiveRange();
      
      if (!range || (range.getNumRows() === 1 && range.getNumColumns() === 1 && !range.getValue())) {
        var dataRange = sheet.getDataRange();
        if (dataRange.getNumRows() > 1 || dataRange.getNumColumns() > 1 || String(dataRange.getValue()).trim() !== "") {
          range = dataRange;
        }
      }
      
      var values = range.getDisplayValues();
      var formulas = range.getFormulas();
      var columns = [];
      var rows = [];
      
      if (values && values.length > 0) {
        columns = values[0].map(String);
        if (values.length > 1) {
          rows = values.slice(1).map(function(row, rIdx) {
            return row.map(function(val, cIdx) {
              var f = formulas[rIdx + 1] && formulas[rIdx + 1][cIdx];
              if (f && String(f).startsWith("=")) {
                return String(f);
              }
              return val;
            });
          });
        }
      }
      
      var rangeName = "";
      var namedRanges = ss.getNamedRanges();
      for (var n = 0; n < namedRanges.length; n++) {
        var nr = namedRanges[n];
        if (nr.getRange().getA1Notation() === range.getA1Notation() && nr.getRange().getSheet().getName() === sheet.getName()) {
          rangeName = nr.getName();
          break;
        }
      }
      
      var contentString = JSON.stringify(values);
      return {
        anchor: { 
          kind: "grid", 
          a1_range: range.getA1Notation(), 
          sheet_name: sheet.getName(), 
          named_range: rangeName,
          content_sha: computeSha256(contentString) 
        },
        grid: { columns: columns, rows: rows }
      };
    }
  } catch(e) {}

  try {
    if (DocumentApp.getActiveDocument()) {
      var doc = DocumentApp.getActiveDocument();
      var selection = doc.getSelection();
      var elements = [];
      if (selection) {
        var rangeEls = selection.getRangeElements();
        for (var i = 0; i < rangeEls.length; i++) elements.push(rangeEls[i].getElement());
      } else {
        var body = doc.getBody();
        var numChildren = body.getNumChildren();
        for (var i = 0; i < numChildren; i++) elements.push(body.getChild(i));
      }
      
      var blocks = [];
      var fullText = "";
      var currentHeadingPath = [];
      
      for (var i = 0; i < elements.length; i++) {
        var element = elements[i];
        if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
          var p = element.asParagraph();
          var text = p.getText().trim();
          if (!text) continue;
          
          var heading = p.getHeading();
          var level = 0;
          if (heading === DocumentApp.ParagraphHeading.HEADING1) level = 1;
          else if (heading === DocumentApp.ParagraphHeading.HEADING2) level = 2;
          else if (heading === DocumentApp.ParagraphHeading.HEADING3) level = 3;
          else if (heading === DocumentApp.ParagraphHeading.HEADING4) level = 4;
          else if (heading === DocumentApp.ParagraphHeading.HEADING5) level = 5;
          else if (heading === DocumentApp.ParagraphHeading.HEADING6) level = 6;
          
          if (level > 0) {
            currentHeadingPath = currentHeadingPath.slice(0, level - 1);
            currentHeadingPath[level - 1] = text;
          }
          
          blocks.push({
            type: (level > 0) ? "heading" : "paragraph",
            text: text,
            heading_path: currentHeadingPath.slice()
          });
          fullText += text + "\n";
        } else if (element.editAsText) {
           var text = element.asText().getText().trim();
           if (text) {
             blocks.push({ type: "paragraph", text: text, heading_path: currentHeadingPath.slice() });
             fullText += text + "\n";
           }
        }
      }
      
      return {
        anchor: { kind: "text", content_sha: computeSha256(fullText) },
        text: { blocks: blocks, char_count: fullText.length }
      };
    }
  } catch(e) {}

  try {
    if (SlidesApp.getActivePresentation()) {
      var selection = SlidesApp.getActivePresentation().getSelection();
      var slide = null;
      if (selection) {
         slide = selection.getCurrentPage();
      }
      var slides = SlidesApp.getActivePresentation().getSlides();
      if (!slide && slides.length > 0) {
         slide = slides[0];
      }
      
      var slideIndex = -1;
      for (var s = 0; s < slides.length; s++) {
        if (slides[s].getObjectId() === slide.getObjectId()) {
          slideIndex = s;
          break;
        }
      }
      
      var fullText = "";
      var blocks = [];
      if (slide) {
        var elements = slide.getPageElements();
        for (var i = 0; i < elements.length; i++) {
          var el = elements[i];
          if (el.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
            var shape = el.asShape();
            if (shape.getText) {
              var text = shape.getText().asString().trim();
              if (text) {
                blocks.push({ type: "paragraph", text: text });
                fullText += text + "\n";
              }
            }
          } else if (el.getPageElementType() === SlidesApp.PageElementType.TABLE) {
            var table = el.asTable();
            var rows = table.getNumRows();
            var cols = table.getNumColumns();
            var tableText = "";
            for (var r = 0; r < rows; r++) {
               var rowText = [];
               for (var c = 0; c < cols; c++) {
                 rowText.push(table.getCell(r, c).getText().asString().trim().replace(/\n/g, ' '));
               }
               tableText += "| " + rowText.join(" | ") + " |\n";
            }
            if (tableText) {
               blocks.push({ type: "table", text: tableText });
               fullText += tableText + "\n";
            }
          }
        }
      }
      
      return {
        anchor: { 
          kind: "shape", 
          slide_id: slide ? slide.getObjectId() : null,
          slide_index: slideIndex > -1 ? slideIndex : null,
          content_sha: computeSha256(fullText) 
        },
        text: { blocks: blocks, char_count: fullText.length }
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
      var range = sheet.getRange(anchor.a1_range);
      range.activate();
      
      var origBgs = range.getBackgrounds();
      range.setBackground("#FEF08A");
      SpreadsheetApp.flush();
      
      Utilities.sleep(1500);
      range.setBackgrounds(origBgs);
    }
  } catch(e) {}

  try {
    if (SlidesApp.getActivePresentation() && anchor.slide_index != null) {
      var pres = SlidesApp.getActivePresentation();
      var slides = pres.getSlides();
      if (anchor.slide_index >= 0 && anchor.slide_index < slides.length) {
        slides[anchor.slide_index].selectAsCurrentPage();
      }
    }
  } catch(e) {}

  try {
    if (DocumentApp.getActiveDocument()) {
      var doc = DocumentApp.getActiveDocument();
      var body = doc.getBody();
      
      if (anchor.start_offset != null && anchor.end_offset != null) {
        var start = anchor.start_offset;
        var end = anchor.end_offset;
        var currentOffset = 0;
        var startElement = null;
        var startOffsetInElement = 0;
        var endElement = null;
        var endOffsetInElement = 0;
        
        var numChildren = body.getNumChildren();
        for (var i = 0; i < numChildren; i++) {
          var child = body.getChild(i);
          if (child.editAsText) {
            var text = child.asText().getText() + "\n";
            var textLen = text.length;
            
            if (!startElement && start >= currentOffset && start < currentOffset + textLen) {
              startElement = child.asText();
              startOffsetInElement = start - currentOffset;
            }
            if (!endElement && end >= currentOffset && end <= currentOffset + textLen) {
              endElement = child.asText();
              endOffsetInElement = end - currentOffset;
              if (endOffsetInElement > child.asText().getText().length) {
                endOffsetInElement = child.asText().getText().length;
              }
            }
            
            currentOffset += textLen;
            if (startElement && endElement) break;
          }
        }
        
        if (startElement && endElement) {
          var rangeBuilder = doc.newRange();
          rangeBuilder.addElement(startElement, startOffsetInElement, endOffsetInElement); // Simplified for same-element selection
          if (startElement === endElement) {
              rangeBuilder = doc.newRange();
              rangeBuilder.addElement(startElement, startOffsetInElement, endOffsetInElement);
          } else {
              rangeBuilder = doc.newRange();
              rangeBuilder.addElementsBetween(startElement, endElement);
          }
          var builtRange = rangeBuilder.build();
          doc.setSelection(builtRange);
          
          var els = builtRange.getRangeElements();
          var origBg = [];
          for (var i = 0; i < els.length; i++) {
             var el = els[i].getElement();
             if (el.editAsText) {
                var bg = el.asText().getBackgroundColor(els[i].getStartOffset() || 0) || null;
                origBg.push({el: el.asText(), bg: bg, start: els[i].getStartOffset(), end: els[i].getEndOffsetInclusive()});
                if (els[i].isPartial()) {
                  el.asText().setBackgroundColor(els[i].getStartOffset(), els[i].getEndOffsetInclusive(), "#FEF08A");
                } else {
                  el.asText().setBackgroundColor("#FEF08A");
                }
             }
          }
          
          Utilities.sleep(1500);
          
          for (var j = 0; j < origBg.length; j++) {
             var item = origBg[j];
             if (item.start != -1 && item.end != -1) {
                item.el.setBackgroundColor(item.start, item.end, item.bg);
             } else {
                item.el.setBackgroundColor(item.bg);
             }
          }
          return;
        }
      }
      
      // Fallback to findText
      var needle = (citation.label || "").trim().substring(0, 200);
      var found = body.findText(needle);
      if (found) {
        doc.setSelection(found);
      }
    }
  } catch(e) {}

  try {
    if (SlidesApp.getActivePresentation()) {
      if (anchor.slide_id) {
         var slide = SlidesApp.getActivePresentation().getSlideById(anchor.slide_id);
         if (slide) slide.selectAsCurrentPage();
      } else if (anchor.slide_index != null) {
         var slides = SlidesApp.getActivePresentation().getSlides();
         if (slides.length > anchor.slide_index) {
           slides[anchor.slide_index].selectAsCurrentPage();
         }
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

function getDocumentIdInGoogle() {
  try { if (SpreadsheetApp.getActiveSpreadsheet()) return SpreadsheetApp.getActiveSpreadsheet().getId(); } catch(e) {}
  try { if (DocumentApp.getActiveDocument()) return DocumentApp.getActiveDocument().getId(); } catch(e) {}
  try { if (SlidesApp.getActivePresentation()) return SlidesApp.getActivePresentation().getId(); } catch(e) {}
  return "google-0";
}

function getDocumentTitleInGoogle() {
  try { if (SpreadsheetApp.getActiveSpreadsheet()) return SpreadsheetApp.getActiveSpreadsheet().getName(); } catch(e) {}
  try { if (DocumentApp.getActiveDocument()) return DocumentApp.getActiveDocument().getName(); } catch(e) {}
  try { if (SlidesApp.getActivePresentation()) return SlidesApp.getActivePresentation().getName(); } catch(e) {}
  return "the open document";
}

function onHomepage(e) { return createSidebarCard(); }
function onDocsHomepage(e) { return createSidebarCard(); }
function onSheetsHomepage(e) { return createSidebarCard(); }
function onSlidesHomepage(e) { return createSidebarCard(); }
