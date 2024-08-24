function mergeTemplate(templateBody, mergeDocBody, toMergeData) {
//  console.log("---- Entering mergeTemplate ----");
//  console.log("temporaryBody (before):", temporaryBody.getText()); // Log the entire temporaryBody content
//  console.log("toMergeData:", toMergeData);

    let temporaryBody = templateBody.copy();
  console.log(`merge record: ${Object.keys(toMergeData)[0]} - ${Object.values(toMergeData)[0]}`);
}

function replacePlaceHolders(tempBody,toMergeData) {
  for (let placeholder in toMergeData) {
    let value = toMergeData[placeholder] || "";
    // Correct regex - only escape special characters once:
    // let escapedPlaceholder = placeholder.replace(/[\[\]]/g, '\\\\$&');
    let escapedPlaceholder = placeholder.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); 
            // Use the 'i' flag for case-insensitive matching and 'g' for global search
            //let regex = new RegExp("{{" + escapedPlaceholder + "}}", "gi");
    let regex = "{{"+escapedPlaceholder+"}}";
    tempBody.replaceText(regex,value);
  }
}
/**
 * Element types are here:
 * https://developers.google.com/apps-script/reference/document/element-type
 * 
*/
function deepCopyDocumentElements(sourceBody, mergeDocBody) {
  for (let i = 0; i < sourceBody.getNumChildren(); i++) {
    // Copy the element from the temporary body
    var element = sourceBody.getChild(i).copy();

    // Determine the element type and append accordingly
    switch (element.getType()) {
      case DocumentApp.ElementType.PARAGRAPH:
        console.log("Appending Paragraph:", element.asParagraph().getText());
        // Insert the new paragraph before the last one
        mergeDocBody.appendParagraph(element);
        break;
      case DocumentApp.ElementType.LIST_ITEM:
        console.log("Appending List Item:", element.asListItem().getText());
        mergeDocBody.appendListItem(element);
        break;
      case DocumentApp.ElementType.TABLE:
        console.log("Appending Table: (table data not easily accessible)");
        const table = element.asTable();
        mergeDocBody.appendTable(table);
        break;
/**      case DocumentApp.ElementType.TABLE:
        console.log("Appending Table: (table data not easily accessible)");
        const table = element.asTable();
        const newTable = mergeDocBody.appendTable();
        for (let row = 0; row < table.getNumRows(); row++) {
          const newRow = newTable.appendTableRow();
          for (let col = 0; col < table.getRow(row).getNumCells(); col++) {
            newRow.appendTableCell(table.getRow(row).getCell(col).copy());
          }
        }
        break;*/
      case DocumentApp.ElementType.HORIZONTAL_RULE:
        console.log("Appending HORIZONTAL_RULE");
        mergeDocBody.appendHorizontalRule();
        break;
      case DocumentApp.ElementType.INLINE_IMAGE:
        console.log("Appending INLINE_IMAGE:", element.asInlineImage().getBlob().getDataAsString());
        destinationBody.appendImage(element.asInlineImage().getBlob());
        const inlineImage = element.asInlineImage();
        const imageBlob = inlineImage.getBlob();
        if (imageBlob && imageBlob.getBytes().length > 0) {
          try {
            const copiedImage = mergeDocBody.appendImage(imageBlob);
            copiedImage.setWidth(inlineImage.getWidth());
            copiedImage.setHeight(inlineImage.getHeight());
            copiedImage.setAltText(inlineImage.getAltText());
            console.log('Image appended successfully:', copiedImage.getAltText());
          } catch (error) {
            console.error('Error appending image:', error);
            // Add more detailed error handling here (e.g., check error message)
          }
        } else {
          console.error('Invalid image blob:', imageBlob);
        }
        break;
      case DocumentApp.ElementType.PAGE_BREAK:
        console.log("Appending PAGE_BREAK"); // No specific value to log
        mergeDocBody.appendPageBreak(element);
        break;
      case DocumentApp.ElementType.HORIZONTAL_RULE:
        console.log("Appending HORIZONTAL_RULE");
        mergeDocBody.appendHorizontalRule();
        break;
/**
 * VALID METHODS ARE HERE:
 * https://developers.google.com/apps-script/reference/document/body#appendtabletable
 * VALID METHODS
 * Method	                  Return type	    Brief description
    appendHorizontalRule()	HorizontalRule	Creates and appends a new HorizontalRule.
    appendImage(image)	    InlineImage	    Creates and appends a new InlineImage from the specified image blob.
    appendImage(image)	    InlineImage	    Appends the given InlineImage.
    appendListItem(listItem)	ListItem	    Appends the given ListItem.
    appendListItem(text)	ListItem	        Creates and appends a new ListItem containing the specified text contents.
    appendPageBreak()	PageBreak	            Creates and appends a new PageBreak.
    appendPageBreak(pageBreak)	PageBreak	  Appends the given PageBreak.
    appendParagraph(paragraph)	Paragraph	  Appends the given Paragraph.
    appendParagraph(text)	Paragraph	        Creates and appends a new Paragraph containing the specified text contents.
    appendTable()	          Table	          Creates and appends a new Table.
    appendTable(cells)	    Table	          Appends a new Table containing a TableCell for each specified string value.
    appendTable(table)	    Table	          Appends the given Table.

 *
 * INVALID METHODS
 * e.g. 
 *        case DocumentApp.ElementType.TABLE_CELL:
        console.log("Appending TABLE_CELL");
        const tableCell = element.asTableCell();
        const newCell = mergeDocBody.appendTableCell();
        deepCopyDocumentElements(tableCell, newCell);
        break;
      case DocumentApp.ElementType.TABLE_ROW:
        const tableRow = element.asTableRow();
        mergeDocBody.appendTableRow(tableRow.copy());
        break;
      case DocumentApp.ElementType.FOOTNOTE:
        const footnote = element.asFootnote();
        mergeDocBody.appendFootnote(footnote.copy());
        break;
      case DocumentApp.ElementType.TEXT:
        const text = element.asText();
        mergeDocBody.appendText(text.copy());
        break;
      case DocumentApp.ElementType.INLINE_DRAWING:
        const drawing = element.asInlineDrawing();
        mergeDocBody.appendInlineDrawing(drawing.copy());
        break;
      case DocumentApp.ElementType.EQUATION:
        const equation = element.asEquation();
        const newEquation = mergeDocBody.appendEquation();
        deepCopyDocumentElements(equation, newEquation);
        break;
      case DocumentApp.ElementType.EQUATION_FUNCTION:
        const eqFunction = element.asEquationFunction();
        const newEqFunction = mergeDocBody.appendEquationFunction(eqFunction.copy());
        break;
      case DocumentApp.ElementType.EQUATION_SYMBOL:
        const eqSymbol = element.asEquationSymbol();
        const newEqSymbol = mergeDocBody.appendEquationSymbol(eqSymbol.copy());
        break;
      case DocumentApp.ElementType.EQUATION_ARGUMENT:
        const eqArgument = element.asEquationArgument();
        const newEqArgument = mergeDocBody.appendEquationArgument(eqArgument.copy());
        break;
*/
      default:
        // Handle other element types as needed
        console.log("Unsupported element type:", element.getType());
        break;
    }
  }
}
