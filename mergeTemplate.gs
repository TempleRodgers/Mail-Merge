function mergeTemplate(temporaryBody, mergeDocBody, toMergeData) {
//  console.log("---- Entering mergeTemplate ----");
//  console.log("temporaryBody (before):", temporaryBody.getText()); // Log the entire temporaryBody content
//  console.log("toMergeData:", toMergeData);

  for (let placeholder in toMergeData) {
    let value = toMergeData[placeholder] || "";
    // Correct regex - only escape special characters once:
    // let escapedPlaceholder = placeholder.replace(/[\[\]]/g, '\\\\$&');
    let escapedPlaceholder = placeholder.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); 
    // Use the 'i' flag for case-insensitive matching and 'g' for global search
    let regex = new RegExp("{{" + escapedPlaceholder + "}}", "gi"); 

//    console.log("Escaped Placeholder:", escapedPlaceholder);
//    console.log("Value:", value);
//    console.log("Regex:", regex);

    // Find all occurrences of the placeholder in the temporary body
    let matches = temporaryBody.getText().matchAll(regex);

    // Replace each occurrence with the value
    for (const match of matches) {
      temporaryBody.replaceText(match[0], value);
    }

//    temporaryBody.replaceText(regex, value);
//    console.log("temporaryBody (after replaceText):", temporaryBody.getText()); // Log after each replacement
  }

  // Get the number of child elements in the temporary body
  var numChildren = temporaryBody.getNumChildren();

  // Loop through each child element
  for (var i = 0; i < numChildren; i++) {
    // Copy the element from the temporary body
    var element = temporaryBody.getChild(i).copy();

    // Determine the element type and append accordingly
    switch (element.getType()) {
      case DocumentApp.ElementType.PARAGRAPH:
//        console.log("Appending Paragraph:", element.asParagraph().getText());
        mergeDocBody.appendParagraph(element);
        break;
      case DocumentApp.ElementType.LIST_ITEM:
//        console.log("Appending List Item:", element.asListItem().getText());
        mergeDocBody.appendListItem(element);
        break;
      case DocumentApp.ElementType.TABLE:
//        console.log("Appending Table: (table data not easily accessible)");
        const table = element.asTable();
        const newTable = mergeDocBody.appendTable();
        for (let row = 0; row < table.getNumRows(); row++) {
          const newRow = newTable.appendTableRow();
          for (let col = 0; col < table.getRow(row).getNumCells(); col++) {
            newRow.appendTableCell(table.getRow(row).getCell(col).copy());
          }
        }
        break;
      case DocumentApp.ElementType.TABLE_ROW:
        const tableRow = element.asTableRow();
        mergeDocBody.appendTableRow(tableRow.copy());
        break;
      case DocumentApp.ElementType.TABLE_CELL:
        const tableCell = element.asTableCell();
        const newCell = mergeDocBody.appendTableCell();
        deepCopyDocumentElements(tableCell, newCell);
        break;
      case DocumentApp.ElementType.HORIZONTAL_RULE:
//        console.log("Appending Horizontal Rule"); // No specific value to log
        mergeDocBody.appendHorizontalRule();
        break;
      case DocumentApp.ElementType.INLINE_IMAGE:
//        console.log("Appending Inline Image:", element.asInlineImage().getBlob().getDataAsString());
//        destinationBody.appendImage(element.asInlineImage().getBlob());
        const inlineImage = element.asInlineImage();
        const imageBlob = inlineImage.getBlob();
        const copiedImage = mergeDocBody.appendImage(imageBlob);
        // Copy other properties of the inline image
        copiedImage.setWidth(inlineImage.getWidth());
        copiedImage.setHeight(inlineImage.getHeight());
        copiedImage.setAltText(inlineImage.getAltText());
        break;
      case DocumentApp.ElementType.PAGE_BREAK:
//        console.log("Appending Page Break"); // No specific value to log
        mergeDocBody.appendPageBreak();
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
      case DocumentApp.ElementType.HEADER:
        const header = element.asHeader();
        const newHeader = mergeDocBody.appendHeader();
        deepCopyDocumentElements(header, newHeader);
        break;
      case DocumentApp.ElementType.FOOTER:
        const footer = element.asFooter();
        const newFooter = mergeDocBody.appendFooter();
         // Ensure all children, including images, are copied into each section's footer.
        for (let j = 0; j < sourceFooter.getNumChildren(); j++) {
          const footerElement = sourceFooter.getChild(j);
          deepCopyDocumentElements(footerElement, destinationFooter);
        }
        break;
      case DocumentApp.ElementType.BOOKMARK:
        const bookmark = element.asBookmark();
        mergeDocBody.appendBookmark(bookmark.copy());
        break;
      case DocumentApp.ElementType.HORIZONTAL_RULER:
        mergeDocBody.appendHorizontalRule();
        break;
      default:
        // Handle other element types as needed
        console.log("Unsupported element type:", element.getType());
        break;
    }
  }
  console.log(`merge record: ${Object.keys(toMergeData)[0]} - ${Object.values(toMergeData)[0]}`);
}
