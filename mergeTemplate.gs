function mergeTemplate(templateBody, mergeDocBody, toMergeData) {
  const numChildren = templateBody.getNumChildren();

  for (let i = 0; i < numChildren; i++) {
    const templateElement = templateBody.getChild(i);
    const copiedElement = templateElement.copy();

    // Handle tables separately
//    if (copiedElement.getType() === DocumentApp.ElementType.TABLE) {
//      mergeDocBody.appendTable(copiedElement);
//    } else {
      // Append other elements as paragraphs
//      mergeDocBody.appendParagraph(copiedElement);
//    }
  }

  processDocument(mergeDocBody, toMergeData);
}

function processDocument(mergeDocBody, toMergeData) {
  const numChildren = mergeDocBody.getNumChildren();
  for (let i = 0; i < numChildren; i++) {
    processElement(mergeDocBody.getChild(i), toMergeData);
  }
}

function processElement(element, toMergeData) {
  const elementType = element.getType();
  Logger.log("Processing element of type: " + elementType);

  switch (elementType) {
        case DocumentApp.ElementType.PARAGRAPH:
            element.setText(replacePlaceholders(element.getText(), toMergeData)); // Perform replacement on paragraph text
            Logger.log('Merged Paragraph:', element.asText().getText());
            mergeDocBody.appendParagraph(element); 
            break;
        case DocumentApp.ElementType.TABLE:
            mergeDocBody.appendTable(element);
            break;
        case DocumentApp.ElementType.LIST_ITEM:
            mergeDocBody.appendListItem(element);
            break;
        case DocumentApp.ElementType.HORIZONTAL_RULE:
            mergeDocBody.appendHorizontalRule();
            break; 
        case DocumentApp.ElementType.IMAGE:
            mergeDocBody.appendImage(element);
            break;
        case DocumentApp.ElementType.TABLE_CELL:
            console.log(type);
            mergeDocBody.appendTable(element);
            break;
        case DocumentApp.ElementType.PAGE_BREAK:
            mergeDocBody.appendPageBreak();
            break;
        default:
            Logger.log(`Element type:  ${elementType} not processed for replacements`);
  }
}

function replacePlaceholdersInText(textElement, toMergeData) {
  let text = textElement.getText();
  Logger.log("Original Text: " + text);

  for (let placeholder in toMergeData) {
    let value = toMergeData[placeholder] || "";
    let regex = new RegExp(
      "{{" + escapeRegExp(placeholder) + "}}",
      "g"
    );
    text = text.replace(regex, value);
  }

  Logger.log("Modified Text: " + text);
  textElement.setText(text);
}

// Helper function to escape special characters in regular expressions
function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); 
} 
