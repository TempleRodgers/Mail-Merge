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
        mergeDocBody.appendTable(element);
        break;
      case DocumentApp.ElementType.HORIZONTAL_RULE:
//        console.log("Appending Horizontal Rule"); // No specific value to log
        mergeDocBody.appendHorizontalRule();
        break;
      case DocumentApp.ElementType.INLINE_IMAGE:
//        console.log("Appending Inline Image:", element.asInlineImage().getBlob().getDataAsString());
        mergeDocBody.appendImage(elementasInlineImage().getBlob());
        break;
      case DocumentApp.ElementType.PAGE_BREAK:
//        console.log("Appending Page Break"); // No specific value to log
        mergeDocBody.appendPageBreak();
        break;
      default:
        // Handle other element types as needed
        console.log("Unsupported element type:", element.getType());
        break;
    }
  }
  console.log("---- Exiting mergeTemplate ----");
}
