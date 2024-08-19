/**
 * see Apps Script reference
 * https://developers.google.com/apps-script/reference/document/element-type
 */
function logDocumentElements() {
  var document = DocumentApp.getActiveDocument(); // Or use .openById("DOCUMENT_ID") if not running the script in the bound document
  var body = document.getBody();
  
  // Start the recursive process with the body element
  logElement(body, 0); // The second argument represents the depth level for better readability in logging
}

function logElement(element, depth) {
  var elementType = element.getType();
  var prefix = " ".repeat(depth * 2); // Indentation for readability, based on the depth in the document structure
  
  // Log the element type (and text content for text elements)
  if (elementType == DocumentApp.ElementType.TEXT) {
    Logger.log(prefix + elementType + ": '" + element.getText().substring(0, 50) + "...'"); // Log the first 50 characters of text elements to avoid clutter
  } else {
    Logger.log(prefix + elementType);
  }
  
  // If the element can contain other elements, process each child element
  if (element.getNumChildren) {
    var numChildren = element.getNumChildren();
    for (var i = 0; i < numChildren; i++) {
      var child = element.getChild(i);
      logElement(child, depth + 1); // Recursive call with increased depth
    }
  }
}
