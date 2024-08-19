/**
 * see Apps Script reference
 * https://developers.google.com/apps-script/reference/document/element-type
 */
function replaceTextOccurrences() {
  var document = DocumentApp.getActiveDocument();
  var body = document.getBody();
  body.replaceText("{{CH/PN}}", "--CH/PN-- WAS SUBSTITUTED");
  body.replaceText("{{FIRST NAME \\[2\\]}}", "--FIRST NAME [2]-- WAS SUBSTITUTED");
}
