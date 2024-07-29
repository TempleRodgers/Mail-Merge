function mergeTemplate_(mergeDoc,toMergeData) {
  try {
    // Create an array to hold the replacement pairs
    const replacements = [];

    // Collect replacement pairs
    //replacements = [];
    for (const [key, value] of Object.entries(toMergeData)) {
      replacements.push({ placeholder: `{{${key.toUpperCase()}}}`, value });
    }

    templateParagraphs.forEach(function(p) {
      var elementType = p.getType(); // Get the element type

      if (elementType == DocumentApp.ElementType.PARAGRAPH) { // Check for paragraph type
        // Create a new paragraph in the merge document
        var newParagraph = mergeDoc.getBody().appendParagraph(p.copy()); // Copy the paragraph

        // Perform text replacements within the new paragraph
        for (const replacement of replacements) {
          newParagraph.replaceText(replacement.placeholder, replacement.value);
        }
      }
    })
    mergeDoc.getBody().appendPageBreak();

    return;
  } catch (error) {
    console.error(`An error occurred: ${error}`);
    return error;
  }
}
