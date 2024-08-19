function replacePlaceholders(element, regex, substitutionMap) {
    Logger.log(`element type: ${element.getType()}`);
    // Perform replacements based on the type of element
    switch (element.getType()) {
        case DocumentApp.ElementType.PARAGRAPH:
            return replaceParagraph(element.asParagraph(), regex, substitutionMap);
        case DocumentApp.ElementType.TEXT:
            const text = element.asText();
            const textContent = text.getText();
            if (textContent) {
                Logger.log("Text Content:", textContent);
                Logger.log("Type of Text Content:", typeof textContent);
                return replaceText(textContent, regex, substitutionMap);
            } else {
                Logger.log("Text Content is empty.");
                return ""; // Return empty string if text content is empty
            }
        case DocumentApp.ElementType.TABLE:
            return replaceTable(element.asTable(), regex, substitutionMap);
        case DocumentApp.ElementType.LIST_ITEM:
            return replaceListItem(element.asListItem(), regex, substitutionMap);
        // Add cases for other element types as needed
        default:
            // For unsupported element types, return the element as is
            return element;
    }
}

function replaceText(cellContent, regex, substitutionMap) {
    // Use the replace() method with a callback function
    const replacedContent = cellContent.replace(regex, function(match, placeholder) {
        // Check if the placeholder exists in the substitution map
        if (substitutionMap.hasOwnProperty(placeholder.trim())) {
            // If it exists, return the corresponding substitution text
            Logger.log(`Text replacement: ${substitutionMap[placeholder.trim()]}`);
            return substitutionMap[placeholder.trim()];
        } else {
            // If not, return the original match
            return match;
        }
    });
    return replacedContent;
}

function replaceParagraph(paragraph, regex, substitutionMap) {
    // Iterate over the elements in the paragraph
    const numElements = paragraph.getNumChildren();
    for (let i = 0; i < numElements; i++) {
        const child = paragraph.getChild(i);
        // Perform replacements recursively for each child element
        replacePlaceholders(child, regex, substitutionMap);
    }
}

function replaceTable(table, regex, substitutionMap) {
    // Iterate over rows and cells in the table
    const numRows = table.getNumRows();
    for (let i = 0; i < numRows; i++) {
        const row = table.getRow(i);
        const numCells = row.getNumCells();
        for (let j = 0; j < numCells; j++) {
            const cell = row.getCell(j);
            const cellContent = cell.getText(); // Extract text content from the cell
            console.log("Cell Content:", cellContent);
            console.log("Type of Cell Content:", typeof cellContent);
            // Perform replacements recursively for each cell's content
            const replacedContent = replaceText(cellContent, regex, substitutionMap);
            // Set the replaced text back to the cell
            cell.setText(replacedContent);
        }
    }
}


function replaceListItem(listItem, regex, substitutionMap) {
    // Perform replacements recursively for each child element in the list item
    const numChildren = listItem.getNumChildren();
    for (let i = 0; i < numChildren; i++) {
        const child = listItem.getChild(i);
        replacePlaceholders(child, substitutionMap);
    }
}
