function mergeTemplate(mergeDoc, toMergeData) {
  try {
    const replacements = [];

    // Collect replacement pairs
    for (const [key, value] of Object.entries(toMergeData)) {
      /**
       * regular expression:
       * This regular expression (/\[/g) is used to escape square brackets ([) 
       * within the key variable, and it's applied conditionally using the ternary 
       * operator. Here's a breakdown:
       * Condition:
       * key.includes('['): Checks if the key string contains a square bracket character ([).
       * If True (key contains [):
       * key.replace(/\[/g, '\\\\\['):
       * This part uses the replace method on the key string.
       * /\[/g: This is the regular expression pattern.
       * \[: Matches a literal square bracket character ([).
       * g: Global flag to match all occurrences of [.
       * '\\\\\[': This is the replacement string.
       * \\: Escapes the backslash character to avoid confusion with special regex meanings.
       * \[: Represents the actual square bracket character to be inserted.
       * Result:
       * If the condition is true, all square brackets in the key string are replaced with \\\[,
       * effectively escaping them to prevent them from being interpreted as part of the regular
       * expression in subsequent steps.
       * If False (key does not contain [):
       * The key remains unchanged.
       */
      const escapedKey = key.includes('[') ? key.replace(/\[|\]/g, '\\\\$&') : key;
      replacements.push({ placeholder: `{{${escapedKey}}}`, value });
      // this generates an array of placeholders and values 
      // where the special characters are escaped
      // e.g. placeholder: "{{MOSAIC \\[ID\\]}}"
      // value: 33367690
    }

    let replacementMade = false;
    templateParagraphs.forEach(function(p) {
      var elementType = p.getType(); // Get the element type

      if (elementType == DocumentApp.ElementType.PARAGRAPH) { // Check for paragraph type
        var newParagraph = mergeDoc.getBody().appendParagraph(p.copy()); // Copy the paragraph
        // Perform text replacements within the new paragraph
        for (const replacement of replacements) {
          /**
           * regular expression:
           * Escapes special characters in the replacement value to avoid conflicts with the regex syntax.
           * \[\\&*\{\}|^$.]: Matches characters that need escaping (\, &, *, }, |, ^, $).
           * \\$&: Escapes the matched character (e.g., \\ becomes \\\\).
           * g: Global flag to escape all occurrences.
           */
//          var escapedValue = replacement.value.toString().replace(/\[\\&*{}|^$.\]/g, '\\$&');
          /**
           * regular expression:
           * Matches the placeholder syntax {{placeholderKey}}.
           * \{\{: Matches the literal opening curly braces.
           * ([^\}]+): Captures any character except } (non-greedy) into capturedKey.
           * \}\}: Matches the literal closing curly braces.
           * g: Global flag to match all occurrences in the string.
           */
          newParagraph.replaceText(/\{\{([^\}]+)\}\}/g, (match, capturedKey) => {
            console.log(`Match: ${match}, Captured Key: ${capturedKey}`);
            // Escape special characters in the captured key
            //const escapedCapturedKey = capturedKey.replace(/[\\^$.*+?()[\]{}|]/g, '\\$&');
            const paragraphReplacements = replacements.filter(
              (replacement) => replacement.placeholder.includes(capturedKey)
            );

            if (paragraphReplacements.length > 0) {
              // Implement your logic for handling multiple matches (e.g., sort by priority)
              const chosenReplacement = paragraphReplacements[0]; // Replace with first by default
              console.log(`Replaced "{{${capturedKey}}}" with: ${chosenReplacement.Value}`);
              return chosenReplacement.value;
            } else {
              console.log(`No replacement found for ${match}`);
              return match; // Return original match if no replacement found
            }
          });
        }
      }
    });

    mergeDoc.getBody().appendPageBreak();

    return;
  } catch (error) {
    console.error(`An error occurred: ${error}`);
    return error;
  }
}
