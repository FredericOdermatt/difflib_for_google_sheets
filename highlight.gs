function highlightDifferences() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getDataRange(); // Assumes the data range contains the text to compare
  const values = range.getValues();

  for (let row = 0; row < values.length; row++) {
    for (let col = 1; col < values[row].length; col++) {
      const cellValue = values[row][col];
      const referenceValue = values[row][col - 1]; // Compare with the previous column

      if (referenceValue && cellValue) {
        const differences = getDiffIndexes(referenceValue, cellValue);
        const richTextBuilder = SpreadsheetApp.newRichTextValue().setText(cellValue);

        differences.forEach(([start, end]) => {
          const textStyle = SpreadsheetApp.newTextStyle()
            .setForegroundColor("red") // Highlight differences in red
            .build();
          richTextBuilder.setTextStyle(start, end, textStyle);
        });

        const richText = richTextBuilder.build();
        sheet.getRange(row + 1, col + 1).setRichTextValue(richText); // Set the styled text
      }
    }
  }
}

/**
 * Get the indexes of differences between two strings.
 * @param {string} text1 - The reference text.
 * @param {string} text2 - The text to compare.
 * @returns {Array} Array of [start, end] pairs for differences.
 */
function getDiffIndexes(text1, text2) {
  const diffIndexes = [];
  const maxLength = Math.max(text1.length, text2.length);

  let start = null;

  for (let i = 0; i < maxLength; i++) {
    if (text1[i] !== text2[i]) {
      if (start === null) start = i; // Start of a difference
    } else if (start !== null) {
      diffIndexes.push([start, i]); // End of a difference
      start = null;
    }
  }

  if (start !== null) {
    diffIndexes.push([start, maxLength]); // Handle trailing differences
  }

  return diffIndexes;
}
