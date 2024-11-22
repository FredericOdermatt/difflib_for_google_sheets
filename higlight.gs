function difflib_highlight() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getActiveRange();
  const values = range.getValues();

  const startColumn = range.getColumn();
  for (let row = 0; row < values.length; row++) {
    const col = 1; // Adjust if needed
    const cellValue = values[row][col];
    const referenceValue = values[row][col - 1];

    if (referenceValue && cellValue) {
      const matcher = new SequenceMatcher(null, referenceValue, cellValue);
      const opcodes = matcher.getOpcodes(); // Get matching blocks
      
      const richTextBuilder = SpreadsheetApp.newRichTextValue().setText(cellValue);

      opcodes.forEach(([tag, i1, i2, j1, j2]) => {
        let textStyle;
        if (tag === "replace" || tag === "delete" || tag === "insert") {
          textStyle = SpreadsheetApp.newTextStyle()
            .setForegroundColor("red") // Highlight differences in red
            .setBold(true) // Optional: Make differences bold
            .build();
        } else if (tag === "equal") {
          textStyle = SpreadsheetApp.newTextStyle()
            .setForegroundColor("black") // Matching text stays default
            .build();
        }

        if (tag !== "delete") {
          // Apply styles to matching range in cellValue
          const start = j1; // Start index in cellValue
          const end = Math.min(j2, cellValue.length); // End index
          if (start < end) {
            richTextBuilder.setTextStyle(start, end, textStyle);
          }
        }
      });

      // Apply the formatted text back to the cell
      const richText = richTextBuilder.build();
      range.offset(row, col).setRichTextValue(richText);
    }
  }
}

function testDifflibIntegration() {
  const matcher = new SequenceMatcher(null, "hello", "hallo");
  Logger.log(matcher.getOpcodes());
}