function convertMarkdownBold() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getDataRange();
  const richTextValues = range.getRichTextValues();
  
  const newRichTextValues = richTextValues.map(row => {
    return row.map(cell => {
      const originalText = cell.getText();
      if (!originalText) return cell;

      // This regex finds the **text** and captures the 'text' inside
      const boldRegex = /\*\*(.*?)\*\*/g;
      let match;
      
      // We'll build the clean text by simply removing all **
      const cleanText = originalText.replace(/\*\*/g, '');
      const builder = SpreadsheetApp.newRichTextValue().setText(cleanText);
      const boldStyle = SpreadsheetApp.newTextStyle().setBold(true).build();

      // We need to track how many asterisks we've skipped to find the new positions
      let asterisksRemoved = 0;

      while ((match = boldRegex.exec(originalText)) !== null) {
        // match[0] is "**text**", match[1] is "text"
        // match.index is the start position in the ORIGINAL string
        
        const startInClean = match.index - asterisksRemoved;
        const endInClean = startInClean + match[1].length;

        builder.setTextStyle(startInClean, endInClean, boldStyle);
        
        // Every time we process a match, we've bypassed 4 asterisks (2 start, 2 end)
        // that won't exist in the clean string.
        asterisksRemoved += 4;
      }

      return builder.build();
    });
  });

  range.setRichTextValues(newRichTextValues);
}