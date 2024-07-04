function getTopLevelIndent(paragraph) {
  var indentFirstLine = paragraph.getIndentFirstLine();
  var indentStart = paragraph.getIndentStart();
  
  // Determine the top-level indent based on the smallest indent
  var topLevelIndent = Math.min(indentFirstLine, indentStart);
  
  return topLevelIndent;
}

function parseTOC( headerRange) {

  var contents = [];
  var doc = DocumentApp.getActiveDocument();

  // Define the search parameters.
  var searchElement  = doc.getBody();
  var searchType = DocumentApp.ElementType.TABLE_OF_CONTENTS;

  // Search for TOC. Assume there's only one.
  var searchResult = searchElement.findElement(searchType);

  if (searchResult) {
    // TOC was found
    var toc = searchResult.getElement().asTableOfContents();

    // Parse all entries in TOC. The TOC contains child Paragraph elements,
    // and each of those has a child Text element. The attributes of both
    // the Paragraph and Text combine to make the TOC item functional.
    var numChildren = toc.getNumChildren();
    var skip = true;
    var index = 0
    for (var i=0; i < numChildren; i++) {
      var itemInfo = {}
      var tocItem = toc.getChild(i).asParagraph();
      var topLevelIndent = getTopLevelIndent(tocItem);
      var tocItemText = tocItem.getChild(0).asText();
      
      if(tocItem.getText() == headerRange[index][0]){
        skip = false;
        // exclusive start
        continue;
      }
      if(tocItem.getText() == headerRange[index][1]){
        skip = true;
        index += 1;
      }
      if(skip) continue;

      // 18 is a single indent which all the features happen to be by format.
      // if( topLevelIndent != 18) continue;
      Logger.log("tocItem: " + tocItem.getText());
      // skip if top level indent


      // Set itemInfo attributes for this TOC item, first from Paragraph
      itemInfo.text = tocItem.getText();                // Displayed text
      itemInfo.indentStart = tocItem.getIndentStart();  // TOC Indentation
      itemInfo.indentFirstLine = tocItem.getIndentFirstLine();
      // ... then from child Text
      itemInfo.linkUrl = tocItemText.getLinkUrl();      // URL Link in document
      contents.push(itemInfo);
    }
  }

  // Return array of objects containing TOC info
  return contents;
}

function LinkHeaders() {
  // It'll use these headers as boundaries to start and stop parsing
  startStopPairs = [ ["Weapon Features", "Armor & Shields"], ["Armor Features","Reload Table"]]
  var contents = parseTOC(startStopPairs);
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var paragraphs = body.getParagraphs();

  for (var i = 0; i < contents.length; i++) {
    var ctx = contents[i];
    Logger.log("ctx: " + ctx.text);
    for (var j = 0; j < paragraphs.length; j++) { // Use j instead of i here
      var paragraph = paragraphs[j];
      var textElement = paragraph.editAsText();
      // skipp table of contents
      if (textElement.getParent() == 'TableOfContents') continue;
      // skip if header2
      if (textElement.getFontSize() == 16 || textElement.getFontSize() == null) continue;
      var searchResult = textElement.findText(ctx.text);

  
      while (searchResult != null ) {
        var startOffset = searchResult.getStartOffset();
        var endOffset = searchResult.getEndOffsetInclusive();
        // Logger.log("matched: " + textElement.getText().substring(startOffset, endOffset + 1));

        // Remove any existing links in the range
        textElement.setLinkUrl(startOffset, endOffset, null);

        // Set the new link
        textElement.setLinkUrl(startOffset, endOffset, ctx.linkUrl);

        // Find the next occurrence within the current textElement
        searchResult = textElement.findText(ctx.text, searchResult);
      }
    }
  }
}

