function compareWordXMLText(xmlStr1, xmlStr2) {
    // Parse the XML strings into document objects
    const parser = new DOMParser();
    const xmlDoc1 = parser.parseFromString(xmlStr1, "application/xml");
    const xmlDoc2 = parser.parseFromString(xmlStr2, "application/xml");
  
    // Extract text content from <w:t> elements
    const text1 = extractTextFromWordXML(xmlDoc1);
    const text2 = extractTextFromWordXML(xmlDoc2);
  
    // Compare the extracted text content
    if (text1 === text2) {
      console.log("The visible content in the Word documents is identical.");
    } else {
      console.log("The visible content in the Word documents is different.");
    }
  }
  
  // Function to extract text from <w:t> elements
  function extractTextFromWordXML(xmlDoc) {
    const wTElements = xmlDoc.getElementsByTagName('w:t');
    let textContent = '';
    for (let i = 0; i < wTElements.length; i++) {
      textContent += wTElements[i].textContent; // Concatenate the text content from each <w:t> element
    }
    return textContent.trim(); // Remove extra spaces
  }

module.exports = compareWordXMLText;