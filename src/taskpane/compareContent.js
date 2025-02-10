function parseXml(xmlString) {
    const parser = new DOMParser();
    return parser.parseFromString(xmlString, "application/xml");
  }
  
  function extractTextContent(node) {
    let textContent = "";
    if (node.nodeType === Node.TEXT_NODE) {
      textContent += node.nodeValue.trim();
    }
    for (let child of node.childNodes) {
      textContent += extractTextContent(child);
    }
    return textContent;
  }
  
  function compareXmlTextContent(xml1, xml2) {
    const parsedXml1 = parseXml(xml1);
    const parsedXml2 = parseXml(xml2);
  
    const textContent1 = extractTextContent(parsedXml1.documentElement);
    const textContent2 = extractTextContent(parsedXml2.documentElement);
  
    return textContent1 === textContent2;
  }


  module.exports = compareXmlTextContent;