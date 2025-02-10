function parseXml(xmlString) {
    const parser = new DOMParser();
    return parser.parseFromString(xmlString, "application/xml");
  }

  function deepCompareNodes(node1, node2) {
    if (node1.nodeType !== node2.nodeType) {
      console.log("First Comparison");
      return false;
    }
  
    if (node1.nodeType === Node.TEXT_NODE) {
      console.log("Second Comparison");
      return node1.nodeValue.trim() === node2.nodeValue.trim();
    }
  
    if (node1.nodeName !== node2.nodeName) {
      console.log("Third Comparison");
      return false;
    }
  
    if (node1.attributes.length !== node2.attributes.length) {
      console.log("Fourth Comparison");
      return false;
    }
  
    for (let i = 0; i < node1.attributes.length; i++) {
      console.log("Inside the for iteration for comparison");
      const attr1 = node1.attributes[i];
      const attr2 = node2.attributes.getNamedItem(attr1.name);
      if (!attr2 || attr1.value !== attr2.value) {
        console.log("Inside the for iteration for comparison");
        return false;
      }
    }
  
    if (node1.childNodes.length !== node2.childNodes.length) {
      console.log("Fifth Comparison");
      return false;
    }
  
    for (let i = 0; i < node1.childNodes.length; i++) {
      if (!deepCompareNodes(node1.childNodes[i], node2.childNodes[i])) {
        return false;
      }
    }
  
    return true;
  }

  function compareXml(xml1, xml2) {
    const parsedXml1 = parseXml(xml1);
    const parsedXml2 = parseXml(xml2);
  
    return deepCompareNodes(parsedXml1.documentElement, parsedXml2.documentElement);
  }

  module.exports = compareXml