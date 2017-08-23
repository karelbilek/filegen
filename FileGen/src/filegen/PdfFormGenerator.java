package filegen;

import com.fasterxml.jackson.core.type.TypeReference;
import com.itextpdf.text.pdf.AcroFields;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;
import com.itextpdf.text.pdf.XfaForm;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;


public class PdfFormGenerator extends ObecnyGenerator {
    
    public NodeList getXmlNodeFromString(String s) throws ParserConfigurationException, SAXException, IOException {
        DocumentBuilder newDocumentBuilder = DocumentBuilderFactory.newInstance().newDocumentBuilder();
        Document parse = newDocumentBuilder.parse(new ByteArrayInputStream(s.getBytes("UTF-8")));
        return parse.getFirstChild().getChildNodes();
    }
    
    public PdfFormGenerator(String templateName, String finalName, String finalFormat, String jsonstring) {
        super(templateName,finalName,finalFormat,jsonstring);
    }

    protected String getNativeFormat() {
        return "pdf";
    }
    
    private void fillNode(PdfFormData data, Node node, String path) throws ParserConfigurationException, SAXException, IOException {
        NodeList childNodes = node.getChildNodes();
        for (int i = 0; i < childNodes.getLength(); i++) {
            Node childNode = childNodes.item(i);
            String childNodePath = path + "/" + childNode.getLocalName();
            if (data.strings.containsKey(childNodePath)) {
                childNode.setTextContent(data.strings.get(childNodePath));
                data.strings.remove(childNodePath); // HACK
            } else if (data.xmlStrings.containsKey(childNodePath)) {
                String s = data.xmlStrings.get(childNodePath);
                NodeList nl = getXmlNodeFromString(s);
                while (childNode.hasChildNodes()){
                    childNode.removeChild(childNode.getFirstChild());
                }
                for (int j = 0; j < nl.getLength(); j++) {
                    Node added = childNode.getOwnerDocument().importNode(nl.item(j), true);
                    childNode.appendChild(added);
                }
            }  else {
                this.fillNode(data, childNode, childNodePath);
            }
        }
    }

    protected File generateNatural(File templateName, Object mapO) throws Exception {
        PdfFormData data = (PdfFormData) mapO;
        
        PdfReader reader = new PdfReader(new FileInputStream(templateName));
        File resFile = File.createTempFile("PdFormTemp", ".pdf");
        
        PdfStamper stamper = new PdfStamper(reader,
                new FileOutputStream(resFile)
                , '\0', true);
        
        AcroFields fields = stamper.getAcroFields();
        XfaForm xfa = fields.getXfa();
        
        Node datasets_node = xfa.getDatasetsNode();
        Node data_node = datasets_node.getChildNodes().item(0);
        Node formname_node = data_node.getChildNodes().item(0);
        Node form_node = formname_node.getChildNodes().item(0);
        
        this.fillNode(data, form_node, "");
        xfa.fillXfaForm(form_node);
        
        stamper.close();
        reader.close();
        return resFile;
    }
    
    protected TypeReference getType() {
        return new TypeReference<PdfFormData>() {
        };
    }

}
