package symbolthree.oracle.excel;

import java.io.File;
import java.io.InputStream;
import java.util.List;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.jdom2.Document;
import org.jdom2.Element;
import org.jdom2.Namespace;
import org.jdom2.filter.Filters;
import org.jdom2.input.DOMBuilder;
import org.jdom2.input.SAXBuilder;
import org.jdom2.output.Format;
import org.jdom2.output.XMLOutputter;
import org.jdom2.xpath.XPathExpression;
import org.jdom2.xpath.XPathFactory;

public class ReadPOMTest {

	public static void main(String[] args) {
		ReadPOMTest test = new ReadPOMTest();
		test.run();
	}
	
	private void run() {
		try {
		DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
		Namespace ns = Namespace.getNamespace("ns", "http://maven.apache.org/POM/4.0.0");
		//Namespace ns = Namespace.getNamespace("");
		//factory.setNamespaceAware(true);
		DocumentBuilder documentBuilder = factory.newDocumentBuilder();
		org.w3c.dom.Document w3cDocument = documentBuilder.parse(new File("pom.xml"));
			
		//SAXBuilder jdomBuilder = new SAXBuilder();
		//Document jdomDoc = jdomBuilder.build(new File("pom.xml"));
		Document jdomDoc = new DOMBuilder().build(w3cDocument);
		
        XMLOutputter outputter = new XMLOutputter();
        outputter.setFormat(Format.getPrettyFormat());
        //System.out.println(outputter.outputString(jdomDoc));
        
		XPathFactory xpfac = XPathFactory.instance();
		XPathExpression<Element> xp = xpfac.compile("//ns:dependency", Filters.element(), null, ns);
	    List<Element> eles = xp.evaluate(jdomDoc);
	    for (Element ele : eles) {
	    	//List<Element> ele2 = ele.getChildren();
	    	String groupId = ele.getChild("groupId", ns).getText();	    	
	    	String version = ele.getChild("version", ns).getText();
	    	System.out.println(groupId + "/" + version);
	    	//System.out.println(outputter.outputString(ele2));
	    }
        
        //Element ele = jdomDoc.getRootElement().getChild("properties", ns);
        //System.out.println(outputter.outputString(ele));
        
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
