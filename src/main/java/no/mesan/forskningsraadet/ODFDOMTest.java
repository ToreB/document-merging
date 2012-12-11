package no.mesan.forskningsraadet;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.xml.namespace.NamespaceContext;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathFactory;

import org.odftoolkit.odfdom.doc.OdfDocument;
import org.odftoolkit.odfdom.doc.OdfTextDocument;
import org.odftoolkit.odfdom.dom.OdfContentDom;
import org.odftoolkit.odfdom.dom.element.OdfStyleBase;
import org.odftoolkit.odfdom.dom.element.office.OfficeTextElement;
import org.odftoolkit.odfdom.dom.element.text.TextPElement;
import org.odftoolkit.odfdom.dom.style.props.OdfTextProperties;
import org.odftoolkit.odfdom.incubator.doc.office.OdfOfficeStyles;
import org.odftoolkit.odfdom.incubator.doc.text.OdfTextHeading;
import org.odftoolkit.odfdom.incubator.doc.text.OdfTextParagraph;
import org.odftoolkit.odfdom.incubator.search.TextNavigation;
import org.odftoolkit.odfdom.incubator.search.TextSelection;
import org.odftoolkit.odfdom.pkg.OdfElement;
import org.odftoolkit.odfdom.pkg.OdfNamespace;

public class ODFDOMTest {

	public static void main(String[] args) {
		String path = "/Users/toreb/Documents/Forskningsraadet/";
		/*String layoutTemplate = path + "template2.odt";
		String output = path + "dokument2.odt";
		createDocumentFromTemplate(layoutTemplate, output);*/
		
		generateContentTemplateWithData(path + "IMMedInnhold.odt");
	}
	
	private static void createDocumentFromTemplateWithContent(String layoutTemplatePath, String contentTemplatePath, String outputPath) {
		Map<String, String> replacements = new HashMap<String, String>();
		replacements.put("<%TITLE%>", "Min Tittel");
		replacements.put("<%CONTENT_PAGE1%>", "Dette er innholdet i dokumentet på side 1.");
		replacements.put("<%CONTENT_PAGE2%>", "Dette er innholdet i dokumentet på side 2.");
		replacements.put("<%HEADING_PAGE1%>", "Dette er en heading 2 på side 1");
		replacements.put("<%HEADING_PAGE2%>", "Dette er en heading 1 på side 2");
		replacements.put("<%HEADER%>", "Dette er en header!");
		replacements.put("<%FOOTER%>", "Dette er en footer");
		
		try {		
			OdfTextDocument layoutTemplateDoc = OdfTextDocument.loadDocument(new File(layoutTemplatePath));			
			System.out.println("Lastet inn utseendemal: " + layoutTemplatePath);
			
			OdfTextDocument contentTemplateDoc = OdfTextDocument.loadDocument(new File(contentTemplatePath));
			
			OfficeTextElement contents = contentTemplateDoc.getContentRoot();
			/*XPath xpath = XPathFactory.newInstance().newXPath();
			TextPElement element = (TextPElement) xpath.evaluate("//text:p", contents, XPathConstants.NODE);
			System.out.println(element);*/
			
			System.out.println(OdfElement.findFirstChildNode(TextPElement.class, contents));
			
			String regex = "<%([a-zA-Z0-9_#\\-]+)%>";
			TextNavigation search = new TextNavigation(regex, layoutTemplateDoc);

			while(search.hasNext()) {
				TextSelection item = (TextSelection) search.getCurrentItem();
				String text = item.getText();
				
				/*Pattern pattern = Pattern.compile(regex);
				Matcher matcher = pattern.matcher(item.getText());
				matcher.find();
				String text = matcher.group(1);*/
				
				if (replacements.containsKey(text)) {
					item.replaceWith(replacements.get(text));
				}
			}
			
			layoutTemplateDoc.save(outputPath);
			System.out.println("Dokument lagret som: " + outputPath);
			
			layoutTemplateDoc.close();
			
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	private static void createDocumentFromTemplate(String templatePath, String outputPath) {
		Map<String, String> replacements = new HashMap<String, String>();
		replacements.put("<%TITLE%>", "Min Tittel");
		replacements.put("<%CONTENT_PAGE1%>", "Dette er innholdet i dokumentet på side 1.");
		replacements.put("<%CONTENT_PAGE2%>", "Dette er innholdet i dokumentet på side 2.");
		replacements.put("<%HEADING_PAGE1%>", "Dette er en heading 2 på side 1");
		replacements.put("<%HEADING_PAGE2%>", "Dette er en heading 1 på side 2");
		replacements.put("<%HEADER%>", "Dette er en header!");
		replacements.put("<%FOOTER%>", "Dette er en footer");
		
		try {		
			OdfTextDocument layoutTemplateDoc = OdfTextDocument.loadDocument(new File(templatePath));
			
			System.out.println("Lastet inn utseendemal: " + templatePath);
			
			/*String contentTemplate = path + "InnholdsMalMedInnhold.odt";
			OdfTextDocument contentTemplateDoc = OdfTextDocument.loadDocument(new File(contentTemplate));
			
			OfficeTextElement contents = contentTemplateDoc.getContentRoot();
			XPath xpath = XPathFactory.newInstance().newXPath();
			TextPElement element = (TextPElement) xpath.evaluate("//text:p[3]", contents, XPathConstants.NODE);
			System.out.println(element);
			
			System.out.println(OdfElement.findFirstChildNode(TextPElement.class, contents));*/
			
			String regex = "<%([a-zA-Z0-9_#\\-]+)%>";
			TextNavigation search = new TextNavigation(regex, layoutTemplateDoc);

			while(search.hasNext()) {
				TextSelection item = (TextSelection) search.getCurrentItem();
				String text = item.getText();
				
				/*Pattern pattern = Pattern.compile(regex);
				Matcher matcher = pattern.matcher(item.getText());
				matcher.find();
				String text = matcher.group(1);*/
				
				if (replacements.containsKey(text)) {
					item.replaceWith(replacements.get(text));
				}
			}
			
			layoutTemplateDoc.save(outputPath);
			System.out.println("Dokument lagret som: " + outputPath);
			
			layoutTemplateDoc.close();
			
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	private static void generateContentTemplateWithData(String filepath) {
		try {	
			//New document that contains one paragraph by default
			OdfDocument doc = OdfTextDocument.newTextDocument();
			OdfContentDom dom = doc.getContentDom();
			OfficeTextElement officeText = (OfficeTextElement) doc.getContentRoot();
			OdfOfficeStyles styles = doc.getOrCreateDocumentStyles();
			
			//Inserts a heading to be used as document title
			OdfTextHeading heading = new OdfTextHeading(dom);
			heading.addStyledContent("Heading_20_1", "Document title");		
			officeText.appendChild(heading);		
			
			String headingName = "Heading_20_2";
			String headingText = "Paragraph";
			String content = 
					"Lorem ipsum dolor sit amet, consectetur adipiscing elit. Fusce sit amet dolor felis. " +
					"Donec mollis auctor euismod. Aliquam suscipit porta odio vitae vehicula. Donec nec tortor sit amet erat " +
					"hendrerit dignissim. Aliquam massa enim, pellentesque non lobortis vel, feugiat eget arcu. Fusce posuere " +
					"sapien a neque tristique commodo. Suspendisse potenti. Ut commodo ante sed dolor dictum placerat. Curabitur " +
					"vehicula, nulla at viverra consequat, metus leo fringilla est, ac consectetur risus sem et augue. " +
					"Sed metus ipsum, elementum id egestas vitae, mattis sed neque. Morbi odio turpis, tristique id tincidunt vitae, " +
					"lobortis et ligula. In dapibus laoreet lacus sed lacinia.";
			
			//Inserts a few paragraphs
			for(int i = 1; i <= 20; i++) {
				//Creates a paragraph
				OdfTextParagraph paragraph = new OdfTextParagraph(dom);
				
				//Adds a heading
				OdfTextHeading paraHeading = new OdfTextHeading(dom);
				paraHeading.addStyledContent(headingName, headingText + i);
				officeText.appendChild(paraHeading);
				
				//Adds the content
				paragraph.addContent(content);
				
				//Appends the paragraph to the root node
				officeText.appendChild(paragraph);
			}
			
			doc.save(filepath);
			System.out.println("Dokument lagret: " + filepath);
			
			doc.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	private static void setFontSize(OdfStyleBase style, String value)
	{
	    style.setProperty(OdfTextProperties.FontSize, value);
	    style.setProperty(OdfTextProperties.FontSizeAsian, value);
	    style.setProperty(OdfTextProperties.FontSizeComplex, value);
	}
}
