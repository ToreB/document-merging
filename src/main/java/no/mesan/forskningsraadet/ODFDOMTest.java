package no.mesan.forskningsraadet;

import java.io.File;
import java.util.HashMap;
import java.util.Map;

import org.odftoolkit.odfdom.doc.OdfDocument;
import org.odftoolkit.odfdom.doc.OdfTextDocument;
import org.odftoolkit.odfdom.dom.OdfContentDom;
import org.odftoolkit.odfdom.dom.element.OdfStyleBase;
import org.odftoolkit.odfdom.dom.element.office.OfficeTextElement;
import org.odftoolkit.odfdom.dom.element.text.TextSectionElement;
import org.odftoolkit.odfdom.dom.style.props.OdfTextProperties;
import org.odftoolkit.odfdom.incubator.doc.office.OdfOfficeStyles;
import org.odftoolkit.odfdom.incubator.doc.text.OdfTextHeading;
import org.odftoolkit.odfdom.incubator.doc.text.OdfTextParagraph;
import org.odftoolkit.odfdom.incubator.search.TextNavigation;
import org.odftoolkit.odfdom.incubator.search.TextSelection;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import sun.misc.Cleaner;

public class ODFDOMTest {

	public static void main(String[] args) {
		String path = "/Users/toreb/Documents/Forskningsraadet/";
		String layoutTemplate = path + "template3.odt";
		String contentTemplate = path + "IMMedInnhold.odt";
		String output = path + "dokument3.odt";
		
		//generateContentTemplateWithData(path + "IMMedInnhold.odt");
		
		createDocumentFromTemplateWithContent(layoutTemplate, contentTemplate, output);
	}
	
	private static void createDocumentFromTemplateWithContent(String layoutTemplatePath, 
																String contentTemplatePath, 
																String outputPath) {
		
		Map<String, String> replacements = new HashMap<String, String>();
		replacements.put("<%TITLE%>", "Min Tittel");
		replacements.put("<%HEADER_MIDDLE%>", "Dette er en header!");
		replacements.put("<%HEADER_LEFT%>", "Venstre side i header");
		replacements.put("<%HEADER_RIGHT%>", "Høyre side i header!");
		replacements.put("<%FOOTER%>", "Dette er en footer");
		
		try {		
			OdfTextDocument layoutTemplateDoc = OdfTextDocument.loadDocument(new File(layoutTemplatePath));
			OfficeTextElement layoutRoot = layoutTemplateDoc.getContentRoot();
			System.out.println("Lastet inn utseendemal: " + layoutTemplatePath);
			
			OdfTextDocument contentTemplateDoc = OdfTextDocument.loadDocument(new File(contentTemplatePath));
			System.out.println("Lastet inn innholdsmal: "+ contentTemplatePath);
			
			OfficeTextElement contents = contentTemplateDoc.getContentRoot();
			
			/*XPath xpath = XPathFactory.newInstance().newXPath();
			TextPElement element = (TextPElement) xpath.evaluate("//text:p", contents, XPathConstants.NODE);
			System.out.println(element);*/		
			
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
				} else if (text.equals("<%CONTENT%>")) {
					item.replaceWith("");
					
					//TextSectionElement section = new TextSectionElement(layoutTemplateDoc.getContentDom());
					//layoutTemplateDoc.getContentRoot().appendChild(section);
					//section.setTextNameAttribute("Content from IM");
					
					NodeList nodes = contents.getChildNodes();
					for(int i = 0; i < nodes.getLength(); i++) {
						Node newNode = nodes.item(i).cloneNode(true);
						Node adoptedNode = layoutTemplateDoc.getContentDom().adoptNode(newNode);
						//section.appendChild(adoptedNode);
						layoutRoot.appendChild(adoptedNode);
					}
				}
			}
			
			layoutTemplateDoc.save(outputPath);
			System.out.println("Dokument lagret som: " + outputPath);
			
			layoutTemplateDoc.close();
			contentTemplateDoc.close();
			
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
			
			cleanOutDocument(officeText);
			
			/*//Inserts a heading to be used as document title
			OdfTextHeading heading = new OdfTextHeading(dom);
			heading.addStyledContent("Heading_20_1", "Document title");		
			officeText.appendChild(heading);*/
			
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
	
	private static void setFontSize(OdfStyleBase style, String value) {
	    style.setProperty(OdfTextProperties.FontSize, value);
	    style.setProperty(OdfTextProperties.FontSizeAsian, value);
	    style.setProperty(OdfTextProperties.FontSizeComplex, value);
	}
	
	private static void cleanOutDocument(OfficeTextElement element) {
	    Node childNode = element.getFirstChild();
	    
	    while (childNode != null)
	    {
	        element.removeChild(childNode);
	        childNode = element.getFirstChild();
	    }
	}
}
