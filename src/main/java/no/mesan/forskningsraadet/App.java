package no.mesan.forskningsraadet;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;

import org.docx4j.model.structure.HeaderFooterPolicy;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.Text;

/*
 * Some helper methods found in this example: 
 * http://www.javacodegeeks.com/2012/07/java-word-docx-documents-with-docx4j.html
 */
public class App 
{
/*
	private static final int HEADER = 1;
	private static final int FOOTER = 2;
	
    public static void main( String[] args )
    {
        try {
			WordprocessingMLPackage template = getTemplate("/Users/toreb/Documents/Forskningsraadet/template.docx");
			
			replaceHeaderFooterPlaceholder(template, "HEADER_PLACEHOLDER", "Dette er en header!", HEADER);
			replaceHeaderFooterPlaceholder(template, "FOOTER_PLACEHOLDER", "Dette er en footer", FOOTER);
			
			Map<String, String> replacements = new HashMap<String, String>();
			replacements.put("TITLE_PLACEHOLDER", "Min Tittel");
			replacements.put("CONTENT_PAGE1_PLACEHOLDER", "Dette er innholdet i dokumentet på side 1.");
			replacements.put("CONTENT_PAGE2_PLACEHOLDER", "Dette er innholdet i dokumentet på side 2.");
			replacements.put("BOLD_TEXT_PLACEHOLDER", "Dette skal komme i fet skrift");
			replacements.put("BOLD_ITALIC_UNDERLINED_TEXT_PLACEHOLDER", "Dette skal komme i bold, italic, underlined skrift");
			replacements.put("HEADING_PAGE1_PLACEHOLDER", "Dette er en heading 2 på side 1");
			replacements.put("HEADING_PAGE2_PLACEHOLDER", "Dette er en heading 1 på side 2");
			
			replacePlaceholder(template, replacements);
			
			writeDocxToStream(template, "/Users/toreb/Documents/Forskningsraadet/dokument.docx");
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (Docx4JException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
    }
    
	private static WordprocessingMLPackage getTemplate(String filename) throws Docx4JException, FileNotFoundException {	
		//loads a document
		WordprocessingMLPackage template = WordprocessingMLPackage.load(new FileInputStream(new File(filename)));
		
		return template;
	}
	
	private static List<Object> getAllElementFromObject(Object obj, Class<?> toSearch) {
		List<Object> result = new ArrayList<Object>();
		if (obj instanceof JAXBElement)
			obj = ((JAXBElement<?>) obj).getValue();

		if (obj.getClass().equals(toSearch))
			result.add(obj);
		else if (obj instanceof ContentAccessor) {
			List<?> children = ((ContentAccessor) obj).getContent();
			for (Object child : children) {
				result.addAll(getAllElementFromObject(child, toSearch));
			}
		}
		
		return result;
	}
	
	private static void replacePlaceholder(WordprocessingMLPackage template, Map<String, String> replacements) {		
		List<Object> texts = getAllElementFromObject(
				template.getMainDocumentPart(), Text.class);
		
		for (Object text : texts) {
			Text textElement = (Text) text;
			String value = textElement.getValue();
			
			if (replacements.isEmpty()) break;
			
			if (replacements.keySet().contains(value)) {
				String newValue = replacements.get(value);
				textElement.setValue(newValue);				
				replacements.remove(value);
			}
		}
	}
	
	private static void replacePlaceholder(WordprocessingMLPackage template, String placeholder, String replacementText) {		
		List<Object> texts = getAllElementFromObject(
				template.getMainDocumentPart(), Text.class);
		
		for (Object text : texts) {
			Text textElement = (Text) text;
			if (textElement.getValue().equals(placeholder)) {
				textElement.setValue(replacementText);
			}
		}
	}
	
	private static void replaceHeaderFooterPlaceholder(WordprocessingMLPackage template, String placeholder, String replacementText, int option) {
		List<SectionWrapper> sectionWrappers = template.getDocumentModel().getSections();
		org.docx4j.wml.ObjectFactory factory = new org.docx4j.wml.ObjectFactory();
		
		for (SectionWrapper sw : sectionWrappers) {
			HeaderFooterPolicy hfp = sw.getHeaderFooterPolicy();
			
			Object element = null;
			if (option == HEADER) {
				element = hfp.getDefaultHeader();
			} else {
				element = hfp.getDefaultFooter();
			}

			String xpath = "//w:r[w:t[contains(text(),'" + placeholder + "')]]";
			List<Object> list = null;
			try {
				if (option == HEADER)
					list = ((HeaderPart) element).getJAXBNodesViaXPath(xpath, false);
				else
					list = ((FooterPart) element).getJAXBNodesViaXPath(xpath, false);
			} catch (JAXBException e) {
				e.printStackTrace();
				list = new ArrayList<Object>();
			}
			
			for (int i = 0; i < list.size(); i++) {
				org.docx4j.wml.R r = (org.docx4j.wml.R) list.get(i);
				org.docx4j.wml.P parent = (org.docx4j.wml.P) r.getParent();
				org.docx4j.wml.RPr rpr = r.getRPr();
				int index = parent.getContent().indexOf(r);
				parent.getContent().remove(r);

				org.docx4j.wml.Text addedTmpText = factory.createText();
				org.docx4j.wml.R mainR = factory.createR();
				mainR.setRPr(rpr);
				addedTmpText.setValue(replacementText);
				mainR.getContent().add(addedTmpText);
				parent.getContent().add(index, mainR);
			}
		}
	}
	
	private static void writeDocxToStream(WordprocessingMLPackage template, String target) 
			throws IOException, Docx4JException {
		File f = new File(target);
		template.save(f);
	}
*/
}