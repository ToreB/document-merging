package no.mesan.forskningsraadet;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.docx4j.openpackaging.exceptions.Docx4JException;

public class App {
    public static void main( String[] args ) {
		String path = "/media/sf_DATA_DRIVE/Documents/Jobb/Forskningsraadet/";
		String layout = path + "template.docx";
		String contents = path + "IMMedInnhold.docx";
    	
        try {
        	//DocxDocument docx = replaceStyledPlaceholders(path + "template2.docx");
        	
			//DocxDocument docx = mergeDocuments(layout, contents);
       	
        	DocxDocument docx = insertFromPlaceholderBlocks(path + "FlettemotorData.docx");
        	
			docx.writeToFile(path + "dokument.docx");
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (Docx4JException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
    }
    
    private static DocxDocument replaceStyledPlaceholders(String layout) 
    		throws FileNotFoundException, Docx4JException {
    	DocxDocument layoutTemplate = new DocxDocument(layout);
		
		Map<String, String> headerReplacements = new HashMap<String, String>();
		headerReplacements.put("<%HEADER_MIDDLE%>", "Dette er en header!");
		headerReplacements.put("<%HEADER_LEFT%>", "Venstre side i header");
		headerReplacements.put("<%HEADER_RIGHT%>", "Høyre side i header!");		
		layoutTemplate.replaceHeaderPlaceholders(headerReplacements);
		
		layoutTemplate.replaceFooterPlaceholder("<%FOOTER_LEFT%>", "Venstre side i footer.");
		layoutTemplate.replaceFooterPlaceholder("<%FOOTER_MIDDLE%>", "Dette er en footer");
		
		Map<String, String> bodyReplacements = new HashMap<String, String>();
		bodyReplacements.put("<%TITLE%>", "Min Tittel");
		bodyReplacements.put("<%CONTENT_NORMAL%>", "Normal skrift");
		bodyReplacements.put("<%CONTENT_BOLD%>", "Fet skrift");
		bodyReplacements.put("<%CONTENT_ITALIC%>", "Kursiv skrift");
		bodyReplacements.put("<%CONTENT_UNDERLINED%>", "Understreket skrift");
		bodyReplacements.put("<%CONTENT_BOLD_ITALIC_UNDERLINED%>", "Fet, kursiv, understreket skrift");
		
		String placeholder = "<Denne lange setningen kan også brukes som en ‘placeholder’, hvis man måtte ønske det, men det er kanskje ikke like lett å få øye på den.>";
		bodyReplacements.put(placeholder, "Denne setningen ble satt inn fra fletteprogrammet :).");
		
		layoutTemplate.replaceBodyPlaceholders(bodyReplacements);
		
		return layoutTemplate;
    }

	private static DocxDocument insertFromPlaceholderBlocks(String documentWithBLocks)
			throws FileNotFoundException, Docx4JException {
		DocxDocument docx = new DocxDocument();
		
		DocxDocument data = new DocxDocument(documentWithBLocks);
		docx.insertElementsFromContentBlock(data.getDocument(), "<title>", "</title>");
		docx.insertElementsFromContentBlock(data.getDocument(), "<content>", "</content>");
		docx.insertElementsFromContentBlock(data.getDocument(), "<images>", "</images>");
		
		return docx;
	}
    
    private static DocxDocument mergeDocuments(String layoutTemplatePath, String contentTemplateWithContentPath) 
    		throws IOException, Docx4JException {
    	DocxDocument layoutTemplate = new DocxDocument(layoutTemplatePath);
		DocxDocument contentsTemplate = new DocxDocument(contentTemplateWithContentPath);
		
		layoutTemplate.replacePlaceholderWithDocumentContent("<%CONTENT%>", contentsTemplate.getDocument());		
		
		Map<String, String> headerReplacements = new HashMap<String, String>();
		headerReplacements.put("<%HEADER_MIDDLE%>", "Dette er en header!");
		headerReplacements.put("<%HEADER_LEFT%>", "Venstre side i header");
		headerReplacements.put("<%HEADER_RIGHT%>", "Høyre side i header!");
		
		layoutTemplate.replaceHeaderPlaceholders(headerReplacements);
		layoutTemplate.replaceFooterPlaceholder("<%FOOTER_LEFT%>", "Venstre side i footer.");
		layoutTemplate.replaceFooterPlaceholder("<%FOOTER_MIDDLE%>", "Dette er en footer");
		layoutTemplate.replaceBodyPlaceholder("<%TITLE%>", "Min Tittel");
		
		return layoutTemplate;
    }
}