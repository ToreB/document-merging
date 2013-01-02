package no.mesan.document_merging;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.docx4j.openpackaging.exceptions.Docx4JException;

public class App {
    public static void main( String[] args ) {
		String path = "./eksempler/";
		String layout = path + "template.docx";
		String contents = path + "IMMedInnhold.docx";
    	
        try {
        	DocxDocument docx1 = replaceStyledPlaceholders(path + "replaceStyledPlaceholdersTemplate.docx");
        	docx1.writeToFile(path + "replaceStyledPlaceholdersOutput.docx");
        	docx1.writeToFile(path + "replaceStyledPlaceholdersOutput.pdf");       	
        	
			DocxDocument docx2 = 
					mergeDocuments(path + "mergeDocumentTemplate.docx", path + "mergeDocumentContent.docx");
			docx2.writeToFile(path + "mergeDocumentOutput.docx");
			docx2.writeToFile(path + "mergeDocumentOutput.pdf");
			
        	DocxDocument docx3 = insertFromContentBlocks(path + "insertFromContentBlocksData.docx");
        	docx3.writeToFile(path + "insertFromContentBlocksOutput.docx");
        	docx3.writeToFile(path + "insertFromContentBlocksOutput.pdf");
			
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
    	
    	// Loads a document
    	DocxDocument layoutTemplate = new DocxDocument(layout);
		
    	// Creates some replacement values for the placeholders in the header
		Map<String, String> headerReplacements = new HashMap<String, String>();
		headerReplacements.put("<%HEADER_MIDDLE%>", "Dette er en header!");
		headerReplacements.put("<%HEADER_LEFT%>", "Venstre side i header");
		headerReplacements.put("<%HEADER_RIGHT%>", "Høyre side i header!");		
		
		// Replaces the header placeholders with the specified values
		layoutTemplate.replaceHeaderPlaceholders(headerReplacements);
		
		// Replaces the footer placeholders
		layoutTemplate.replaceFooterPlaceholder("<%FOOTER_LEFT%>", "Venstre side i footer.");
		layoutTemplate.replaceFooterPlaceholder("<%FOOTER_MIDDLE%>", "Dette er en footer");
		
		// Creates some replacements for the placeholders in the body
		Map<String, String> bodyReplacements = new HashMap<String, String>();
		bodyReplacements.put("<%TITLE%>", "Min Tittel");
		bodyReplacements.put("<%CONTENT_NORMAL%>", "Normal skrift");
		bodyReplacements.put("<%CONTENT_BOLD%>", "Fet skrift");
		bodyReplacements.put("<%CONTENT_ITALIC%>", "Kursiv skrift");
		bodyReplacements.put("<%CONTENT_UNDERLINED%>", "Understreket skrift");
		bodyReplacements.put("<%CONTENT_BOLD_ITALIC_UNDERLINED%>", "Fet, kursiv, understreket skrift");
		
		// Replaces the body placeholders
		layoutTemplate.replaceBodyPlaceholders(bodyReplacements);
		
		return layoutTemplate;
    }

	private static DocxDocument insertFromContentBlocks(String documentWithBLocks)
			throws FileNotFoundException, Docx4JException {
		
		// Creates a new document
		DocxDocument docx = new DocxDocument();
		
		// Loads the document that contains the content blocks 
		DocxDocument data = new DocxDocument(documentWithBLocks);
		
		// Inserts some content
		docx.insertElementsFromContentBlock(data.getDocument(), "<title>", "</title>");
		docx.insertElementsFromContentBlock(data.getDocument(), "<content>", "</content>");
		docx.insertElementsFromContentBlock(data.getDocument(), "<images>", "</images>");
		
		return docx;
	}
    
    private static DocxDocument mergeDocuments(String layoutTemplatePath, String contentTemplateWithContentPath) 
    		throws IOException, Docx4JException {
    	
    	// Loads some documents
    	DocxDocument layoutTemplate = new DocxDocument(layoutTemplatePath);
		DocxDocument contentsTemplate = new DocxDocument(contentTemplateWithContentPath);
		
		// Replaces the placeholder in one document with the entire body of the other document
		layoutTemplate.replacePlaceholderWithDocumentContent("<%CONTENT%>", contentsTemplate.getDocument());		
		
		// Creates some replacement values
		Map<String, String> headerReplacements = new HashMap<String, String>();
		headerReplacements.put("<%HEADER_MIDDLE%>", "Dette er en header!");
		headerReplacements.put("<%HEADER_LEFT%>", "Venstre side i header");
		headerReplacements.put("<%HEADER_RIGHT%>", "Høyre side i header!");
		
		// Replaces some placeholders
		layoutTemplate.replaceHeaderPlaceholders(headerReplacements);
		layoutTemplate.replaceFooterPlaceholder("<%FOOTER_LEFT%>", "Venstre side i footer.");
		layoutTemplate.replaceFooterPlaceholder("<%FOOTER_MIDDLE%>", "Dette er en footer");
		layoutTemplate.replaceBodyPlaceholder("<%TITLE%>", "Min Tittel");
		
		return layoutTemplate;
    }
}