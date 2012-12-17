package no.mesan.forskningsraadet;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

/*
 * Some helper methods found in this example: 
 * http://www.javacodegeeks.com/2012/07/java-word-docx-documents-with-docx4j.html
 */
public class App 
{
	private static final int HEADER = 1;
	private static final int FOOTER = 2;
	
    public static void main( String[] args )
    {
		String path = "/media/sf_DATA_DRIVE/Documents/Jobb/Forskningsraadet/";
		String layout = path + "template.docx";
		String contents = path + "IMMedInnhold.docx";
    	
        try {
			DocxDocument docx = mergeDocuments(layout, contents);
			docx.writeToFile(path + "dokument.docx");
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (Docx4JException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
    }
    
    private static DocxDocument mergeDocuments(String layoutTemplatePath, String contentTemplateWithContentPath) 
    		throws IOException, Docx4JException {
    	DocxDocument layoutTemplate = new DocxDocument(layoutTemplatePath);
		DocxDocument contentsTemplate = new DocxDocument(contentTemplateWithContentPath);
		
		layoutTemplate.replacePlaceholderWithDocumentContent("<%CONTENT%>", contentsTemplate);		
		
		Map<String, String> headerReplacements = new HashMap<String, String>();
		headerReplacements.put("<%HEADER_MIDDLE%>", "Dette er en header!");
		headerReplacements.put("<%HEADER_LEFT%>", "Venstre side i header");
		headerReplacements.put("<%HEADER_RIGHT%>", "Høyre side i header!");
		
		layoutTemplate.replaceHeaderPlaceholders(headerReplacements);
		layoutTemplate.replaceFooterPlaceholder("<%FOOTER%>", "Dette er en footer");
		layoutTemplate.replacePlaceholder("<%TITLE%>", "Min Tittel");
		
		return layoutTemplate;
    }
}