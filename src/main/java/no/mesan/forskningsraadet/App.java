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
	private static final int HEADER = 1;
	private static final int FOOTER = 2;
	
    public static void main( String[] args )
    {
		String path = "/media/sf_DATA_DRIVE/Documents/Jobb/Forskningsraadet/";
		String layout = path + "template.docx";
		String contents = path + "IMMedInnhold.docx";
    	
        try {
			DocxDocument layoutTemplate = new DocxDocument(layout);
			DocxDocument contentsTemplate = new DocxDocument(contents);
			
			layoutTemplate.replacePlaceholderWithDocumentContent("<%CONTENT%>", contentsTemplate);
			
			/*
			Map<String, String> headerReplacements = new HashMap<String, String>();
			headerReplacements.put("<%HEADER_MIDDLE%>", "Dette er en header!");
			headerReplacements.put("<%HEADER_LEFT%>", "Venstre side i header");
			headerReplacements.put("<%HEADER_RIGHT%>", "Høyre side i header!");
			
			layoutTemplate.replaceHeaderPlaceholders(headerReplacements);
			layoutTemplate.replaceFooterPlaceholder("<%FOOTER%>", "Dette er en footer");
			layoutTemplate.replacePlaceholder("<%TITLE%>", "Min Tittel");
			
			
			layoutTemplate.writeToFile(path + "dokument.docx");*/
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (Docx4JException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
    }
}