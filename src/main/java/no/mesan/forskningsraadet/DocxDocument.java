package no.mesan.forskningsraadet;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;

import org.apache.log4j.Level;
import org.apache.log4j.Logger;
import org.apache.log4j.spi.LoggerRepository;
import org.docx4j.convert.out.pdf.PdfConversion;
import org.docx4j.convert.out.pdf.viaXSLFO.Conversion;
import org.docx4j.fonts.BestMatchingMapper;
import org.docx4j.model.structure.HeaderFooterPolicy;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Body;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Text;

public class DocxDocument {
	private static final int HEADER = 1;
	private static final int FOOTER = 2;	
	
	private String filePath;
	private WordprocessingMLPackage document;
	
	public DocxDocument() {
		try {
			document = WordprocessingMLPackage.createPackage();
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		}
	}
	
	public DocxDocument(String filePath) throws FileNotFoundException, Docx4JException {
		document = WordprocessingMLPackage.load(new FileInputStream(filePath));
	}

	public String getFilePath() {
		return filePath;
	}

	public void setFilePath(String filePath) {
		this.filePath = filePath;
	}
	
	public WordprocessingMLPackage getDocument() {
		return document;
	}
	
	public void setDocument(WordprocessingMLPackage document) {
		this.document = document;
	}
	
	public void writeToFile(String filePath) throws IOException, Docx4JException {
		File file = new File(filePath);
		
		if (filePath.endsWith(".docx")) {
			document.save(file);			
		} else if (filePath.endsWith(".pdf")) {
			try {
				//Works better on windows
				//document.setFontMapper(new IdentityPlusMapper());
				
				//Works better on linux and OS X
				document.setFontMapper(new BestMatchingMapper());
				
				//Turns off logging, so that we don't get log messages in the
				//created pdf document
				Logger log = Logger.getLogger("org/docx4j/convert/out/pdf/viaXSLFO/");
                LoggerRepository repository = log.getLoggerRepository();
                repository.setThreshold(Level.OFF);
				
				PdfConversion conversion = new Conversion(document);
				
				OutputStream os = new FileOutputStream(file);
				
				conversion.output(os, null);
			} catch (Exception e) {
				e.printStackTrace();
			}
		} else {
			throw new IllegalArgumentException("Currently only .docx and .pdf extensions are supported.");
		}
	}
	
	private List<Object> getAllElementFromObject(Object obj, Class<?> toSearch) {
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
	
	public void replacePlaceholders(Map<String, String> replacements) {		
		List<Object> texts = getAllElementFromObject(
				document.getMainDocumentPart(), Text.class);
		
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
	
	public void replacePlaceholder(String placeholder, String replacementText) {		
		List<Object> texts = getAllElementFromObject(
				document.getMainDocumentPart(), Text.class);
		
		for (Object text : texts) {
			Text textElement = (Text) text;
			if (textElement.getValue().equals(placeholder)) {
				textElement.setValue(replacementText);
			}
		}
	}
	
	private void replaceHeaderFooterPlaceholder(String placeholder, String replacementText, int option) {
		List<SectionWrapper> sectionWrappers = document.getDocumentModel().getSections();
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
	
	public void replaceHeaderPlaceholder(String placeholder, String replacementText) {
		replaceHeaderFooterPlaceholder(placeholder, replacementText, HEADER);
	}
	
	public void replaceFooterPlaceholder(String placeholder, String replacementText) {
		replaceHeaderFooterPlaceholder(placeholder, replacementText, FOOTER);
	}
	
	public void replaceHeaderPlaceholders(Map<String, String> replacements) {
		for(String key: replacements.keySet()) {
			replaceHeaderFooterPlaceholder(key, replacements.get(key), HEADER);
		}
	}
	
	public void replaceFooterPlaceholders(Map<String, String> replacements) {
		for(String key: replacements.keySet()) {
			replaceHeaderFooterPlaceholder(key, replacements.get(key), FOOTER);
		}
	}
	
	public void replacePlaceholderWithDocumentContent(String placeholder, WordprocessingMLPackage document) {
		
		MainDocumentPart documentPart = this.document.getMainDocumentPart();

        String xpath = "//w:p[w:r[w:t[contains(text(),'" + placeholder + "')]]]";
        List<Object> list = null;
		try {
			list = documentPart.getJAXBNodesViaXPath(xpath, false);
		} catch (JAXBException e) {
			e.printStackTrace();
		}
		
		for (Object element : list) {
			P paragraph = (P) element;
			
			//Gets the body and the paragraph's position in the content list
			Body body = (Body) paragraph.getParent();
			int paragraphIndex = body.getContent().indexOf(paragraph);
			
			//Removes the paragraph
			body.getContent().remove(paragraphIndex);
			
			//Inserts content from document
			List<Object> documentElements = document.getMainDocumentPart().getContent();
			for(int i = 0; i < documentElements.size(); i++) {
				body.getContent().add(paragraphIndex + i, documentElements.get(i));
			}
		}
	}
	
	public List<Object> getElementsInsidePlaceholderBlock(String blockStart, String blockEnd) {
		List<Object> result = new ArrayList<Object>();
		
		MainDocumentPart documentPart = document.getMainDocumentPart();
		
		//find all paragraphs
		String xpath = "//w:p[w:r[w:t]]";
		List<Object> list = null;
		try {
			list = documentPart.getJAXBNodesViaXPath(xpath, false);
		} catch (JAXBException e) {
			e.printStackTrace();
		}
		
		boolean shouldAdd = false;
		for(int i = 0; i < list.size(); i++) {
			P paragraph = (P) list.get(i);
			
			//gets all run elemenets
			List<Object> runs = getAllElementFromObject(paragraph, R.class);
			
			//loops through the run elements to find all the text elements,
			//in case the text elements have been split up
			String content = "";
			for(Object obj: runs) {
				R run = (R) obj;
				
				//gets all text elements
				List<Object> texts = getAllElementFromObject(run, Text.class);
				
				for(Object text: texts) {
					Text t = (Text) text;
					content += t.getValue();
				}
			}
			System.out.println(content);
			
			if (shouldAdd && !content.equals(blockEnd)) {
				result.add(paragraph);
			}
			
			if (content.equals(blockStart)) {
				shouldAdd = true;
			} else if (content.equals(blockEnd)) {
				shouldAdd = false;
				
				return result;
			}			
		}
		
		return result;
	}
}
