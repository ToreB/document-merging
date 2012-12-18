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
import org.docx4j.XmlUtils;
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
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.wml.Body;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.PPrBase.PStyle;
import org.docx4j.wml.R;
import org.docx4j.wml.RPr;
import org.docx4j.wml.Style;
import org.docx4j.wml.Styles;
import org.docx4j.wml.Text;

public class DocxDocument {
	private static final int HEADER = 1;
	private static final int FOOTER = 2;	
	
	private String filePath;
	private WordprocessingMLPackage document;
	
	public DocxDocument() {
		/*String filePath = "newWordDocumentTemplate.docx";
		document = WordprocessingMLPackage.load(new FileInputStream(filePath));*/
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
				
				//Turns off logging, to prevent getting log messages in the
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
				R run = (R) list.get(i);
				P parent = (P) run.getParent();
				RPr rpr = run.getRPr();
				int index = parent.getContent().indexOf(run);
				parent.getContent().remove(run);

				Text addedTmpText = factory.createText();
				R mainR = factory.createR();
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
	
	public void insertElementsFromContentBlock(WordprocessingMLPackage document, String blockStart, String blockEnd) {
		
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
		int startIndex = 0;
		int endIndex = 0;
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
				
				//Append the values of each text element
				for(Object text: texts) {
					Text t = (Text) text;
					content += t.getValue();
				}
			}
			
			if (shouldAdd && !content.equals(blockEnd)) {
				//Adds a deep copy of the paragraph, to preserve the styling
				P deepCopy = XmlUtils.deepCopy(paragraph);
				
				this.document.getMainDocumentPart().getContent().add(deepCopy);
			}
			
			if (content.equals(blockStart)) {
				shouldAdd = true;
				startIndex = documentPart.getContent().indexOf(paragraph);
			} else if (content.equals(blockEnd)) {
				shouldAdd = false;
				endIndex = documentPart.getContent().indexOf(paragraph);
				
				
				
				//Uncomment this line to prevent finding more than one matching block
				//break;
			}			
		}
		
		//Adds styles
		addAllStylesFromDocument(document);
	}
	
	/**
	 * Copy style definitions from a document into this document, in case we copy elements
	 * from a document and want to preserve their original styling.
	 * 
	 * @param document to copy styles from
	 */
	private void addAllStylesFromDocument(WordprocessingMLPackage document) {
		List<Style> stylesToCopy = 
			document.getMainDocumentPart().getStyleDefinitionsPart().getJaxbElement().getStyle();
		List<Style> currentDocumentStyles = 
			this.document.getMainDocumentPart().getStyleDefinitionsPart().getJaxbElement().getStyle();
		
		for(Style style: stylesToCopy) {
			if (!currentDocumentStyles.contains(style)) {
				currentDocumentStyles.add(style);
			}
		}
	}
}
