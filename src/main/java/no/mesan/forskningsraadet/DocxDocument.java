package no.mesan.forskningsraadet;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

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
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.Parts;
import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart.AddPartBehaviour;
import org.docx4j.wml.Body;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.RPr;
import org.docx4j.wml.Style;
import org.docx4j.wml.Text;

public class DocxDocument {
	private static final int HEADER = 1;
	private static final int FOOTER = 2;
	private static final String PLACEHOLDER_START= "<";
	private static final String PLACEHOLDER_END = ">";
	
	private String filePath;
	private WordprocessingMLPackage document;
	
	//TODO: Make all the replacePlaceholder methods be able to handle text split into more than one run element.
	//TODO: Research the Office OpenXML structure
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
				
				//Works better on Linux and OS X
				document.setFontMapper(new BestMatchingMapper());
				
				//Turns off logging, to prevent getting log messages in the
				//created PDF document
				Logger log = Logger.getLogger("org/docx4j/convert/out/pdf/viaXSLFO/");
                LoggerRepository repository = log.getLoggerRepository();
                repository.setThreshold(Level.OFF);
				
				PdfConversion conversion = new Conversion(document);
				
				OutputStream os = new FileOutputStream(file);
				
				conversion.output(os, null);
				System.out.println("Saved as PDF");
			} catch (Exception e) {
				e.printStackTrace();
			}
		} else {
			throw new IllegalArgumentException("Currently only .docx and .pdf extensions are supported.");
		}
	}
	
	/**
	 * Method that gets all elements of a specific class from an element,
	 * e.g. all runs in a paragraph or all paragraphs in a document.
	 * 
	 * Shamelessly stolen from an example at: 
	 * http://www.javacodegeeks.com/2012/07/java-word-docx-documents-with-docx4j.html
	 * 
	 * @param obj the object to retrieve elements from
	 * @param toSearch the class to search for
	 */
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
	
	public void replaceBodyPlaceholders(Map<String, String> replacements) {		
		//gets all paragraphs from the body of the document
		List<Object> paragraphs = getAllElementFromObject(
				document.getMainDocumentPart(), P.class);
		
		for (Object obj : paragraphs) {
			P paragraph = (P) obj;
			
			//gets the text content
			String content = getAllTextInParagraph(paragraph);
			
			if (replacements.isEmpty()) break;
			
			//loops through the keys and checks if content contains the placeholder (key)
			for(Iterator<String> iterator = replacements.keySet().iterator(); iterator.hasNext();) {
				String key = iterator.next();
				if (content.contains(key)) {
					String replacementText = replacements.get(key);
					
					replacePlaceholderInParagraph(paragraph, key, replacementText);
					
					iterator.remove();
					break;
				}
			}
		}
	}
	
	public void replaceBodyPlaceholder(String placeholder, String replacementText) {		
		//gets all paragraphs from the body of the document
		List<Object> paragraphs = getAllElementFromObject(
				document.getMainDocumentPart(), P.class);
		
		for (Object obj : paragraphs) {
			P paragraph = (P) obj;
			
			//gets the text content
			String content = getAllTextInParagraph(paragraph);
			
			if (content.contains(placeholder)) {
				replacePlaceholderInParagraph(paragraph, placeholder, replacementText);
			}
		}
	}
	
	private void replaceHeaderFooterPlaceholder(String placeholder, String replacementText, int option) {
		List<SectionWrapper> sectionWrappers = document.getDocumentModel().getSections();
		
		for (SectionWrapper sw : sectionWrappers) {	
			HeaderFooterPolicy hfp = sw.getHeaderFooterPolicy();
			
			Object element = null;
			if (option == HEADER) {
				element = hfp.getDefaultHeader();
			} else {
				element = hfp.getDefaultFooter();
			}

			//String xpath = "//w:r[w:t[contains(text(),'" + placeholder + "')]]";
			String xpath = "//w:p[w:r[w:t]]";
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
			
			//Loops through all the paragraphs in the header/footer
			//not sure if there will ever be more than one, but just in case
			for (Object p: list) {
				P paragraph = (P) p;
				
				replacePlaceholderInParagraph(paragraph, placeholder, replacementText);
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

        //String xpath = "//w:p[w:r[w:t[contains(text(),'" + placeholder + "')]]]";
        List<P> paragraphs = getAllParagraphsContainingPlaceholder(documentPart, placeholder);
		
		for (P paragraph: paragraphs) {
			
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
		
		addAllStylesFromDocument(document);
	}
	
	//TODO: Make it possible to write one liner blocks, e.g. blockStart(.+)blockEnd
	public void insertElementsFromContentBlock(WordprocessingMLPackage document, String blockStart, String blockEnd) {
			
		MainDocumentPart documentPart = document.getMainDocumentPart();
		
		List<Object> list = getAllElementFromObject(documentPart, P.class);
		
		int startIndex = 0;
		int endIndex = 0;
		for(int i = 0; i < list.size(); i++) {
			P paragraph = (P) list.get(i);
			
			String content = getAllTextInParagraph(paragraph);
			
			if (content.equals(blockStart)) {
				startIndex = documentPart.getContent().indexOf(paragraph) + 1;
			} else if (content.equals(blockEnd)) {
				endIndex = documentPart.getContent().indexOf(paragraph);
				
				//Gets all the elements inside the block
				List<Object> elementsToAdd = documentPart.getContent().subList(startIndex, endIndex);
				
				//Adds them to the document	
				this.document.getMainDocumentPart().getContent().addAll(elementsToAdd);
				
				//Copy images, if any
				//TODO: Update relationships between image and image xml
				Parts parts = document.getParts();
				
				Map<PartName, Part> partsMap = parts.getParts();
				Iterator<PartName> iterator = partsMap.keySet().iterator();
				while(iterator.hasNext()) {
					PartName name = iterator.next();
					
					//not sure if more any other extension than png is used, so
					//may be more than the ones listed
					if (name.getExtension().equals("png") || 
						name.getExtension().equals("jpg") ||
						name.getExtension().equals("jpeg")) {
						
						try {
							this.document.addTargetPart(partsMap.get(name), 
											AddPartBehaviour.OVERWRITE_IF_NAME_EXISTS);
						} catch (InvalidFormatException e) {
							e.printStackTrace();
						}
					}
				}
				
			} /*else {
				int start = content.indexOf(blockStart);
				int end = content.indexOf(blockEnd);
				
				//the block is a one liner, e.g. blockStart(.+)blockEnd
				if (start != -1 && end != -1) {
					List<Object> texts = getAllElementFromObject(paragraph, Text.class);
					
					for(Object obj: texts) {
						Text text = (Text) obj;
						String textContent = text.getValue().trim();
						
						if (textContent.equals(blockStart) || textContent.equals(blockEnd)) {
							R run = (R) text.getParent();
							
						}
						
//						if (textContent.equals(content)) {
//							String subStartBlock = textContent.substring(start, start + blockStart.length() - 1);
//							String subEndBlock = textContent.substring(end);
//							
//							textContent.replace(subStartBlock, "");
//							textContent.replace(subEndBlock, "");
//							
//							text.setValue(textContent);
//							break;
//						}
					}
					
					this.document.getMainDocumentPart().getContent().add(paragraph);
				}
			}*/
		}
		
		addAllStylesFromDocument(document);
	}
	
	/**
	 * Copy style definitions from a document into this document, in case we copy elements
	 * from a document which is using different styles than this document, and we want to 
	 * preserve the original styling.
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
	
	/**
	 * Sometimes the text content in a paragraph is split into several
	 * elements. This is a helper method to get all the text content of a
	 * paragraph as a single string.
	 * 
	 * @param a paragraph
	 * @return all the text content of a paragraph as a String
	 */
	private String getAllTextInParagraph(P paragraph) {
		//TODO: Could probably skip getting runs and just get text elements straight from the start
		
		//gets all run elements
		List<Object> runs = getAllElementFromObject(paragraph, R.class);
		
		//loops through the run elements to find all the text elements,
		//in case the text elements have been split up
		StringBuilder contentBuilder = new StringBuilder();
		for(Object obj: runs) {
			R run = (R) obj;
			
			//gets all text elements
			List<Object> texts = getAllElementFromObject(run, Text.class);
			
			//Append the values of each text element
			for(Object text: texts) {
				Text t = (Text) text;
				contentBuilder.append(t.getValue());
			}
		}
		
		return contentBuilder.toString();
	}
	
	/**
	 * Searches a document for all paragraphs who's content contains the placeholder
	 * text.
	 * 
	 * @param document to search
	 * @param placeholder to search for
	 * @return a list of paragraphs which contents are the placeholder text
	 */
	private List<P> getAllParagraphsContainingPlaceholder(MainDocumentPart mainDocument, String placeholder) {
		String xpath = "//w:p[w:r[w:t]]";
		List<Object> list = null;
		try {
			list = mainDocument.getJAXBNodesViaXPath(xpath, false);
		} catch (JAXBException e) {
			e.printStackTrace();
		}
		
		//Finds all the paragraphs that contains the placeholder text
		List<P> paragraphs = new ArrayList<P>();
		for(Object obj: list) {
			P paragraph = (P) obj;
			String content = getAllTextInParagraph(paragraph);
			
			if (content.indexOf(placeholder) != -1) {
				paragraphs.add(paragraph);
			}
		}
		
		return paragraphs;
	}
	
	/**
	 * Method that replaces a placeholder in a paragraph with the replacement text.
	 * This method is able to handle that a placeholder is split into several runs
	 * inside a paragraph.
	 * 
	 * @param paragraph to modify
	 * @param placeholder to replace
	 * @param replacementText to replace with placeholder
	 */
	private void replacePlaceholderInParagraph(P paragraph, String placeholder, String replacementText) {
		//gets all the runs
		List<Object> runs = getAllElementFromObject(paragraph, R.class);
		
		//loops through all the runs in the paragraph to search for
		//the placeholder
		String content = "";
		int startIndex = 0, endIndex = 0;
		boolean append = false;
		R startRun = null;//, endRun = null;
		Set<R> placeholderRuns = new HashSet<R>();
		RPr rpr = null;
		for(Object r: runs) {
			R run = (R) r;
			
			//gets all the texts, not sure if there will ever be more than one
			List<Object> texts = getAllElementFromObject(run, Text.class);
			
			for(Object t: texts) {
				Text text = (Text) t;
				String textContent = text.getValue();
				
				startIndex = textContent.indexOf(PLACEHOLDER_START);						
				endIndex = textContent.indexOf(PLACEHOLDER_END);
				
				//Checks if the text contains the whole placeholder
				if (startIndex != -1 && endIndex != -1) {
					content = textContent.substring(startIndex, endIndex + PLACEHOLDER_END.length());
					placeholderRuns.add(run);
					
					if (content.trim().equals(placeholder)) {
						text.setValue(replacementText);
						
						return;
					}
				} 
				//only start of the placeholder
				else if (startIndex != -1 && endIndex == -1) {
					append = true;
					String sub = textContent.substring(startIndex);
					content += sub;
					rpr = run.getRPr();
					startRun = run;
					placeholderRuns.add(run);
				}
				//only the end of the placeholder
				else if (startIndex == -1 && endIndex != -1) {
					append = false;
					String sub = textContent.substring(0, endIndex + PLACEHOLDER_END.length());
					content += sub;
					//endRun = run;
					placeholderRuns.add(run);
				}
				//if both is -1, then we append the text if append is true
				else if (startIndex == -1 && endIndex == -1 && append) {
					content += textContent;
					placeholderRuns.add(run);
				}
			}
			
			if (content.trim().equals(placeholder)) {
				org.docx4j.wml.ObjectFactory factory = new org.docx4j.wml.ObjectFactory();
				
				//TODO: Reuse (replace text in) the startRun instead of creating a new one?
				
				//Creates a new run and adds the original run's properties
				R newRun = factory.createR();
				newRun.setRPr(rpr);
				
				//Creates and add a new text element
				Text newText = factory.createText();
				newText.setValue(replacementText);
				newRun.getContent().add(newText);
				
				//gets the position to insert the new run
				int placementIndex = paragraph.getContent().indexOf(startRun);
				
				//removes the split runs that contained the placeholder
				paragraph.getContent().removeAll(placeholderRuns);
				
				//adds the new run
				paragraph.getContent().add(placementIndex, newRun);
				//content = "";
				
				return;
			}
		}
	}
}
