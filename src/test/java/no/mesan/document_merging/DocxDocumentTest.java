package no.mesan.document_merging;

import static org.junit.Assert.*;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;

import no.mesan.document_merging.DocxDocument;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.junit.Before;
import org.junit.Test;
import org.junit.experimental.categories.Categories.ExcludeCategory;

public class DocxDocumentTest {
	private DocxDocument doc;
	private String path = "./eksempler/replaceStyledPlaceholdersTemplate.docx";
	
	@Test
	public void testDefaultCtor_createsdocument() {		
		doc = new DocxDocument();
		
		assertNotNull(doc.getDocument());
	}

	@Test
	public void testParameterizedCtor_WhenGivenAValidFilePath_ShouldNotThrowException() {
		try {
			doc = new DocxDocument(path);			
		} catch(Exception e) {
			fail();
		}
	}
	
	@Test
	public void testParameterizedCtor_WhenGivenAnInvalidFilePath_ShouldThrowException() {
		try {
			doc = new DocxDocument("invalid/file/path.docx");
			
			fail();
		} catch(Exception e) {	}
	}

	@Test
	public void testWriteToFile_ShouldNotThrowException() {
		writeToFile("testDocument.docx");
	}
	
	@Test
	public void testWriteToPDF_ShouldNotThrowException() {
		writeToFile("testDocument.pdf");
	}
	
	private void writeToFile(String filename) {	
		try {
			doc = new DocxDocument(path);
			doc.writeToFile(filename);
		} catch (Exception e) {
			fail();
		}
		
		File file = new File(filename);
		if (!file.exists()) {
			fail("File should exist");
		} else { //File exists, so clean up
			try {
				file.delete();
			} catch(SecurityException e) {
				//Don't really care, file will get overwritten anyway
			}
		}
	}
}
