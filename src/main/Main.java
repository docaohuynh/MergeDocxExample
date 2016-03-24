package main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

import manager.DocManager;
import manager.MergeDocx;

public class Main {
	static String pathInput1File = "C:\\Users\\docaohuynh\\Google Drive\\MyProject\\Java\\Process\\doc1.docx";
	static String pathInput2File = "C:\\Users\\docaohuynh\\Google Drive\\MyProject\\Java\\Process\\doc2.docx";
	static String pathInput3File = "C:\\Users\\docaohuynh\\Google Drive\\MyProject\\Java\\Process\\doc3.docx";
	static String pathOutputFile = "C:\\Users\\docaohuynh\\Google Drive\\MyProject\\Java\\Process\\docfull.docx";

	public static void main(String[] args) {
		DocManager.getInstance().readFile(pathInput1File);
		MainDocumentPart documentPart = DocManager.getInstance().getWordMLPackage().getMainDocumentPart();
		List<Object> contents = documentPart.getContent();

		List<Object> newContents = new ArrayList<>();
		newContents.add(contents.get(0));

		List<String> testString = null;

		DocManager.getInstance().writeFile(newContents, pathOutputFile);

		DocManager.getInstance().duplicateContent(DocManager.getInstance().getWordMLPackage(), pathOutputFile);
		DocManager.getInstance().mergeFile(pathInput1File, pathOutputFile);

		try {
			InputStream inputStream1 = new FileInputStream(new File(pathInput1File));
			InputStream inputStream2 = new FileInputStream(new File(pathInput2File));
			InputStream inputStream3 = new FileInputStream(new File(pathInput3File));
			OutputStream outStream = new FileOutputStream(new File(pathOutputFile));
			MergeDocx mergeDocx = new MergeDocx();
			mergeDocx.mergeDocx(inputStream1, inputStream2, outStream);
		} catch (Exception e) {
			e.printStackTrace();
		}
		// List<Integer> positionsRemove = new ArrayList<>();
		// positionsRemove.add(1);
		// positionsRemove.add(2);
		// DocManager.getInstance().removeContentFile(DocManager.getInstance().getWordMLPackage(),positionsRemove,pathOutputFile);
	}
}
