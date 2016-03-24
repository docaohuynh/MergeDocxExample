package manager;

import java.io.File;
import java.io.FileOutputStream;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;
import java.util.Map.Entry;

import org.docx4j.XmlUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.io.SaveToZipFile;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.ObjectFactory;

public class DocManager {
	private WordprocessingMLPackage wordMLPackage;
	private static DocManager instance;

	public static DocManager getInstance() {
		if (instance == null)
			instance = new DocManager();
		return instance;
	}

	public DocManager() {

	}

	public void readFile(String pathInputFile) {
		try {
			wordMLPackage = WordprocessingMLPackage.load(new java.io.File(pathInputFile));
		} catch (Docx4JException e) {
			e.printStackTrace();
		}
	}

	public void writeFile(List<Object> contents, String pathOutputFile) {
		try {
			// Create the package
			WordprocessingMLPackage wordMLPackage = new WordprocessingMLPackage();

			// Create the main document part (word/document.xml)
			MainDocumentPart wordDocumentPart = new MainDocumentPart();

			// Create main document part content
			ObjectFactory factory = Context.getWmlObjectFactory();
			org.docx4j.wml.Body body = factory.createBody();
			org.docx4j.wml.Document wmlDocumentEl = factory.createDocument();
			wmlDocumentEl.setBody(body);

			// Put the content in the part
			for (Object content : contents)
				wordDocumentPart.addObject(content);

			// Add the main document part to the package relationships
			wordMLPackage.addTargetPart(wordDocumentPart);
			wordMLPackage.save(new File(pathOutputFile));
		} catch (Docx4JException e) {
			e.printStackTrace();
		}
	}

	public void removeContentFile(WordprocessingMLPackage wordMLPackage, List<Integer> positionsRemove,
			String pathOutputFile) {
		for (Integer positionRemove : positionsRemove)
			wordMLPackage.getMainDocumentPart().getContent().remove(positionRemove);
		try {
			wordMLPackage.save(new File(pathOutputFile));
		} catch (Docx4JException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public void duplicateContent(WordprocessingMLPackage wordMLPackage, String pathOutputFile) {
		try {
			wordMLPackage.addTargetPart(wordMLPackage.getMainDocumentPart());
			wordMLPackage.addTargetPart(wordMLPackage.getMainDocumentPart());
			wordMLPackage.addTargetPart(wordMLPackage.getMainDocumentPart());
			wordMLPackage.save(new File(pathOutputFile));
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	public void mergeFile(String inputFile, String outputfilepath) {
		boolean save = true;
		String dataPath = "";
		ArrayList<String> sourceDocxNames = new ArrayList<>();
		sourceDocxNames.add(inputFile);
		sourceDocxNames.add(inputFile);
		
		// Create list of docx packages to merge
		List<WordprocessingMLPackage> wmlPkgList = new ArrayList<WordprocessingMLPackage>();
		try {
			for (int i = 0; i < sourceDocxNames.size(); i++) {
				String filename = dataPath + sourceDocxNames.get(i);
				System.out.println("Loading " + filename);
				wmlPkgList.add(WordprocessingMLPackage.load(new java.io.File(filename)));
			}

			// Use reflection, so docx4j can be built
			// by users who don't have the MergeDocx utility
			Class<?> documentBuilder = Class.forName("com.plutext.merge.DocumentBuilder");
			// Method method = documentBuilder.getMethod("merge",
			// wmlPkgList.getClass());
			Method[] methods = documentBuilder.getMethods();
			Method method = null;
			for (int j = 0; j < methods.length; j++) {
				System.out.println(methods[j].getName());
				if (methods[j].getName().equals("merge")) {
					method = methods[j];
					break;
				}
			}
			if (method == null)
				throw new NoSuchMethodException();

			WordprocessingMLPackage resultPkg = (WordprocessingMLPackage) method.invoke(null, wmlPkgList);

			if (save) {
				SaveToZipFile saver = new SaveToZipFile(resultPkg);
				saver.save(outputfilepath);
				System.out.println("Generated " + outputfilepath);
			} else {
				String result = XmlUtils.marshaltoString(resultPkg.getMainDocumentPart().getJaxbElement(), true, true);
				System.out.println(result);
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public ArrayList<BinaryPartAbstractImage> getPartImage() {
		ArrayList<BinaryPartAbstractImage> images = new ArrayList<>();
		for (Entry<PartName, Part> entry : wordMLPackage.getParts().getParts().entrySet()) {
			if (entry.getValue() instanceof BinaryPartAbstractImage) {
				images.add((BinaryPartAbstractImage) entry.getValue());
			}
		}
		return images;
	}

	public WordprocessingMLPackage getWordMLPackage() {
		return wordMLPackage;
	}
}
