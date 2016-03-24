package manager;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.util.IOUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.contenttype.ContentType;
import org.docx4j.openpackaging.io.SaveToZipFile;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.WordprocessingML.AlternativeFormatInputPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.CTAltChunk;

public class MergeDocx {
	private static long chunk = 0;
	private static final String CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

	public void mergeDocx(InputStream s1, InputStream s2, OutputStream os) throws Exception {
		WordprocessingMLPackage target = WordprocessingMLPackage.load(s1);
		insertDocx(target.getMainDocumentPart(), IOUtils.toByteArray(s2));
		SaveToZipFile saver = new SaveToZipFile(target);
		saver.save(os);
	}
	
	public void mergeListDocx(List<InputStream> s1, OutputStream os) throws Exception {
		WordprocessingMLPackage target = WordprocessingMLPackage.load(s1.get(0));
		insertListDocx(target.getMainDocumentPart(), s1);
		SaveToZipFile saver = new SaveToZipFile(target);
		saver.save(os);
	}

	private static void insertDocx(MainDocumentPart main, byte[] bytes) throws Exception {
		AlternativeFormatInputPart afiPart = new AlternativeFormatInputPart(
				new PartName("/part" + (chunk++) + ".docx"));
		afiPart.setContentType(new ContentType(CONTENT_TYPE));
		afiPart.setBinaryData(bytes);
		Relationship altChunkRel = main.addTargetPart(afiPart);

		CTAltChunk chunk = Context.getWmlObjectFactory().createCTAltChunk();
		chunk.setId(altChunkRel.getId());
		
		CTAltChunk chunk1 = Context.getWmlObjectFactory().createCTAltChunk();
		chunk1.setId(altChunkRel.getId());
		
		CTAltChunk chunk2 = Context.getWmlObjectFactory().createCTAltChunk();
		chunk2.setId(altChunkRel.getId());

		main.addObject(chunk);
		main.addObject(chunk1);
		main.addObject(chunk2);
	}
	
	private static void insertListDocx(MainDocumentPart main, List<InputStream> listDocx) throws Exception {
		int i=0;
		for(InputStream input : listDocx){
			if(i != 0 ){
				AlternativeFormatInputPart afiPart = new AlternativeFormatInputPart(
						new PartName("/part" + (chunk++) + ".docx"));
				System.out.println("count "+ i);
				byte[] bytes = IOUtils.toByteArray(input);
				afiPart.setContentType(new ContentType(CONTENT_TYPE));
				afiPart.setBinaryData(bytes);
				Relationship altChunkRel = main.addTargetPart(afiPart);
				CTAltChunk chunk = Context.getWmlObjectFactory().createCTAltChunk();
				chunk.setId(altChunkRel.getId());
				main.addObject(chunk);
			}
			i++;
		}
	}
}