package com.pdf.service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;

import javax.transaction.Transactional;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.data.domain.Sort;
import org.springframework.stereotype.Service;
import com.lowagie.text.Chunk;
import com.lowagie.text.Document;
import com.lowagie.text.Element;
import com.lowagie.text.Font;
import com.lowagie.text.FontFactory;
import com.lowagie.text.Paragraph;
import com.lowagie.text.pdf.PdfWriter;
import com.pdf.entity.PdfEntity;
import com.pdf.repo.PdfEntityRepo;

@Service
@Transactional
public class PdfService {
	
    
   @Autowired
   private PdfEntityRepo repo;
    
   public List<PdfEntity> listAll() {
       return repo.findAll(Sort.by("id").ascending());
   }
    

	
	private Logger logger = LoggerFactory.getLogger(PdfService.class);
	

	public ByteArrayInputStream createPdf() {
		
		logger.info("Create PDF Started: ");
		
		String title = "Welcome to TejTechWorld\n";
		String content = "\n We Provide Social Promotion In All Over India \n Share Our Contacts to Your Friends Or Relatives \n";
		
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
		
		Document document = new Document();
		
		PdfWriter.getInstance(document, outputStream);
		
		document.open();
		
		
		Font titleFont = FontFactory.getFont(FontFactory.TIMES_BOLD,24);
		Paragraph titlePara = new Paragraph(title,titleFont);	
		titlePara.setAlignment(Element.ALIGN_CENTER);
		document.add(titlePara);
		
		Font paraFont = FontFactory.getFont(FontFactory.COURIER, 12);
		 Paragraph paragraph = new Paragraph(content,paraFont);
		 paragraph.add(new Chunk("\n Chunk is the method \n We Can Add Data After Creation Paragraph with the Help Of Chunk \n"));
		 paragraph.setAlignment(Element.ALIGN_CENTER);
		 document.add(paragraph);
		 
		 document.close();
		 
		 return new ByteArrayInputStream(outputStream.toByteArray());
}
	
//	public IoStream

}
