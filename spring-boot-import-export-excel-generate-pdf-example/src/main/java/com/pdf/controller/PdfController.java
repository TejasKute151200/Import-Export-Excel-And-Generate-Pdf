package com.pdf.controller;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

import javax.servlet.http.HttpServletResponse;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.lowagie.text.DocumentException;
import com.pdf.entity.PdfEntity;
import com.pdf.ie.PdfExportor;
import com.pdf.service.PdfService;

@RestController
@RequestMapping("/pdf")
public class PdfController {
	
	@Autowired 
	private PdfService service;
	
	@GetMapping("/TejTechWorld")
	public ResponseEntity<InputStreamResource> createPdf() {
		
		ByteArrayInputStream pdf = service.createPdf();
		
//		DateFormat dateFormat = new SimpleDateFormat("yyyy-mm-dd_HH:mm:ss");
//		String currentDate = dateFormat.format(new Date());
		
		HttpHeaders httpHeaders = new HttpHeaders();
		httpHeaders.add("CONTENT_DISPOSITION", "inline;file=lcwd.pdf");	
		
		return ResponseEntity.ok()
				.headers(httpHeaders)
				.contentType(MediaType.APPLICATION_PDF)
				.body(new InputStreamResource(pdf));
		
	}

         
    @GetMapping("/export")
    public void exportToPDF(HttpServletResponse response) throws DocumentException, IOException {
        response.setContentType("application/pdf");
        DateFormat dateFormatter = new SimpleDateFormat("yyyy-MM-dd_HH:mm:ss");
        String currentDateTime = dateFormatter.format(new Date());
         
        String headerKey = "Content-Disposition";
        String headerValue = "attachment; filename=TejTech_" + currentDateTime + ".pdf";
        response.setHeader(headerKey, headerValue);
         
        List<PdfEntity> listUsers = service.listAll();
         
        PdfExportor exporter = new PdfExportor(listUsers);
        exporter.export(response);
         
    }
	

}