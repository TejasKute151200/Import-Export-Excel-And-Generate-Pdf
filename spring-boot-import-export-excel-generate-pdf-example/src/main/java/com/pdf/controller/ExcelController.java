package com.pdf.controller;

import java.io.IOException;
import java.util.List;

import javax.servlet.http.HttpServletResponse;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RestController;

import com.pdf.entity.PdfEntity;
import com.pdf.ie.ExcelExportor;
import com.pdf.ie.ExcelImportor;
import com.pdf.repo.PdfEntityRepo;

@RestController
public class ExcelController {
	
//	@Autowired
//	private ExcelService service;
    
	@Autowired
	private PdfEntityRepo repo;

	@GetMapping("/export/excel")
	public String exportToExcel(HttpServletResponse response) throws IOException {
		response.setContentType("application/octet-stream");
		String headerKey = "Content-Disposition";
		String headervalue = "attachment; filename=PdfEntity.xlsx";

		response.setHeader(headerKey, headervalue);
		List<PdfEntity> listOfUsers = repo.findAll();
		ExcelExportor exp = new ExcelExportor(listOfUsers);
		exp.export(response);
		return "Export Successfully";
	}

	@PostMapping("/import/excel")
	public String importFromExcel() {
		ExcelImportor excelImporter=new ExcelImportor();
		List<PdfEntity> listOfUsers= excelImporter.excelImport();
		repo.saveAll(listOfUsers);
		return "Import Successfully";
	}

//    @PostMapping("/import")
//    public ResponseEntity<?> sendUsersData(@RequestParam("file")MultipartFile file){
//        this.service.saveUsersToDatabase(file);
//        return ResponseEntity.ok(
//        		Map.of("Message" , " Users data uploaded and saved to database successfully"));
//    }
//
//    @GetMapping("/export")
//    public ResponseEntity<List<PdfEntity>> getUsers(){
//        return new ResponseEntity<>(service.getUsers(), HttpStatus.FOUND);
//    }

}
