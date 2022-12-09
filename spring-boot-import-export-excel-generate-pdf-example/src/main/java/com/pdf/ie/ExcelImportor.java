package com.pdf.ie;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.pdf.entity.PdfEntity;

public class ExcelImportor {

public List<PdfEntity> excelImport(){
	List<PdfEntity> listOfUsers=new ArrayList<>();
	int id=0;
	String fullName="";
	String email="";
	String location="";
	
	String excelFilePath="C:\\Users\\hp\\Desktop\\PdfEntity.xlsx";
	
	long start = System.currentTimeMillis();
	
	FileInputStream inputStream;
	try {
		inputStream = new FileInputStream(excelFilePath);
		
		Workbook workbook=new XSSFWorkbook(inputStream);
		Sheet firstSheet=workbook.getSheetAt(0);
		Iterator<Row> rowIterator=firstSheet.iterator();
		rowIterator.next();
		
		while(rowIterator.hasNext()) {
			Row nextRow = rowIterator.next();
			Iterator<Cell> cellIterator=nextRow.cellIterator();
			while(cellIterator.hasNext()) {
				Cell nextCell=cellIterator.next();
				int columnIndex=nextCell.getColumnIndex();
				switch (columnIndex) {
				case 0:
					id=(int) nextCell.getNumericCellValue();
					System.out.println(id);
					break;
				case 1:
					fullName=nextCell.getStringCellValue();
					System.out.println(fullName);
					break;
				case 2:
					email=nextCell.getStringCellValue();
					System.out.println(email);
					break;
				case 3:
					location=nextCell.getStringCellValue();
					System.out.println(location);
					break;
				
				}
				listOfUsers.add(new PdfEntity(id, fullName, email, location));
			}
		}
		
		workbook.close();
		long end = System.currentTimeMillis();
		System.out.printf("Import done in %d m\n", (end - start));
		
	} catch (Exception e) {
		
		e.printStackTrace();
	}
	
	return listOfUsers;
}

}