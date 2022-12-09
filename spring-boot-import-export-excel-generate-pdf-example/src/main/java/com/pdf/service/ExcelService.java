package com.pdf.service;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Objects;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import com.pdf.entity.PdfEntity;
import com.pdf.repo.PdfEntityRepo;

@Service
public class ExcelService {
	
	@Autowired
	private PdfEntityRepo repo;

    public void saveUsersToDatabase(MultipartFile file){
        if(ExcelService.isValidExcelFile(file)){
            try {
                List<PdfEntity> users = ExcelService.getUsersDataFromExcel(file.getInputStream());
                this.repo.saveAll(users);
            } catch (IOException e) {
                throw new IllegalArgumentException("The file is not a valid excel file");
            }
        }
    }

    public List<PdfEntity> getUsers(){
        return repo.findAll();
    }
	
    public static boolean isValidExcelFile(MultipartFile file){
        return Objects.equals(file.getContentType(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" );
    }
   public static List<PdfEntity> getUsersDataFromExcel(InputStream inputStream){
        List<PdfEntity> users = new ArrayList<>();
       try {
           XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
           XSSFSheet sheet = workbook.getSheet("C:\\Users\\hp\\Desktop\\PdfEntity.xlsx");
           int rowIndex =0;
           for (Row row : sheet){
               if (rowIndex ==0){
                   rowIndex++;
                   continue;
               }
               Iterator<Cell> cellIterator = row.iterator();
               int cellIndex = 0;
               PdfEntity user = new PdfEntity();
               while (cellIterator.hasNext()){
                   Cell cell = cellIterator.next();
                   switch (cellIndex){
                       case 0 -> user.setId((int) cell.getNumericCellValue());
                       case 1 -> user.setFullName(cell.getStringCellValue());
                       case 2 -> user.setEmail(cell.getStringCellValue());
                       case 3 -> user.setLocation(cell.getStringCellValue());
                       default -> {
                       }
                   }
                   cellIndex++;
               }
               users.add(user);
           }
       } catch (IOException e) {
           e.getStackTrace();
       }
       return users;
   }

}
