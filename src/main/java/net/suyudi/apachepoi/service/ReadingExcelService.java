package net.suyudi.apachepoi.service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

/**
 * ReadingExcelService
 */
@Service
public class ReadingExcelService {

    public String GetData(String filePath) {

        System.out.println(filePath);

        try {

            FileInputStream excelFile = new FileInputStream(new File(filePath));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();

            iterator.next();

            while (iterator.hasNext()) {

                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();

                while (cellIterator.hasNext()) {

                    Cell currentCell = cellIterator.next();

                    switch (currentCell.getCellType()) {
                        case NUMERIC:
                            System.out.print(currentCell.getNumericCellValue() + " | ");
                            break;

                        case BOOLEAN:
                            System.out.print(currentCell.getBooleanCellValue() + " | ");
                            break;

                        case FORMULA:
                            System.out.print(currentCell.getCellFormula() + " | ");
                            break;
                    
                        default:
                            System.out.print(currentCell.getStringCellValue() + " | ");
                            break;
                    }

                }

                System.out.println(" ");

            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        
        return "extract file done!";        
    }
}