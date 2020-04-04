package net.suyudi.apachepoi.service;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

@Service
public class WritingExcelService {

    @Value("${local.path.file}")
    private String localPath;

    public String buildFile(String file) {
        System.out.println("building file");

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Author");
        Object[][] datatypes = {
                {"fuad 1", "Suyudi", "fuad1", "fuad1"},
                {"fuad 2", "Suyudi", "fuad2", "fuad2"},
                {"fuad 3", "Suyudi", "fuad3", "fuad3"},
                {"fuad 4", "Suyudi", "fuad4", "fuad4"},
                {"fuad 5", "Suyudi", "fuad5", "fuad5"},
                {"fuad 6", "Suyudi", "fuad6", "fuad6"}
        };

        sheet.setColumnWidth(0, 1000);
        sheet.setColumnWidth(1, 4000);
        sheet.setColumnWidth(2, 4000);
        sheet.setColumnWidth(3, 4000);
        sheet.setColumnWidth(4, 4000);

        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        
        XSSFFont font = ((XSSFWorkbook) workbook).createFont();
        font.setBold(true);
        headerStyle.setFont(font);

        Row row = sheet.createRow(0);

        Cell cell = row.createCell(0);
        cell.setCellValue((String) "No");
        cell.setCellStyle(headerStyle);

        cell = row.createCell(1);
        cell.setCellValue((String) "First Name");
        cell.setCellStyle(headerStyle);

        cell = row.createCell(2);
        cell.setCellValue((String) "Last Name");
        cell.setCellStyle(headerStyle);

        cell = row.createCell(3);
        cell.setCellValue((String) "Username");
        cell.setCellStyle(headerStyle);

        cell = row.createCell(4);
        cell.setCellValue((String) "Password");
        cell.setCellStyle(headerStyle);

        Integer rowNum = 1;

        for (Object[] datatype : datatypes) {
            row = sheet.createRow(rowNum++);
            cell = row.createCell(0);
            cell.setCellValue((Integer) rowNum - 1);

            Integer colNum = 1;
            for (Object field : datatype) {
                cell = row.createCell(colNum++);
                if (field instanceof String) {
                    cell.setCellValue((String) field);
                } else if (field instanceof Integer) {
                    cell.setCellValue((Integer) field);
                }
            }
        }

        try {
            FileOutputStream outputStream = new FileOutputStream(localPath + "/" + file + ".xlsx");
            workbook.write(outputStream);
            workbook.close();

            System.out.println("file generated");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        
        return "build file done!";
    }
}