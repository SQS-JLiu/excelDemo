package org.example;


import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class ExcelDataTypeTest {
    private String path = System.getProperty("user.dir");

    public static void main(String[] args) throws IOException {
        new ExcelDataTypeTest().testCellType();
    }

    public void testCellType() throws IOException {
        //获取文件流
        FileInputStream fis = new FileInputStream(path+ File.separator+"test-03.xls");
        Workbook workbook = new HSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(0);
        if(row != null){
            int cellCount = row.getPhysicalNumberOfCells();
            for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                Cell cell = row.getCell(cellNum);
                if(cell != null){
                    CellType cellType = cell.getCellType();
                    String cellValue="";
                    switch (cellType){
                        case BLANK:
                            System.out.println("Blank type");
                            break;
                        case BOOLEAN:
                            System.out.println("Boolean type");
                            cellValue = String.valueOf(cell.getBooleanCellValue());
                            break;
                        case STRING:
                            System.out.println("String type");
                            cellValue = cell.getStringCellValue();
                            break;
                        case NUMERIC:
                            System.out.println("Numeric type");
                            cellValue = String.valueOf(cell.getNumericCellValue());
                            break;
                    }
                    System.out.println(cellValue);
                }
            }
        }
    }
}
