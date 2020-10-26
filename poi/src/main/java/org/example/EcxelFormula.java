package org.example;

import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class EcxelFormula {
    private String path = System.getProperty("user.dir");

    public static void main(String[] args) throws IOException {
        new EcxelFormula().formulaTest();
    }

    public void formulaTest() throws IOException {
        FileInputStream fis = new FileInputStream(path+ File.separator+"formula.xls");
        Workbook workbook = new HSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);

        Row row = sheet.getRow(4);
        Cell cell = row.getCell(0);

        //拿到计算公式 eval
        FormulaEvaluator formulaEvaluator = new HSSFFormulaEvaluator((HSSFWorkbook) workbook);

        //输出单元格的内容
        CellType cellType = cell.getCellType();
        switch (cellType){
            case FORMULA: //公式
                String formula = cell.getCellFormula();
                System.out.println(formula);
                //计算
                CellValue cellValue = formulaEvaluator.evaluate(cell);
                String valueStr = cellValue.formatAsString();
                System.out.println(valueStr);
                break;
        }
    }
}
