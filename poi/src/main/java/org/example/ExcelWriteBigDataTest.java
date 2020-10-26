package org.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelWriteBigDataTest {

    private String path = System.getProperty("user.dir");

    public void testWrite03BigData() throws IOException {
        //时间
        long begin = System.currentTimeMillis();

        //创建一个工作簿
        Workbook workbook = new HSSFWorkbook();
        //创建表
        Sheet sheet = workbook.createSheet();
        //写入数据, 最大行数65536,超过会异常
        for (int rowNum = 0; rowNum < 65536; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("over");
        FileOutputStream fos = new FileOutputStream(path+ File.separator+"test-03-bigData.xls");
        workbook.write(fos);
        fos.close();
        long end = System.currentTimeMillis();
        System.out.println((double)(end-begin)/1000);
    }

    public static void main(String[] args) throws IOException {
        //new ExcelWriteBigDataTest().testWrite03BigData();
        //new ExcelWriteBigDataTest().testWrite07BigData();
        new ExcelWriteBigDataTest().testWrite07BigDataS();
    }

    public void testWrite07BigDataS() throws IOException {
        //时间
        long begin = System.currentTimeMillis();

        //创建一个工作簿
        Workbook workbook = new SXSSFWorkbook(); //增强版XSSF写速度快
        //创建表
        Sheet sheet = workbook.createSheet();
        //写入数据, 行数70000
        for (int rowNum = 0; rowNum < 70000; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("over");
        FileOutputStream fos = new FileOutputStream(path+ File.separator+"test-07-bigDataS.xlsx");
        workbook.write(fos);
        fos.close();
        ((SXSSFWorkbook)workbook).dispose(); //清除临时文件！！！
        long end = System.currentTimeMillis();
        System.out.println((double)(end-begin)/1000);
    }

    public void testWrite07BigData() throws IOException {
        //时间
        long begin = System.currentTimeMillis();

        //创建一个工作簿
        Workbook workbook = new XSSFWorkbook(); //写速度慢
        //创建表
        Sheet sheet = workbook.createSheet();
        //写入数据, 行数65539
        for (int rowNum = 0; rowNum < 65539; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("over");
        FileOutputStream fos = new FileOutputStream(path+ File.separator+"test-07-bigData.xlsx");
        workbook.write(fos);
        fos.close();
        long end = System.currentTimeMillis();
        System.out.println((double)(end-begin)/1000);
    }
}
