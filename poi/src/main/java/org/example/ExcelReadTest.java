package org.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class ExcelReadTest {
    private String path = System.getProperty("user.dir");

    public void testRead03() throws IOException {
        //1.获取文件流
        FileInputStream inputStream = new FileInputStream(path+ File.separator+"test-03.xls");

        //2.创建工作簿，使用excel能操作的,Java均可以操作
        Workbook workbook = new HSSFWorkbook(inputStream);
        //获取工作表
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(0); //得到0行
        Cell cell = row.getCell(0);//得到0列
        //获取(0,0)值
        System.out.println(cell.getStringCellValue()); //获取字符串类型的值

        Cell cell1 = row.getCell(1);//得到1列
        System.out.println((int)cell1.getNumericCellValue());//获取数字类型的值
        inputStream.close();
    }

    public static void main(String[] args) throws IOException {
        //读取03和07基本一样
        new ExcelReadTest().testRead03();
        new ExcelReadTest().testRead07();
    }

    public void testRead07() throws IOException {
        //1.获取文件流
        FileInputStream inputStream = new FileInputStream(path+ File.separator+"test-07.xlsx");

        //2.创建工作簿，使用excel能操作的,Java均可以操作
        Workbook workbook = new XSSFWorkbook(inputStream);
        //获取工作表
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(0); //得到0行
        Cell cell = row.getCell(0);//得到0列
        //获取(0,0)值
        System.out.println(cell.getStringCellValue()); //获取字符串类型的值

        Cell cell1 = row.getCell(1);//得到1列
        System.out.println((int)cell1.getNumericCellValue());//获取数字类型的值
        inputStream.close();
    }
}
