package org.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelWriteTest {
    private String path = System.getProperty("user.dir");

    public  void testWrite03() throws IOException {
        //1. 创建工作簿
        Workbook workbook = new HSSFWorkbook();
        //2. 创建一个工作表
        Sheet sheet = workbook.createSheet("test观众统计表");
        //3. 创建第一行 0
        Row row1 = sheet.createRow(0);//默认从第一行,0开始
        //4. 创建第一行第一列(0,0)
        Cell cell11 = row1.createCell(0);// 构成(0,0)单元格
        cell11.setCellValue("今日新增观众");
        //第一行第二列(0,1)
        Cell cell12 = row1.createCell(1);// 构成(0,1)单元格
        cell12.setCellValue(666);

        //3. 创建第二行 1
        Row row2 = sheet.createRow(1);//默认从第一行,0开始
        //4. 创建第一行第一列(0,0)
        Cell cell21 = row2.createCell(0);// 构成(1,0)单元格
        cell21.setCellValue("统计时间");
        //第一行第二列(0,1)
        Cell cell22 = row2.createCell(1);// 构成(1,1)单元格
        String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell22.setCellValue(time);

        //生成一张表(IO流)
        FileOutputStream fos = new FileOutputStream(path+ File.separator+"test-03.xls");
        workbook.write(fos);
        fos.close();
        System.out.println("excel生成完毕.");
    }

    public static void main(String[] args) throws IOException {
        //new ExcelWriteTest().testWrite03();
        new ExcelWriteTest().testWrite07();
    }

    public  void testWrite07() throws IOException {
        //1. 创建工作簿
        Workbook workbook = new XSSFWorkbook();
        //2. 创建一个工作表
        Sheet sheet = workbook.createSheet("test观众统计表");
        //3. 创建第一行 0
        Row row1 = sheet.createRow(0);//默认从第一行,0开始
        //4. 创建第一行第一列(0,0)
        Cell cell11 = row1.createCell(0);// 构成(0,0)单元格
        cell11.setCellValue("今日新增观众");
        //第一行第二列(0,1)
        Cell cell12 = row1.createCell(1);// 构成(0,1)单元格
        cell12.setCellValue(666);

        //3. 创建第二行 1
        Row row2 = sheet.createRow(1);//默认从第一行,0开始
        //4. 创建第一行第一列(0,0)
        Cell cell21 = row2.createCell(0);// 构成(1,0)单元格
        cell21.setCellValue("统计时间");
        //第一行第二列(0,1)
        Cell cell22 = row2.createCell(1);// 构成(1,1)单元格
        String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell22.setCellValue(time);

        //生成一张表(IO流)
        FileOutputStream fos = new FileOutputStream(path+ File.separator+"test-07.xlsx");
        workbook.write(fos);
        fos.close();
        System.out.println("excel生成完毕.");
    }
}
