package org.example.easy;

import com.alibaba.excel.EasyExcel;

import java.io.File;

public class ReadTest {
    private final String path = System.getProperty("user.dir");

    public static void main(String[] args) {
        new ReadTest().simpleRead();
    }

    public void simpleRead() {
        String fileName = path+ File.separator+"easy.xlsx";
        // 这里 需要指定读用哪个class去读，然后读取第一个sheet 文件流会自动关闭
        EasyExcel.read(fileName, DemoData.class, new DemoDataListener()).sheet().doRead();
    }
}
