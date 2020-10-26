package org.example.easy;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;

import java.io.File;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class WriteTest {
    private final String path = System.getProperty("user.dir");

    public static void main(String[] args) {
        new WriteTest().simpleWrite();
    }

    public void simpleWrite(){ //写excel文件
        //写法1
        String fileName = path+ File.separator+"easy.xlsx";
        EasyExcel.write(fileName,DemoData.class).sheet("工作表名").doWrite(data());

        //写法2
        String fileName2 = path+ File.separator+"easy2.xlsx";
        ExcelWriter excelWriter = EasyExcel.write(fileName2,DemoData.class).build();
        WriteSheet writeSheet = EasyExcel.writerSheet("工作表名").build();
        excelWriter.write(data(),writeSheet);
        excelWriter.finish(); //关闭流
    }

    public List<DemoData> data(){
        List<DemoData> list = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            DemoData data = new DemoData();
            data.setString("字符串"+i);
            data.setDate(new Date());
            data.setDoubleData(0.56);
            data.setIgnore("ignore");
            list.add(data);
        }
        return list;
    }
}
