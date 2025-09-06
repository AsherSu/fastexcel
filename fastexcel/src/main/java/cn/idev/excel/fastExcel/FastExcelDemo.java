package cn.idev.excel.fastExcel;

import cn.idev.excel.ExcelWriter;
import cn.idev.excel.FastExcel;
import cn.idev.excel.context.AnalysisContext;
import cn.idev.excel.enums.CellExtraTypeEnum;
import cn.idev.excel.read.listener.ReadListener;
import cn.idev.excel.write.metadata.WriteSheet;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class FastExcelDemo {
    public static void main(String[] args) {
        String fileName = "E:\\project\\java\\swagger-demo\\1.xlsx";


//        // 方式1：简单写入
//        FastExcel.write(fileName, DemoData.class)
//                .sheet("模板")
//                .doWrite(data());
//
//        // 方式2：复杂写入
//        ExcelWriter excelWriter = FastExcel.write(fileName, DemoData.class).build();
//        WriteSheet writeSheet = FastExcel.writerSheet("数据表1").build();
//        excelWriter.write(data(), writeSheet);
//        excelWriter.finish();
        FastExcel.read(fileName, DemoData.class, (ReadListener<DemoData>) (data, context) -> {
                    System.out.println(data);
                })
                .extraRead(CellExtraTypeEnum.MERGE_IMAGE)
                .doReadAll();
    }

}
