package cn.idev.excel.fastExcel;

import cn.idev.excel.annotation.ExcelProperty;
import cn.idev.excel.converters.bytearray.ByteArrayImageConverter;
import lombok.Data;

@Data
public class DemoData {
    @ExcelProperty("1")
    private String name;

    @ExcelProperty("2")
    private String age;

    @ExcelProperty("3")
    private String birthDate;

    @ExcelProperty(value = "4", converter = ByteArrayImageConverter.class)
    private byte[] pic;

//    @ExcelProperty(value = "5", converter = ByteArrayImageConverter.class)
//    private byte[] pic;
}
