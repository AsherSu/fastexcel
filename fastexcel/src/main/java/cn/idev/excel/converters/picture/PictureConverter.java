package cn.idev.excel.converters.picture;

import cn.idev.excel.converters.Converter;
import cn.idev.excel.enums.CellDataTypeEnum;
import cn.idev.excel.enums.CellExtraTypeEnum;
import cn.idev.excel.metadata.CellExtra;
import cn.idev.excel.metadata.GlobalConfiguration;
import cn.idev.excel.metadata.data.ReadCellData;
import cn.idev.excel.metadata.data.WriteCellData;
import cn.idev.excel.metadata.property.ExcelContentProperty;

import java.util.Base64;

/**
 * Picture/Image data converter
 * Converts between Excel picture data and Java objects
 *
 * This converter handles picture data stored in CellExtra information
 * and converts it to various Java representations (byte[], Base64 String, etc.)
 *
 */
public class PictureConverter implements Converter<PictureData> {

    @Override
    public Class<?> supportJavaTypeKey() {
        return PictureData.class;
    }

    @Override
    public CellDataTypeEnum supportExcelTypeKey() {
        // Pictures are not stored in regular cells, but as extra information
        // This converter works with the extra information mechanism
        return CellDataTypeEnum.EMPTY;
    }

    @Override
    public PictureData convertToJavaData(
            ReadCellData<?> cellData,
            ExcelContentProperty contentProperty,
            GlobalConfiguration globalConfiguration) {

        // Check if cellData contains any binary data that could be picture bytes
        if (cellData != null && cellData.getData() instanceof byte[]) {
            PictureData pictureData = new PictureData();
            pictureData.setPictureBytes((byte[]) cellData.getData());

            // Convert to Base64 for easy handling
            String base64Data = Base64.getEncoder().encodeToString((byte[]) cellData.getData());
            pictureData.setBase64Data(base64Data);

            return pictureData;
        }

        return null;
    }

    @Override
    public WriteCellData<?> convertToExcelData(
            PictureData value,
            ExcelContentProperty contentProperty,
            GlobalConfiguration globalConfiguration) {

        // For writing pictures, FastExcel uses ImageData mechanism
        // This converter focuses on reading pictures from Excel
        // Writing pictures should use the existing ImageData approach

        return new WriteCellData<>(CellDataTypeEnum.EMPTY);
    }

    /**
     * Convert picture data from CellExtra to PictureData object
     * This method should be called from ReadListener.extra() when CellExtraTypeEnum.PICTURE is encountered
     */
    public static PictureData convertFromCellExtra(CellExtra cellExtra) {
        if (cellExtra == null || cellExtra.getType() != CellExtraTypeEnum.PICTURE) {
            return null;
        }

        PictureData pictureData = new PictureData();
        pictureData.setPictureBytes(cellExtra.getImageData());
        pictureData.setPictureFormat(cellExtra.getImageFormat());
        pictureData.setRowIndex(cellExtra.getRowIndex());
        pictureData.setColumnIndex(cellExtra.getColumnIndex());

        // Convert to Base64 for easy handling
        if (cellExtra.getImageData() != null) {
            String base64Data = Base64.getEncoder().encodeToString(cellExtra.getImageData());
            pictureData.setBase64Data(base64Data);
        }

        return pictureData;
    }
}

/**
 * Picture data holder class
 */
class PictureData {
    private byte[] pictureBytes;
    private String pictureFormat;
    private Integer rowIndex;
    private Integer columnIndex;
    private String base64Data;

    // Getters and setters
    public byte[] getPictureBytes() {
        return pictureBytes;
    }

    public void setPictureBytes(byte[] pictureBytes) {
        this.pictureBytes = pictureBytes;
    }

    public String getPictureFormat() {
        return pictureFormat;
    }

    public void setPictureFormat(String pictureFormat) {
        this.pictureFormat = pictureFormat;
    }

    public Integer getRowIndex() {
        return rowIndex;
    }

    public void setRowIndex(Integer rowIndex) {
        this.rowIndex = rowIndex;
    }

    public Integer getColumnIndex() {
        return columnIndex;
    }

    public void setColumnIndex(Integer columnIndex) {
        this.columnIndex = columnIndex;
    }

    public String getBase64Data() {
        return base64Data;
    }

    public void setBase64Data(String base64Data) {
        this.base64Data = base64Data;
    }
}
