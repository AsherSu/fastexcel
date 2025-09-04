package cn.idev.excel.metadata;

import cn.idev.excel.constant.ExcelXmlConstants;
import cn.idev.excel.enums.CellExtraTypeEnum;
import org.apache.poi.ss.util.CellReference;

/**
 * Cell extra information.
 *
 *
 */
public class CellExtra extends AbstractCell {
    /**
     * Cell extra type
     */
    private CellExtraTypeEnum type;
    /**
     * Cell extra data
     */
    private String text;
    /**
     * First row index, if this object is an interval
     */
    private Integer firstRowIndex;
    /**
     * Last row index, if this object is an interval
     */
    private Integer lastRowIndex;
    /**
     * First column index, if this object is an interval
     */
    private Integer firstColumnIndex;
    /**
     * Last column index, if this object is an interval
     */
    private Integer lastColumnIndex;
    /**
     * Picture data (byte array), only for PICTURE type
     */
    private byte[] imageData;
    /**
     * Picture format (e.g., "PNG", "JPG", "GIF"), only for PICTURE type
     */
    private String imageFormat;

    public CellExtra(CellExtraTypeEnum type, String text, String range) {
        super();
        this.type = type;
        this.text = text;
        String[] ranges = range.split(ExcelXmlConstants.CELL_RANGE_SPLIT);
        CellReference first = new CellReference(ranges[0]);
        CellReference last = first;
        this.firstRowIndex = first.getRow();
        this.firstColumnIndex = (int) first.getCol();
        setRowIndex(this.firstRowIndex);
        setColumnIndex(this.firstColumnIndex);
        if (ranges.length > 1) {
            last = new CellReference(ranges[1]);
        }
        this.lastRowIndex = last.getRow();
        this.lastColumnIndex = (int) last.getCol();
    }

    public CellExtra(CellExtraTypeEnum type, String text, Integer rowIndex, Integer columnIndex) {
        this(type, text, rowIndex, rowIndex, columnIndex, columnIndex);
    }

    public CellExtra(
            CellExtraTypeEnum type,
            String text,
            Integer firstRowIndex,
            Integer lastRowIndex,
            Integer firstColumnIndex,
            Integer lastColumnIndex) {
        super();
        setRowIndex(firstRowIndex);
        setColumnIndex(firstColumnIndex);
        this.type = type;
        this.text = text;
        this.firstRowIndex = firstRowIndex;
        this.firstColumnIndex = firstColumnIndex;
        this.lastRowIndex = lastRowIndex;
        this.lastColumnIndex = lastColumnIndex;
    }

    /**
     * Constructor for picture type CellExtra
     */
    public CellExtra(CellExtraTypeEnum type, byte[] pictureData, String imageFormat,
                     Integer rowIndex, Integer columnIndex) {
        this(type, null, rowIndex, rowIndex, columnIndex, columnIndex);
        this.imageData = pictureData;
        this.imageFormat = imageFormat;
    }

    public CellExtraTypeEnum getType() {
        return type;
    }

    public void setType(CellExtraTypeEnum type) {
        this.type = type;
    }

    public String getText() {
        return text;
    }

    public void setText(String text) {
        this.text = text;
    }

    public Integer getFirstRowIndex() {
        return firstRowIndex;
    }

    public void setFirstRowIndex(Integer firstRowIndex) {
        this.firstRowIndex = firstRowIndex;
    }

    public Integer getFirstColumnIndex() {
        return firstColumnIndex;
    }

    public void setFirstColumnIndex(Integer firstColumnIndex) {
        this.firstColumnIndex = firstColumnIndex;
    }

    public Integer getLastRowIndex() {
        return lastRowIndex;
    }

    public void setLastRowIndex(Integer lastRowIndex) {
        this.lastRowIndex = lastRowIndex;
    }

    public Integer getLastColumnIndex() {
        return lastColumnIndex;
    }

    public void setLastColumnIndex(Integer lastColumnIndex) {
        this.lastColumnIndex = lastColumnIndex;
    }

    public byte[] getImageData() {
        return imageData;
    }

    public void setImageData(byte[] imageData) {
        this.imageData = imageData;
    }

    public String getImageFormat() {
        return imageFormat;
    }

    public void setImageFormat(String imageFormat) {
        this.imageFormat = imageFormat;
    }
}
