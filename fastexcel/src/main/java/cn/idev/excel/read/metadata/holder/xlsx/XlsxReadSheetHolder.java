package cn.idev.excel.read.metadata.holder.xlsx;

import cn.idev.excel.analysis.v07.XlsxSaxAnalyser;
import cn.idev.excel.read.metadata.ReadSheet;
import cn.idev.excel.read.metadata.holder.ReadSheetHolder;
import cn.idev.excel.read.metadata.holder.ReadWorkbookHolder;

import java.net.URI;
import java.util.Deque;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import lombok.EqualsAndHashCode;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackageRelationshipCollection;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;

/**
 * sheet holder
 */
@Getter
@Setter
@EqualsAndHashCode
public class XlsxReadSheetHolder extends ReadSheetHolder {
    /**
     * Record the label of the current operation to prevent NPE.
     */
    private Deque<String> tagDeque;
    /**
     * Current Column
     */
    private Integer columnIndex;
    /**
     * Data for current label.
     */
    private StringBuilder tempData;
    /**
     * Formula for current label.
     */
    private StringBuilder tempFormula;
    /**
     * excel Relationship
     */
    private PackageRelationshipCollection packageRelationshipCollection;

    /**
     * Current sheet's package part
     */
    private PackagePart packagePart;

    private List<XlsxSaxAnalyser.EmbeddedImage> images;

    /**
     * 图片数据映射：行号 -> 列号 -> 图片对象
     * 用于在SAX解析过程中快速查找特定位置的图片
     */
    private Map<Integer, Map<Integer, XlsxSaxAnalyser.EmbeddedImage>> imageMap;

    public XlsxReadSheetHolder(ReadSheet readSheet, ReadWorkbookHolder readWorkbookHolder) {
        super(readSheet, readWorkbookHolder);
        this.tagDeque = new LinkedList<>();
        packageRelationshipCollection = ((XlsxReadWorkbookHolder) readWorkbookHolder)
                .getPackageRelationshipCollectionMap()
                .get(readSheet.getSheetNo());
        packagePart = ((XlsxReadWorkbookHolder) readWorkbookHolder)
                .getPackagePartMap()
                .get(readSheet.getSheetNo());
    }
}
