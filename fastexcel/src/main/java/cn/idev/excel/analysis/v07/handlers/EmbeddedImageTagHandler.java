package cn.idev.excel.analysis.v07.handlers;

import cn.idev.excel.context.xlsx.XlsxReadContext;
import cn.idev.excel.metadata.data.FormulaData;
import cn.idev.excel.read.metadata.holder.xlsx.XlsxReadSheetHolder;
import lombok.extern.slf4j.Slf4j;
import org.xml.sax.Attributes;

/**
 * 内嵌图片处理器 - 处理单元格内嵌图片
 * Embedded Image Handler - Handles embedded images in cells
 */
@Slf4j
public class EmbeddedImageTagHandler extends AbstractXlsxTagHandler {

    @Override
    public void startElement(XlsxReadContext xlsxReadContext, String name, Attributes attributes) {
        XlsxReadSheetHolder xlsxReadSheetHolder = xlsxReadContext.xlsxReadSheetHolder();
        xlsxReadSheetHolder.setTempFormula(new StringBuilder());

        // 解析图片相关属性
        String imageId = attributes.getValue("r:embed");
        if (imageId != null) {
            log.info("Found embedded image with ID: {}", imageId);
            // 可以在这里加载图片数据，例如从 xlsxReadContext 获取包关系
            // byte[] imageData = loadImageFromPackage(xlsxReadContext, imageId);
            // xlsxReadSheetHolder.getTempCellData().setByteArrayValue(imageData);
        }
    }

    @Override
    public void endElement(XlsxReadContext xlsxReadContext, String name) {
        XlsxReadSheetHolder xlsxReadSheetHolder = xlsxReadContext.xlsxReadSheetHolder();
        String formulaValue = xlsxReadSheetHolder.getTempFormula().toString();

        FormulaData formulaData = new FormulaData();
        formulaData.setFormulaValue(formulaValue);
        xlsxReadSheetHolder.getTempCellData().setFormulaData(formulaData);

        // 保留占位符图像数据，实际图像处理需在后续加载
        xlsxReadSheetHolder.getTempCellData().setByteArrayValue("IMAGE_PLACiEHOLDER".getBytes());
    }

    @Override
    public void characters(XlsxReadContext xlsxReadContext, char[] ch, int start, int length) {
        xlsxReadContext.xlsxReadSheetHolder().getTempFormula().append(ch, start, length);
    }
}
