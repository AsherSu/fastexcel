package cn.idev.excel.analysis.v07.handlers.sax;

import cn.idev.excel.analysis.v07.handlers.*;
import cn.idev.excel.constant.ExcelXmlConstants;
import cn.idev.excel.context.xlsx.XlsxReadContext;
import java.util.HashMap;
import java.util.Map;
import lombok.extern.slf4j.Slf4j;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

/**
 *
 */
@Slf4j
public class XlsxRowHandler extends DefaultHandler {
    private final XlsxReadContext xlsxReadContext;
    private static final Map<String, XlsxTagHandler> XLSX_CELL_HANDLER_MAP = new HashMap<>(64);

    static {
        CellFormulaTagHandler cellFormulaTagHandler = new CellFormulaTagHandler();
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.CELL_FORMULA_TAG, cellFormulaTagHandler);
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.X_CELL_FORMULA_TAG, cellFormulaTagHandler);
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.NS2_CELL_FORMULA_TAG, cellFormulaTagHandler);
        CellInlineStringValueTagHandler cellInlineStringValueTagHandler = new CellInlineStringValueTagHandler();
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.CELL_INLINE_STRING_VALUE_TAG, cellInlineStringValueTagHandler);
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.X_CELL_INLINE_STRING_VALUE_TAG, cellInlineStringValueTagHandler);
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.NS2_CELL_INLINE_STRING_VALUE_TAG, cellInlineStringValueTagHandler);
        CellTagHandler cellTagHandler = new CellTagHandler();
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.CELL_TAG, cellTagHandler);
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.X_CELL_TAG, cellTagHandler);
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.NS2_CELL_TAG, cellTagHandler);
        CellValueTagHandler cellValueTagHandler = new CellValueTagHandler();
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.CELL_VALUE_TAG, cellValueTagHandler);
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.X_CELL_VALUE_TAG, cellValueTagHandler);
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.NS2_CELL_VALUE_TAG, cellValueTagHandler);
        CountTagHandler countTagHandler = new CountTagHandler();
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.DIMENSION_TAG, countTagHandler);
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.X_DIMENSION_TAG, countTagHandler);
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.NS2_DIMENSION_TAG, countTagHandler);
        HyperlinkTagHandler hyperlinkTagHandler = new HyperlinkTagHandler();
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.HYPERLINK_TAG, hyperlinkTagHandler);
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.X_HYPERLINK_TAG, hyperlinkTagHandler);
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.NS2_HYPERLINK_TAG, hyperlinkTagHandler);
        MergeCellTagHandler mergeCellTagHandler = new MergeCellTagHandler();
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.MERGE_CELL_TAG, mergeCellTagHandler);
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.X_MERGE_CELL_TAG, mergeCellTagHandler);
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.NS2_MERGE_CELL_TAG, mergeCellTagHandler);
        RowTagHandler rowTagHandler = new RowTagHandler();
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.ROW_TAG, rowTagHandler);
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.X_ROW_TAG, rowTagHandler);
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.NS2_ROW_TAG, rowTagHandler);
        FloatingImageTagHandler floatingImageTagHandler = new FloatingImageTagHandler();
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.IMAGE_TAG, floatingImageTagHandler);
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.X_IMAGE_TAG, floatingImageTagHandler);
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.NS2_IMAGE_TAG, floatingImageTagHandler);

        // 添加更多图片相关标签的支持
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.TWO_CELL_ANCHOR_TAG, floatingImageTagHandler);
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.X_TWO_CELL_ANCHOR_TAG, floatingImageTagHandler);
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.NS2_TWO_CELL_ANCHOR_TAG, floatingImageTagHandler);

        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.ONE_CELL_ANCHOR_TAG, floatingImageTagHandler);
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.X_ONE_CELL_ANCHOR_TAG, floatingImageTagHandler);
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.NS2_ONE_CELL_ANCHOR_TAG, floatingImageTagHandler);

        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.BLIP_TAG, floatingImageTagHandler);
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.X_BLIP_TAG, floatingImageTagHandler);
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.NS2_BLIP_TAG, floatingImageTagHandler);
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.A_BLIP_TAG, floatingImageTagHandler);

        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.FROM_TAG, floatingImageTagHandler);
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.X_FROM_TAG, floatingImageTagHandler);
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.NS2_FROM_TAG, floatingImageTagHandler);

        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.TO_TAG, floatingImageTagHandler);
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.X_TO_TAG, floatingImageTagHandler);
        XLSX_CELL_HANDLER_MAP.put(ExcelXmlConstants.NS2_TO_TAG, floatingImageTagHandler);

    }

    public XlsxRowHandler(XlsxReadContext xlsxReadContext) {
        this.xlsxReadContext = xlsxReadContext;
    }

    @Override
    public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
        XlsxTagHandler handler = XLSX_CELL_HANDLER_MAP.get(name);
        if (handler == null || !handler.support(xlsxReadContext)) {
            return;
        }
        xlsxReadContext.xlsxReadSheetHolder().getTagDeque().push(name);
        handler.startElement(xlsxReadContext, name, attributes);
    }

    @Override
    public void characters(char[] ch, int start, int length) throws SAXException {
        String currentTag = xlsxReadContext.xlsxReadSheetHolder().getTagDeque().peek();
        if (currentTag == null) {
            return;
        }
        XlsxTagHandler handler = XLSX_CELL_HANDLER_MAP.get(currentTag);
        if (handler == null || !handler.support(xlsxReadContext)) {
            return;
        }
        handler.characters(xlsxReadContext, ch, start, length);
    }

    @Override
    public void endElement(String uri, String localName, String name) throws SAXException {
        XlsxTagHandler handler = XLSX_CELL_HANDLER_MAP.get(name);
        if (handler == null || !handler.support(xlsxReadContext)) {
            return;
        }
        handler.endElement(xlsxReadContext, name);
        xlsxReadContext.xlsxReadSheetHolder().getTagDeque().pop();
    }
}
