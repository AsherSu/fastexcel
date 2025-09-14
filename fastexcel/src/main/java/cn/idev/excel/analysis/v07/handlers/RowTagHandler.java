package cn.idev.excel.analysis.v07.handlers;

import cn.idev.excel.analysis.v07.XlsxSaxAnalyser;
import cn.idev.excel.constant.ExcelXmlConstants;
import cn.idev.excel.context.xlsx.XlsxReadContext;
import cn.idev.excel.enums.CellDataTypeEnum;
import cn.idev.excel.enums.RowTypeEnum;
import cn.idev.excel.metadata.Cell;
import cn.idev.excel.metadata.data.ReadCellData;
import cn.idev.excel.read.metadata.holder.ReadRowHolder;
import cn.idev.excel.read.metadata.holder.xlsx.XlsxReadSheetHolder;
import cn.idev.excel.util.PositionUtils;
import java.util.LinkedHashMap;
import java.util.Map;

import org.apache.commons.collections4.MapUtils;
import org.xml.sax.Attributes;

/**
 * Cell Handler
 *
 */
public class RowTagHandler extends AbstractXlsxTagHandler {

    @Override
    public void startElement(XlsxReadContext xlsxReadContext, String name, Attributes attributes) {
        XlsxReadSheetHolder xlsxReadSheetHolder = xlsxReadContext.xlsxReadSheetHolder();
        int rowIndex = PositionUtils.getRowByRowTagt(
                attributes.getValue(ExcelXmlConstants.ATTRIBUTE_R), xlsxReadSheetHolder.getRowIndex());
        String value = attributes.getValue(ExcelXmlConstants.SPAN);
        if (value!=null) {
            String totalCol = value.split(":")[1];
            xlsxReadSheetHolder.setTotalCol(Integer.parseInt(totalCol));
        }
        Integer lastRowIndex = xlsxReadContext.readSheetHolder().getRowIndex();
        while (lastRowIndex + 1 < rowIndex) {
            xlsxReadContext.readRowHolder(new ReadRowHolder(
                    lastRowIndex + 1,
                    RowTypeEnum.EMPTY,
                    xlsxReadSheetHolder.getGlobalConfiguration(),
                    new LinkedHashMap<Integer, Cell>()));
            xlsxReadContext.analysisEventProcessor().endRow(xlsxReadContext);
            xlsxReadSheetHolder.setColumnIndex(null);
            xlsxReadSheetHolder.setCellMap(new LinkedHashMap<Integer, Cell>());
            lastRowIndex++;
        }
        xlsxReadSheetHolder.setRowIndex(rowIndex);
    }

    @Override
    public void endElement(XlsxReadContext xlsxReadContext, String name) {
        XlsxReadSheetHolder xlsxReadSheetHolder = xlsxReadContext.xlsxReadSheetHolder();
        RowTypeEnum rowType = MapUtils.isEmpty(xlsxReadSheetHolder.getCellMap()) ? RowTypeEnum.EMPTY : RowTypeEnum.DATA;
        // It's possible that all of the cells in the row are empty
        if (rowType == RowTypeEnum.DATA) {
            boolean hasData = false;
            for (Cell cell : xlsxReadSheetHolder.getCellMap().values()) {
                if (!(cell instanceof ReadCellData)) {
                    hasData = true;
                    break;
                }
                ReadCellData<?> readCellData = (ReadCellData<?>) cell;
                if (readCellData.getType() != CellDataTypeEnum.EMPTY) {
                    hasData = true;
                    break;
                }
            }
            if (!hasData) {
                rowType = RowTypeEnum.EMPTY;
            }
        }
        Map<Integer, Cell> cellMap = xlsxReadSheetHolder.getCellMap();
        for (int i=0;i< (xlsxReadSheetHolder.getTotalCol()==null?0 :xlsxReadSheetHolder.getTotalCol());i++){
            if (cellMap.get(i)==null){
                Map<Integer, Map<Integer, XlsxSaxAnalyser.EmbeddedImage>> imageMap = xlsxReadSheetHolder.getImageMap();
                if (imageMap!=null && imageMap.get(xlsxReadSheetHolder.getRowIndex())!=null){
                    XlsxSaxAnalyser.EmbeddedImage embeddedImage = imageMap.get(xlsxReadSheetHolder.getRowIndex()).get(i);
                    if (embeddedImage!=null){
                        cellMap.put(i,new ReadCellData<>(embeddedImage.getData()));
                    }
                }
            }
        }
        xlsxReadContext.readRowHolder(new ReadRowHolder(
                xlsxReadSheetHolder.getRowIndex(),
                rowType,
                xlsxReadSheetHolder.getGlobalConfiguration(),
                xlsxReadSheetHolder.getCellMap()));
        xlsxReadContext.analysisEventProcessor().endRow(xlsxReadContext);
        xlsxReadSheetHolder.setColumnIndex(null);
        xlsxReadSheetHolder.setCellMap(new LinkedHashMap<>());
    }
}


