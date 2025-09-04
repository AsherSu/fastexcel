package cn.idev.excel.analysis.v07.handlers;

import cn.idev.excel.context.xlsx.XlsxReadContext;
import cn.idev.excel.metadata.data.FormulaData;
import cn.idev.excel.read.metadata.holder.xlsx.XlsxReadSheetHolder;
import cn.idev.excel.util.DispimgFormulaUtil;
import lombok.extern.slf4j.Slf4j;
import org.xml.sax.Attributes;

/**
 * 公式单元格处理器 - 支持DISPIMG图片公式解析
 * Cell Formula Handler - Supports DISPIMG picture formula parsing
 */
@Slf4j
public class CellFormulaTagHandler extends AbstractXlsxTagHandler {

    @Override
    public void startElement(XlsxReadContext xlsxReadContext, String name, Attributes attributes) {
        XlsxReadSheetHolder xlsxReadSheetHolder = xlsxReadContext.xlsxReadSheetHolder();
        xlsxReadSheetHolder.setTempFormula(new StringBuilder());
    }

    @Override
    public void endElement(XlsxReadContext xlsxReadContext, String name) {
        XlsxReadSheetHolder xlsxReadSheetHolder = xlsxReadContext.xlsxReadSheetHolder();
        String formulaValue = xlsxReadSheetHolder.getTempFormula().toString();

        FormulaData formulaData = new FormulaData();
        formulaData.setFormulaValue(formulaValue);
        xlsxReadSheetHolder.getTempCellData().setFormulaData(formulaData);

        // 检查是否为DISPIMG公式并处理图片
        processDispimgFormulaIfNeeded(xlsxReadContext, formulaValue);
    }

    @Override
    public void characters(XlsxReadContext xlsxReadContext, char[] ch, int start, int length) {
        xlsxReadContext.xlsxReadSheetHolder().getTempFormula().append(ch, start, length);
    }

    /**
     * 如需要，处理DISPIMG公式图片
     */
    private void processDispimgFormulaIfNeeded(XlsxReadContext xlsxReadContext, String formula) {
        if (DispimgFormulaUtil.isDispimgFormula(formula)) {
            try {
                // 获取当前单元格的行列信息
                XlsxReadSheetHolder sheetHolder = xlsxReadContext.xlsxReadSheetHolder();
                Integer rowIndex = xlsxReadContext.readSheetHolder().getRowIndex();
                Integer columnIndex = sheetHolder.getColumnIndex();

                boolean processed = DispimgFormulaUtil.processDispimgFormula(
                        xlsxReadContext, formula, rowIndex, columnIndex);

                if (processed) {
                    log.debug("成功处理DISPIMG公式: row={}, col={}, formula={}",
                            rowIndex, columnIndex, formula);
                }

            } catch (Exception e) {
                log.warn("处理DISPIMG公式时发生错误: formula={}, error={}", formula, e.getMessage());
            }
        }
    }
}
