package cn.idev.excel.analysis.v03.handlers;

import cn.idev.excel.analysis.v03.IgnorableXlsRecordHandler;
import cn.idev.excel.context.xls.XlsReadContext;
import cn.idev.excel.enums.CellExtraTypeEnum;
import cn.idev.excel.metadata.CellExtra;
import org.apache.poi.hssf.record.ObjRecord;
import org.apache.poi.hssf.record.Record;

/**
 * Picture Record Handler for XLS format
 * Handles ObjRecord that contains picture information
 *
 */
public class ImageRecordHandler extends AbstractXlsRecordHandler implements IgnorableXlsRecordHandler {

    @Override
    public boolean support(XlsReadContext xlsReadContext, Record record) {
        return xlsReadContext.readWorkbookHolder().getExtraReadSet().contains(CellExtraTypeEnum.PICTURE)
                && record instanceof ObjRecord;
    }

    @Override
    public void processRecord(XlsReadContext xlsReadContext, Record record) {
        if (!(record instanceof ObjRecord)) {
            return;
        }

        ObjRecord objRecord = (ObjRecord) record;

        try {
            // In XLS format, pictures are stored as OBJ records
            // This is a simplified implementation - real implementation would need
            // to parse the OBJ record structure to extract picture data

            // For demonstration purposes, we create a placeholder
            // Real implementation would extract actual picture data from the record
            byte[] pictureData = extractPictureData(objRecord);
            String pictureFormat = determinePictureFormat(objRecord);

            if (pictureData != null && pictureData.length > 0) {
                // Extract position information
                Integer rowIndex = extractRowIndex(objRecord);
                Integer columnIndex = extractColumnIndex(objRecord);

                if (rowIndex != null && columnIndex != null) {
                    CellExtra cellExtra = new CellExtra(
                            CellExtraTypeEnum.PICTURE,
                            pictureData,
                            pictureFormat,
                            rowIndex,
                            columnIndex
                    );

                    xlsReadContext.xlsReadSheetHolder().setCellExtra(cellExtra);
                    xlsReadContext.analysisEventProcessor().extra(xlsReadContext);
                }
            }
        } catch (Exception e) {
            // Log error but don't break the reading process
            // Following the error tolerance principle
        }
    }

    private byte[] extractPictureData(ObjRecord objRecord) {
        // This is a placeholder implementation
        // Real implementation would parse the OBJ record to extract picture data
        // XLS format stores pictures in a complex way involving multiple records

        // For now, return null to indicate no picture data found
        // A complete implementation would require:
        // 1. Parse OBJ record to find picture references
        // 2. Look up corresponding IMDATA records
        // 3. Extract and decompress picture data

        return null;
    }

    private String determinePictureFormat(ObjRecord objRecord) {
        // Determine picture format from OBJ record
        // This is a simplified implementation
        return "UNKNOWN";
    }

    private Integer extractRowIndex(ObjRecord objRecord) {
        // Extract row information from OBJ record
        // This would require parsing the anchor information in the record
        return null;
    }

    private Integer extractColumnIndex(ObjRecord objRecord) {
        // Extract column information from OBJ record
        // This would require parsing the anchor information in the record
        return null;
    }
}
