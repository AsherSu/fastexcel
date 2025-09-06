package cn.idev.excel.fastExcel;

import cn.idev.excel.context.AnalysisContext;
import cn.idev.excel.enums.CellExtraTypeEnum;
import cn.idev.excel.metadata.CellExtra;
import cn.idev.excel.read.listener.ReadListener;
import org.apache.commons.io.FileUtils;

public class DemoDataListener implements ReadListener<DemoData> {

    byte[] a;

    @Override
    public void extra(CellExtra extra, AnalysisContext context) {
        a=null;

        // Handle picture data
        if (CellExtraTypeEnum.MERGE_IMAGE == extra.getType()) {
            try {
                // 保存图片到文件系统
                String fileName = "image_" + context.readSheetHolder().getSheetNo() + "_" + extra.getRowIndex() + "_" + extra.getColumnIndex() + "." + extra.getImageFormat().toLowerCase();
                a=extra.getImageData();
                FileUtils.writeByteArrayToFile(new java.io.File(fileName), extra.getImageData());
                System.out.println("图片已保存: " + fileName);
            } catch (Exception e) {
                e.printStackTrace();
            }
            // Convert using PictureConverter
//            Object pictureData = PictureConverter.convertFromCellExtra(extra);
//            if (pictureData != null) {
//                pictureDataList.add(pictureData);
//
//                // Save picture to file (optional)
//                savePictureToFile(extra, context.readSheetHolder().getSheetNo());
//            }
        }
    }

    @Override
    public void invoke(DemoData data, AnalysisContext context) {
        data.setPic(a);
        System.out.println(data);
    }



    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        System.out.println("所有数据解析完成！");
    }
}
