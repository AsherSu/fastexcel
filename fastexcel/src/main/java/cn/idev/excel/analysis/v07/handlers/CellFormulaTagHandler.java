package cn.idev.excel.analysis.v07.handlers;

import cn.idev.excel.context.xlsx.XlsxReadContext;
import cn.idev.excel.enums.CellExtraTypeEnum;
import cn.idev.excel.metadata.CellExtra;
import cn.idev.excel.metadata.data.FormulaData;
import cn.idev.excel.read.metadata.holder.xlsx.XlsxReadSheetHolder;
import cn.idev.excel.util.DispimgFormulaUtil;
import cn.idev.excel.util.StringUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.xml.sax.Attributes;

import java.io.InputStream;
import java.util.Map;
import java.util.HashMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 公式单元格处理器 - 支持DISPIMG图片公式解析
 * Cell Formula Handler - Supports DISPIMG picture formula parsing
 */
@Slf4j
public class CellFormulaTagHandler extends AbstractXlsxTagHandler {
    
    // 解析DISPIMG公式的正则表达式 - 兼容WPS格式
    private static final Pattern DISPIMG_PATTERN = Pattern.compile(
            "^(?:@)?(?:_xlfn\\.)?DISPIMG\\(\\s*\"([^\"]+)\"(?:\\s*,\\s*\\d+\\s*)?\\)\\s*$",
            Pattern.CASE_INSENSITIVE);

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
        // 不再调用setByteArrayValue()方法，因为该方法需要参数

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
        if (isDispimgFormula(formula)) {
            try {
                // 获取当前单元格的行列信息
                XlsxReadSheetHolder sheetHolder = xlsxReadContext.xlsxReadSheetHolder();
                Integer rowIndex = xlsxReadContext.readSheetHolder().getRowIndex();
                Integer columnIndex = sheetHolder.getColumnIndex();
                
                // 解析DISPIMG公式中的图片ID
                String imageId = parseDispimgId(formula);
                if (!StringUtils.isEmpty(imageId)) {
                    // 处理DISPIMG图片
                    boolean processed = processDispimgPicture(xlsxReadContext, imageId, rowIndex, columnIndex);
                    
                    if (processed) {
                        log.debug("成功处理DISPIMG公式: row={}, col={}, formula={}",
                                rowIndex, columnIndex, formula);
                    }
                }

            } catch (Exception e) {
                log.warn("处理DISPIMG公式时发生错误: formula={}, error={}", formula, e.getMessage());
            }
        }
    }
    
    /**
     * 解析DISPIMG公式中的图片ID
     */
    private String parseDispimgId(String formula) {
        if (StringUtils.isEmpty(formula)) {
            return null;
        }

        Matcher matcher = DISPIMG_PATTERN.matcher(formula.trim());
        return matcher.matches() ? matcher.group(1) : null;
    }

    /**
     * 检查公式是否是DISPIMG类型
     */
    private boolean isDispimgFormula(String formula) {
        return !StringUtils.isEmpty(parseDispimgId(formula));
    }
    
    /**
     * 处理DISPIMG图片
     */
    private boolean processDispimgPicture(XlsxReadContext xlsxReadContext, String imageId, 
                                         Integer rowIndex, Integer columnIndex) {
        if (!xlsxReadContext.readWorkbookHolder().getExtraReadSet().contains(CellExtraTypeEnum.MERGE_IMAGE)) {
            return false;
        }

        try {
            OPCPackage pkg = xlsxReadContext.xlsxReadWorkbookHolder().getOpcPackage();
            Map<String, PackagePart> id2ImagePart = buildImageIdMapping(pkg);

            PackagePart imagePart = id2ImagePart.get(imageId);
            if (imagePart != null) {
                createCellExtraFromDispimgImage(xlsxReadContext, imagePart, imageId, rowIndex, columnIndex);
                return true;
            }
        } catch (Exception e) {
            log.warn("处理DISPIMG图片时发生异常: imageId={}, error={}", imageId, e.getMessage());
        }

        return false;
    }
    
    /**
     * 构建图片ID到PackagePart的映射
     */
    private Map<String, PackagePart> buildImageIdMapping(OPCPackage pkg) {
        Map<String, PackagePart> id2ImagePart = new HashMap<>();

        try {
            // 查找所有cellimages*.xml Part
            for (PackagePart part : pkg.getParts()) {
                String partName = part.getPartName().getName().toLowerCase();
                if (partName.contains("cellimages") && partName.endsWith(".xml")) {
                    parseCellImagesForDispimg(part, id2ImagePart);
                }
            }

        } catch (Exception e) {
            log.warn("构建图片映射时发生异常: {}", e.getMessage());
        }

        return id2ImagePart;
    }
    
    /**
     * 解析CellImages XML文件以构建DISPIMG映射
     */
    private void parseCellImagesForDispimg(PackagePart cellImagesPart, Map<String, PackagePart> id2ImagePart) {
        try {
            // 首先建立该Part的rId到图片Part的映射
            Map<String, PackagePart> rid2ImagePart = new HashMap<>();
            for (PackageRelationship rel : cellImagesPart.getRelationships()) {
                String relType = rel.getRelationshipType();
                if (relType != null && relType.toLowerCase().endsWith("/image")) {
                    PackagePart imagePart = cellImagesPart.getRelatedPart(rel);
                    if (imagePart != null) {
                        rid2ImagePart.put(rel.getId(), imagePart);
                    }
                }
            }

            if (rid2ImagePart.isEmpty()) {
                return;
            }
            
            // TODO: 实现XML解析以找到name="ID_xxx"和r:embed的对应关系
            // 这里可以添加完整的XML解析逻辑来处理cellImages.xml文件
            
        } catch (Exception e) {
            log.warn("解析CellImages文件时发生异常: {}", e.getMessage());
        }
    }
    
    /**
     * 从DISPIMG图片创建CellExtra
     */
    private void createCellExtraFromDispimgImage(XlsxReadContext xlsxReadContext, PackagePart imagePart,
                                                String imageId, Integer rowIndex, Integer columnIndex) {
        try (InputStream inputStream = imagePart.getInputStream()) {
            byte[] pictureData = readInputStreamToByteArray(inputStream);
            String pictureFormat = determinePictureFormat(imagePart.getContentType(),
                    imagePart.getPartName().getName());

            // 使用提供的行列索引，如果为null则使用0
            int finalRowIndex = rowIndex != null ? rowIndex : 0;
            int finalColumnIndex = columnIndex != null ? columnIndex : 0;

            CellExtra cellExtra = new CellExtra(
                    CellExtraTypeEnum.MERGE_IMAGE,
                    pictureData,
                    pictureFormat,
                    finalRowIndex,
                    finalColumnIndex
            );

            xlsxReadContext.readSheetHolder().setCellExtra(cellExtra);
            xlsxReadContext.analysisEventProcessor().extra(xlsxReadContext);

        } catch (Exception e) {
            log.warn("从DISPIMG图片创建CellExtra时发生异常: {}", e.getMessage());
        }
    }
    
    private static byte[] readInputStreamToByteArray(InputStream inputStream) throws Exception {
        java.io.ByteArrayOutputStream buffer = new java.io.ByteArrayOutputStream();
        byte[] data = new byte[1024];
        int nRead;
        while ((nRead = inputStream.read(data, 0, data.length)) != -1) {
            buffer.write(data, 0, nRead);
        }
        return buffer.toByteArray();
    }
    
    /**
     * 判断图片格式
     */
    private String determinePictureFormat(String contentType, String partName) {
        if (!StringUtils.isEmpty(contentType)) {
            String ct = contentType.toLowerCase();
            if (ct.contains("png")) return "png";
            if (ct.contains("jpeg") || ct.contains("jpg")) return "jpg";
            if (ct.contains("gif")) return "gif";
            if (ct.contains("bmp")) return "bmp";
            if (ct.contains("tiff")) return "tiff";
            if (ct.contains("webp")) return "webp";
        }

        if (!StringUtils.isEmpty(partName)) {
            int dotIndex = partName.lastIndexOf('.');
            if (dotIndex > 0 && dotIndex < partName.length() - 1) {
                return partName.substring(dotIndex + 1).toLowerCase();
            }
        }

        return "bin";
    }
}