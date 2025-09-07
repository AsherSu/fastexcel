package cn.idev.excel.analysis.v07.handlers;

import cn.idev.excel.context.xlsx.XlsxReadContext;
import cn.idev.excel.enums.CellExtraTypeEnum;
import cn.idev.excel.metadata.CellExtra;
import cn.idev.excel.util.StringUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.opc.*;
import org.xml.sax.Attributes;

import java.io.InputStream;
import java.util.*;

/**
 * 浮动图片处理器 - 专门处理Excel中的浮动图片
 * Floating Picture Handler for XLSX format
 * Handles floating pictures in Excel drawings (twoCellAnchor/oneCellAnchor)
 *
 * 处理存储在xl/drawings/drawing*.xml中的传统浮动图片
 */
@Slf4j
public class FloatingImageTagHandler extends AbstractXlsxTagHandler {



    // 当前解析状态 - 用于SAX解析状态管理
    private static class ParseState {
        boolean inFrom = false;
        boolean inTo = false;
        int row1 = -1;
        int col1 = -1;
        String currentRid = null;

        void reset() {
            inFrom = inTo = false;
            row1 = col1 = -1;
            currentRid = null;
        }
    }

    private final ParseState parseState = new ParseState();

    @Override
    public boolean support(XlsxReadContext xlsxReadContext) {
        return xlsxReadContext.readWorkbookHolder().getExtraReadSet().contains(CellExtraTypeEnum.MERGE_IMAGE);
    }

    @Override
    public void startElement(XlsxReadContext xlsxReadContext, String name, Attributes attributes) {
        try {
            // 处理不同类型的图片元素
            switch (name.toLowerCase()) {
                case "pic":
                case "picture":
                    handleFloatingPicture(xlsxReadContext, attributes);
                    break;
                case "twocellanchor":
                case "onecellanchor":
                    parseState.reset();
                    break;
                case "from":
                    parseState.inFrom = true;
                    parseState.inTo = false;
                    break;
                case "to":
                    parseState.inFrom = false;
                    parseState.inTo = true;
                    break;
                case "row":
                    if (parseState.inFrom) {
                        parseState.row1 = parseIntegerFromText(attributes);
                    }
                    break;
                case "col":
                    if (parseState.inFrom) {
                        parseState.col1 = parseIntegerFromText(attributes);
                    }
                    break;
                case "blip":
                    // a:blip r:embed="rIdX"
                    String embed = attributes.getValue("r:embed");
                    if (embed == null) {
                        embed = attributes.getValue("embed");
                    }
                    if (!StringUtils.isEmpty(embed)) {
                        parseState.currentRid = embed;
                    }
                    break;
                default:
                    // 仅处理浮动图片相关标签，其他标签忽略
                    if (log.isDebugEnabled()) {
                        log.debug("忽略非浮动图片元素: name={}", name);
                    }
                    break;
            }
        } catch (Exception e) {
            log.warn("处理图片元素时发生错误: name={}, error={}", name, e.getMessage());
        }
    }

    @Override
    public void endElement(XlsxReadContext xlsxReadContext, String name) {
        try {
            switch (name.toLowerCase()) {
                case "from":
                case "to":
                    parseState.inFrom = parseState.inTo = false;
                    break;
                case "twocellanchor":
                case "onecellanchor":
                    // 锚点结束，若收集到完整信息，就创建图片对象
                    if (parseState.row1 >= 0 && parseState.col1 >= 0 && !StringUtils.isEmpty(parseState.currentRid)) {
                        createPictureFromAnchor(xlsxReadContext, parseState.currentRid, parseState.row1, parseState.col1);
                    }
                    parseState.reset();
                    break;
            }
        } catch (Exception e) {
            log.warn("结束图片元素处理时发生错误: name={}, error={}", name, e.getMessage());
        }
    }

    /**
     * 处理浮动图片 (传统drawing中的图片)
     */
    private void handleFloatingPicture(XlsxReadContext xlsxReadContext, Attributes attributes) {
        String rId = attributes.getValue("r:embed");
        if (StringUtils.isEmpty(rId)) {
            rId = attributes.getValue("embed");
        }
        if (StringUtils.isEmpty(rId)) {
            return;
        }

        PackageRelationshipCollection packageRelationshipCollection =
                xlsxReadContext.xlsxReadSheetHolder().getPackageRelationshipCollection();
        if (packageRelationshipCollection == null) {
            return;
        }

        String finalRId = rId;
        Optional.ofNullable(packageRelationshipCollection.getRelationshipByID(rId))
                .ifPresent(relationship -> {
                    try {
                        PackagePart picturePart = xlsxReadContext.xlsxReadSheetHolder()
                                .getPackagePart().getRelatedPart(relationship);

                        if (picturePart != null) {
                            createPictureFromPart(xlsxReadContext, picturePart,
                                    extractRowIndex(attributes), extractColumnIndex(attributes));
                        }
                    } catch (Exception e) {
                        log.warn("处理浮动图片时发生错误: rId={}, error={}", finalRId, e.getMessage());
                    }
                });
    }

    /**
     * 从锚点信息创建图片对象
     */
    private void createPictureFromAnchor(XlsxReadContext xlsxReadContext, String rId, int row, int col) {
        PackageRelationshipCollection packageRelationshipCollection =
                xlsxReadContext.xlsxReadSheetHolder().getPackageRelationshipCollection();
        if (packageRelationshipCollection == null) {
            return;
        }

        Optional.ofNullable(packageRelationshipCollection.getRelationshipByID(rId))
                .ifPresent(relationship -> {
                    try {
                        PackagePart picturePart = xlsxReadContext.xlsxReadSheetHolder()
                                .getPackagePart().getRelatedPart(relationship);

                        if (picturePart != null) {
                            createPictureFromPart(xlsxReadContext, picturePart, row, col);
                        }
                    } catch (Exception e) {
                        log.warn("从锚点创建图片时发生错误: rId={}, row={}, col={}, error={}", rId, row, col, e.getMessage());
                    }
                });
    }

    /**
     * 从PackagePart创建图片CellExtra对象
     */
    private void createPictureFromPart(XlsxReadContext xlsxReadContext, PackagePart picturePart,
                                       Integer rowIndex, Integer columnIndex) {
        try (InputStream inputStream = picturePart.getInputStream()) {
            byte[] pictureData = readInputStreamToByteArray(inputStream);
            String pictureFormat = determinePictureFormat(picturePart.getContentType(),
                    picturePart.getPartName().getName());

            // 使用有效的行列索引，如果未提供则使用0
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

            log.debug("成功解析图片: row={}, col={}, format={}, size={}bytes",
                    finalRowIndex, finalColumnIndex, pictureFormat, pictureData.length);

        } catch (Exception e) {
            log.warn("从PackagePart创建图片时发生错误: {}", e.getMessage());
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

        // 从文件名推断格式
        if (!StringUtils.isEmpty(partName)) {
            int dotIndex = partName.lastIndexOf('.');
            if (dotIndex > 0 && dotIndex < partName.length() - 1) {
                return partName.substring(dotIndex + 1).toLowerCase();
            }
        }

        return "bin";
    }

    /**
     * 从属性中提取行索引
     */
    private Integer extractRowIndex(Attributes attributes) {
        // 尝试多种可能的属性名
        String[] possibleNames = {"from_row", "row", "r", "fromRow"};
        for (String name : possibleNames) {
            String value = attributes.getValue(name);
            if (!StringUtils.isEmpty(value)) {
                try {
                    return Integer.parseInt(value);
                } catch (NumberFormatException e) {
                    // 继续尝试下一个
                }
            }
        }
        return null;
    }

    /**
     * 从属性中提取列索引
     */
    private Integer extractColumnIndex(Attributes attributes) {
        // 尝试多种可能的属性名
        String[] possibleNames = {"from_col", "col", "c", "fromCol"};
        for (String name : possibleNames) {
            String value = attributes.getValue(name);
            if (!StringUtils.isEmpty(value)) {
                try {
                    return Integer.parseInt(value);
                } catch (NumberFormatException e) {
                    // 继续尝试下一个
                }
            }
        }
        return null;
    }

    /**
     * 从属性中解析整数值
     */
    private int parseIntegerFromText(Attributes attributes) {
        // 这个方法需要在characters方法中配合使用
        // 这里先返回-1，实际解析在characters方法中进行
        return -1;
    }

    @Override
    public void characters(XlsxReadContext xlsxReadContext, char[] ch, int start, int length) {
        // 处理XML文本内容，主要用于解析row/col标签的值
        String currentTag = xlsxReadContext.xlsxReadSheetHolder().getTagDeque().peek();
        if (!StringUtils.isEmpty(currentTag)) {
            String text = new String(ch, start, length).trim();

            switch (currentTag.toLowerCase()) {
                case "row":
                    if (parseState.inFrom && !StringUtils.isEmpty(text)) {
                        try {
                            parseState.row1 = Integer.parseInt(text);
                        } catch (NumberFormatException e) {
                            log.debug("解析行号失败: {}", text);
                        }
                    }
                    break;
                case "col":
                    if (parseState.inFrom && !StringUtils.isEmpty(text)) {
                        try {
                            parseState.col1 = Integer.parseInt(text);
                        } catch (NumberFormatException e) {
                            log.debug("解析列号失败: {}", text);
                        }
                    }
                    break;
            }
        }
    }
}
