package cn.idev.excel.analysis.v07.handlers;

import cn.idev.excel.constant.ExcelXmlConstants;
import cn.idev.excel.context.xlsx.XlsxReadContext;
import cn.idev.excel.enums.RowTypeEnum;
import cn.idev.excel.metadata.Cell;
import cn.idev.excel.read.metadata.holder.ReadRowHolder;
import cn.idev.excel.read.metadata.holder.xlsx.XlsxReadSheetHolder;
import cn.idev.excel.util.PositionUtils;
import cn.idev.excel.util.StringUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackageRelationshipCollection;
import org.xml.sax.Attributes;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.DocumentBuilder;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;
import org.w3c.dom.Node;
import org.w3c.dom.NamedNodeMap;

import java.io.InputStream;
import java.util.LinkedHashMap;
import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import cn.idev.excel.metadata.data.FormulaData;

/**
 * 公式单元格处理器 - 支持DISPIMG图片公式解析
 * Cell Formula Handler - Supports DISPIMG picture formula parsing
 */
@Slf4j
public class CellFormulaTagHandler extends AbstractXlsxTagHandler {

    // 解析DISPIMG公式：支持可选的 '='、'@'、'_xlfn.' 前缀，分隔符支持 ',' 或 ';'，可带模式位
    private static final Pattern DISPIMG_PATTERN = Pattern.compile(
            "^(?:@)?(?:=)?(?:_xlfn\\.)?DISPIMG\\(\\s*\"([^\"]+)\"\\s*(?:[,;]\\s*(\\d+)\\s*)?\\)\\s*$",
            Pattern.CASE_INSENSITIVE);

    private static class DispimgInfo {
        final String imageId; // 如 ID_XXXX
        @SuppressWarnings("unused")
        final Integer mode;   // 0 裁剪, 1 缩放，未知为 null
        DispimgInfo(String imageId, Integer mode) { this.imageId = imageId; this.mode = mode; }
    }

    @Override
    public void startElement(XlsxReadContext xlsxReadContext, String name, Attributes attributes) {
        XlsxReadSheetHolder xlsxReadSheetHolder = xlsxReadContext.xlsxReadSheetHolder();
        int rowIndex = PositionUtils.getRowByRowTagt(
                attributes.getValue(ExcelXmlConstants.ATTRIBUTE_R), xlsxReadSheetHolder.getRowIndex());
        Integer lastRowIndex = xlsxReadContext.readSheetHolder().getRowIndex();
        while (lastRowIndex + 1 < rowIndex) {
            xlsxReadContext.readRowHolder(new ReadRowHolder(
                    lastRowIndex + 1,
                    RowTypeEnum.EMPTY,
                    xlsxReadSheetHolder.getGlobalConfiguration(),
                    new LinkedHashMap<>()));
            xlsxReadContext.analysisEventProcessor().endRow(xlsxReadContext);
            xlsxReadSheetHolder.setColumnIndex(null);
            xlsxReadSheetHolder.setCellMap(new LinkedHashMap<Integer, Cell>());
            lastRowIndex++;
        }
        xlsxReadSheetHolder.setRowIndex(rowIndex);
        // 为公式内容准备缓冲
        xlsxReadSheetHolder.setTempFormula(new StringBuilder());
    }

    @Override
    public void endElement(XlsxReadContext xlsxReadContext, String name) {
        try {
            XlsxReadSheetHolder xlsxReadSheetHolder = xlsxReadContext.xlsxReadSheetHolder();
            // 写回公式数据到单元格
            String formulaValue = xlsxReadSheetHolder.getTempFormula() == null
                    ? null
                    : xlsxReadSheetHolder.getTempFormula().toString();
            if (formulaValue != null) {
                FormulaData formulaData = new FormulaData();
                formulaData.setFormulaValue(formulaValue);
                xlsxReadSheetHolder.getTempCellData().setFormulaData(formulaData);
            }

            DispimgInfo info = parseDispimgInfo(formulaValue);
            if (info == null || StringUtils.isEmpty(info.imageId)) {
                return;
            }

            // 第一步：ID -> rId（来自 xl/cellimages.xml）
            PackagePart picturePart = resolvePicturePartByImageId(xlsxReadContext, info.imageId);
            if (picturePart == null) {
                // 兼容：有些文件直接在公式里就是 rId，尝试按 rId 查找
                picturePart = resolvePicturePartByRid(xlsxReadContext, info.imageId);
            }
            if (picturePart != null) {
                createPictureFromPart(xlsxReadContext, picturePart,
                        xlsxReadSheetHolder.getRowIndex(), xlsxReadSheetHolder.getColumnIndex());
                return;
            }

            // 无法解析则不再回退到基于SAX的锚点路径，保持“直接查找”的策略
        } catch (Exception e) {
            log.info("结束图片元素处理时发生错误: name={}, error={}", name, e.getMessage());
        }
    }

    @Override
    public void characters(XlsxReadContext xlsxReadContext, char[] ch, int start, int length) {
        XlsxReadSheetHolder xlsxReadSheetHolder = xlsxReadContext.xlsxReadSheetHolder();
        if (xlsxReadSheetHolder.getTempFormula() != null) {
            xlsxReadSheetHolder.getTempFormula().append(ch, start, length);
        }
    }


    /**
     * 从锚点信息创建图片对象
     */
    @SuppressWarnings("unused")
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
     * 在整个OPCPackage范围内，通过rId直接定位图片PackagePart（不进行SAX解析）
     */
    private PackagePart resolvePicturePartByRid(XlsxReadContext xlsxReadContext, String rId) {
        if (StringUtils.isEmpty(rId)) {
            return null;
        }
        OPCPackage opcPackage = xlsxReadContext.xlsxReadWorkbookHolder().getOpcPackage();
        if (opcPackage == null) {
            return null;
        }
        try {
            for (PackagePart part : opcPackage.getParts()) {
                // 跳过关系部件，避免对其调用 getRelationships 抛出异常
                try {
                    if (part.isRelationshipPart()) {
                        continue;
                    }
                } catch (Exception ignored) {
                    // 防御性判断，少数实现可能不支持该检测
                }
                PackageRelationshipCollection rels = part.getRelationships();
                if (rels == null) {
                    continue;
                }
                PackageRelationship relationship = rels.getRelationshipByID(rId);
                if (relationship == null) {
                    continue;
                }
                try {
                    PackagePart pic = part.getRelatedPart(relationship);
                    if (pic != null) {
                        // 简单判定：内容类型包含图片或路径位于 /xl/media/
                        String contentType = pic.getContentType();
                        String name = pic.getPartName() != null ? pic.getPartName().getName() : "";
                        if ((contentType != null && contentType.toLowerCase().contains("image/"))
                                || name.contains("/xl/media/")) {
                            return pic;
                        }
                    }
                } catch (Exception ignored) {
                    // 略过单个关系错误，继续查找
                }
            }
        } catch (Exception e) {
            log.debug("OPCPackage按rId查找图片失败: rId={}, error={}", rId, e.getMessage());
        }
        return null;
    }

    /**
     * 根据图片ID（如 ID_XXXX）解析到图片 PackagePart：
     * 1) 在 /xl/cellimages.xml 中查找 name="ID_XXXX"，取 r:embed/rId
     * 2) 通过 cellimages.xml 的 relationships 查到实际图片部件
     */
    private PackagePart resolvePicturePartByImageId(XlsxReadContext xlsxReadContext, String imageId) {
        if (StringUtils.isEmpty(imageId)) {
            return null;
        }
        OPCPackage opcPackage = xlsxReadContext.xlsxReadWorkbookHolder().getOpcPackage();
        if (opcPackage == null) {
            return null;
        }
        try {
            PackagePart cellImagesPart = null;
            for (PackagePart part : opcPackage.getParts()) {
                try {
                    if (part.isRelationshipPart()) {
                        continue;
                    }
                } catch (Exception ignored) {
                }
                String name = part.getPartName() != null ? part.getPartName().getName() : "";
                if (name.endsWith("/cellimages.xml") || name.equals("/xl/cellimages.xml") || name.equals("xl/cellimages.xml")) {
                    cellImagesPart = part;
                    break;
                }
            }
            if (cellImagesPart == null) {
                return null;
            }

            // 解析 cellimages.xml，找到指定 name 的条目，并读取其 r:embed/rId
            String embedRid = extractEmbedRidFromCellImages(cellImagesPart, imageId);
            if (StringUtils.isEmpty(embedRid)) {
                return null;
            }

            PackageRelationshipCollection rels = cellImagesPart.getRelationships();
            if (rels == null) {
                return null;
            }
            PackageRelationship relationship = rels.getRelationshipByID(embedRid);
            if (relationship == null) {
                return null;
            }
            try {
                return cellImagesPart.getRelatedPart(relationship);
            } catch (Exception e) {
                log.debug("通过 cellimages.xml 的关系解析图片失败: id={}, rId={}, error={}", imageId, embedRid, e.getMessage());
                return null;
            }
        } catch (Exception e) {
            log.debug("解析图片ID映射失败: imageId={}, error={}", imageId, e.getMessage());
            return null;
        }
    }

    private String extractEmbedRidFromCellImages(PackagePart cellImagesPart, String imageId) {
        try (InputStream is = cellImagesPart.getInputStream()) {
            DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
            dbf.setNamespaceAware(true);
            DocumentBuilder db = dbf.newDocumentBuilder();
            Document doc = db.parse(is);

            // 遍历全部元素，定位 xdr:cNvPr[name=ID_xxx]
            NodeList all = doc.getElementsByTagName("*");
            for (int i = 0; i < all.getLength(); i++) {
                Node node = all.item(i);
                if (!"cNvPr".equalsIgnoreCase(localName(node))) {
                    continue;
                }
                NamedNodeMap attrs = node.getAttributes();
                if (attrs == null || attrs.getLength() == 0) {
                    continue;
                }
                String nameAttr = getAttr(attrs, "name");
                if (nameAttr == null || !imageId.equals(nameAttr)) {
                    continue;
                }
                // 找到父系的 xdr:pic，再在其子树中查找 a:blip，取 r:embed
                Node pic = findAncestorByLocalName(node, "pic");
                if (pic == null) {
                    continue;
                }
                Node blip = findDescendantByLocalName(pic, "blip");
                if (blip == null) {
                    continue;
                }
                NamedNodeMap blipAttrs = blip.getAttributes();
                if (blipAttrs == null) {
                    continue;
                }
                String embed = getAttr(blipAttrs, "r:embed");
                if (embed == null) {
                    embed = getAttrEndsWith(blipAttrs, ":embed");
                    if (embed == null) {
                        embed = getAttr(blipAttrs, "embed");
                    }
                }
                if (!StringUtils.isEmpty(embed)) {
                    return embed;
                }
            }
        } catch (Exception e) {
            log.debug("解析 cellimages.xml 失败: error={}", e.getMessage());
        }
        return null;
    }

    private String localName(Node node) {
        if (node == null) {
            return null;
        }
        String ln = node.getLocalName();
        if (ln != null) {
            return ln;
        }
        String nm = node.getNodeName();
        if (nm == null) {
            return null;
        }
        int idx = nm.indexOf(':');
        return idx >= 0 ? nm.substring(idx + 1) : nm;
    }

    private Node findAncestorByLocalName(Node node, String targetLocalName) {
        Node p = node;
        while (p != null) {
            if (targetLocalName.equalsIgnoreCase(localName(p))) {
                return p;
            }
            p = p.getParentNode();
        }
        return null;
    }

    private Node findDescendantByLocalName(Node node, String targetLocalName) {
        if (node == null) {
            return null;
        }
        NodeList children = node.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            Node c = children.item(i);
            if (targetLocalName.equalsIgnoreCase(localName(c))) {
                return c;
            }
            Node hit = findDescendantByLocalName(c, targetLocalName);
            if (hit != null) {
                return hit;
            }
        }
        return null;
    }

    private String getAttr(NamedNodeMap attrs, String key) {
        Node n = attrs.getNamedItem(key);
        return n == null ? null : n.getNodeValue();
    }

    private String getAttrEndsWith(NamedNodeMap attrs, String suffix) {
        for (int i = 0; i < attrs.getLength(); i++) {
            Node a = attrs.item(i);
            if (a != null) {
                String nm = a.getNodeName();
                if (nm != null && nm.endsWith(suffix)) {
                    return a.getNodeValue();
                }
            }
        }
        return null;
    }

    private DispimgInfo parseDispimgInfo(String formula) {
        if (StringUtils.isEmpty(formula)) {
            return null;
        }
        Matcher m = DISPIMG_PATTERN.matcher(formula.trim());
        if (!m.matches()) {
            return null;
        }
        String id = m.group(1);
        Integer mode = null;
        try {
            String g2 = m.group(2);
            if (!StringUtils.isEmpty(g2)) {
                mode = Integer.parseInt(g2);
            }
        } catch (Exception ignore) {
        }
        return new DispimgInfo(id, mode);
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

            XlsxReadSheetHolder xlsxReadSheetHolder = xlsxReadContext.xlsxReadSheetHolder();
            xlsxReadSheetHolder.getTempCellData().setByteArrayValue(pictureData);

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
     * 解析DISPIMG公式中的图片ID
     */
    public static String parseDispimgId(String formula) {
        if (StringUtils.isEmpty(formula)) {
            return null;
        }

        Matcher matcher = DISPIMG_PATTERN.matcher(formula.trim());
        return matcher.matches() ? matcher.group(1) : null;
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




}
