package cn.idev.excel.util;

import cn.idev.excel.context.xlsx.XlsxReadContext;
import cn.idev.excel.enums.CellExtraTypeEnum;
import cn.idev.excel.metadata.CellExtra;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;

import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamConstants;
import javax.xml.stream.XMLStreamReader;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * DISPIMG公式图片处理器
 * DISPIMG Formula Picture Handler
 *
 * 专门处理WPS等软件中使用的DISPIMG公式引用图片
 */
public class DispimgFormulaUtil {

    // 解析DISPIMG公式的正则表达式 - 兼容多种格式
    private static final Pattern DISPIMG_PATTERN = Pattern.compile(
            "^(?:@)?(?:_xlfn\\.)?DISPIMG\\(\\s*\"([^\"]+)\"(?:\\s*,\\s*\\d+\\s*)?\\)\\s*$",
            Pattern.CASE_INSENSITIVE);

    // 全局图片ID到PackagePart的映射缓存 - 避免重复解析
    private static final Map<OPCPackage, Map<String, PackagePart>> PACKAGE_IMAGE_CACHE = new ConcurrentHashMap<>();

    /**
     * 处理DISPIMG公式引用的图片
     *
     * @param xlsxReadContext XLSX读取上下文
     * @param formula 公式内容
     * @param rowIndex 行索引
     * @param columnIndex 列索引
     * @return 是否成功处理了DISPIMG公式
     */
    public static boolean processDispimgFormula(XlsxReadContext xlsxReadContext, String formula,
                                               Integer rowIndex, Integer columnIndex) {
        if (!xlsxReadContext.readWorkbookHolder().getExtraReadSet().contains(CellExtraTypeEnum.MERGE_IMAGE)) {
            return false;
        }

        String imageId = parseDispimgId(formula);
        if (StringUtils.isEmpty(imageId)) {
            return false;
        }

        try {
            OPCPackage pkg = xlsxReadContext.xlsxReadWorkbookHolder().getOpcPackage();
            Map<String, PackagePart> id2ImagePart = getOrBuildImageMapping(pkg);

            PackagePart imagePart = id2ImagePart.get(imageId);
            if (imagePart != null) {
                createCellExtraFromDispimgImage(xlsxReadContext, imagePart, imageId, rowIndex, columnIndex);
                return true;
            }
        } catch (Exception e) {
        }

        return false;
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
     * 检查公式是否是DISPIMG类型
     */
    public static boolean isDispimgFormula(String formula) {
        return !StringUtils.isEmpty(parseDispimgId(formula));
    }

    /**
     * 获取或构建图片ID到PackagePart的映射
     */
    private static Map<String, PackagePart> getOrBuildImageMapping(OPCPackage pkg) {
        return PACKAGE_IMAGE_CACHE.computeIfAbsent(pkg, DispimgFormulaUtil::buildImageIdMapping);
    }

    /**
     * 构建图片ID到PackagePart的映射
     */
    private static Map<String, PackagePart> buildImageIdMapping(OPCPackage pkg) {
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
        }

        return id2ImagePart;
    }

    /**
     * 解析CellImages XML文件以构建DISPIMG映射
     */
    private static void parseCellImagesForDispimg(PackagePart cellImagesPart, Map<String, PackagePart> id2ImagePart) {
        try {
            // 首先建立该Part的rId到图片Part的映射
            Map<String, PackagePart> rid2ImagePart = new HashMap<>();
            for (PackageRelationship rel : cellImagesPart.getRelationships()) {
                String relType = rel.getRelationshipType();
                if (relType != null && relType.toLowerCase().endsWith("/image")) {
                    PackagePart imagePart = getRelatedPart(cellImagesPart.getPackage(), cellImagesPart, rel);
                    if (imagePart != null) {
                        rid2ImagePart.put(rel.getId(), imagePart);
                    }
                }
            }

            if (rid2ImagePart.isEmpty()) {
                return;
            }

            // 解析XML以找到name="ID_xxx"和r:embed的对应关系
            try (InputStream xmlStream = cellImagesPart.getInputStream()) {
                XMLInputFactory factory = XMLInputFactory.newInstance();
                XMLStreamReader reader = factory.createXMLStreamReader(xmlStream);

                String currentImageId = null;
                String currentRid = null;

                while (reader.hasNext()) {
                    int event = reader.next();

                    if (event == XMLStreamConstants.START_ELEMENT) {
                        // 检查所有属性
                        for (int i = 0; i < reader.getAttributeCount(); i++) {
                            String attrName = reader.getAttributeLocalName(i);
                            String attrValue = reader.getAttributeValue(i);

                            // 查找name="ID_xxx"
                            if ("name".equals(attrName) && attrValue != null && attrValue.startsWith("ID_")) {
                                currentImageId = attrValue;
                            }

                            // 查找r:embed="rIdX"
                            if ("embed".equals(attrName)) {
                                currentRid = attrValue;
                            }
                        }
                    } else if (event == XMLStreamConstants.END_ELEMENT) {
                        String localName = reader.getLocalName();

                        // 在合适的结束标签处建立映射
                        if (("twoCellAnchor".equals(localName) || "oneCellAnchor".equals(localName) ||
                             "pic".equals(localName) || "cellImage".equals(localName)) &&
                            currentImageId != null && currentRid != null) {

                            PackagePart imagePart = rid2ImagePart.get(currentRid);
                            if (imagePart != null) {
                                id2ImagePart.put(currentImageId, imagePart);
                            }

                            // 重置状态
                            currentImageId = null;
                            currentRid = null;
                        }
                    }
                }

                reader.close();
            }

        } catch (Exception e) {
        }
    }

    /**
     * 从DISPIMG图片创建CellExtra
     */
    private static void createCellExtraFromDispimgImage(XlsxReadContext xlsxReadContext, PackagePart imagePart,
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
    private static String determinePictureFormat(String contentType, String partName) {
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

    /**
     * 获取相关的Part
     */
    private static PackagePart getRelatedPart(OPCPackage pkg, PackagePart from, PackageRelationship rel) {
        try {
            java.net.URI resolved = PackagingURIHelper.resolvePartUri(
                    from.getPartName().getURI(), rel.getTargetURI());
            PackagePartName target = PackagingURIHelper.createPartName(resolved);
            return pkg.getPart(target);
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * 清理缓存（通常在工作簿关闭时调用）
     */
    public static void clearCache(OPCPackage pkg) {
        if (pkg != null) {
            PACKAGE_IMAGE_CACHE.remove(pkg);
        }
    }

    /**
     * 清理所有缓存
     */
    public static void clearAllCache() {
        PACKAGE_IMAGE_CACHE.clear();
    }
}
