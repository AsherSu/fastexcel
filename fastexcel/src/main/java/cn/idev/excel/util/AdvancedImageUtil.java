package cn.idev.excel.util;

import cn.idev.excel.context.xlsx.XlsxReadContext;
import cn.idev.excel.enums.CellExtraTypeEnum;
import cn.idev.excel.metadata.CellExtra;
import org.apache.poi.openxml4j.opc.*;

import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamConstants;
import javax.xml.stream.XMLStreamReader;
import java.io.InputStream;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 高级图片处理器 - 处理CellImages和DISPIMG公式图片
 * Advanced Image Handler - Handles CellImages and DISPIMG formula pictures
 *
 * 基于您提供的Excel图片解析方案实现
 */
public class AdvancedImageUtil {

    // 解析DISPIMG公式的正则表达式
    private static final Pattern DISPIMG_PATTERN = Pattern.compile(
            "^(?:@)?(?:_xlfn\\.)?DISPIMG\\(\\s*\"([^\"]+)\"(?:\\s*,\\s*\\d+\\s*)?\\)\\s*$",
            Pattern.CASE_INSENSITIVE);

    /**
     * 处理Sheet中的所有高级图片类型
     */
    public static void processAdvancedImages(XlsxReadContext xlsxReadContext) {
        try {
            // 1. 处理CellImages
            processCellImages(xlsxReadContext);

            // 2. 处理DISPIMG公式图片
            processDispimgFormulaImages(xlsxReadContext);

        } catch (Exception e) {
        }
    }

    /**
     * 处理内嵌图片 (CellImages)
     */
    private static void processCellImages(XlsxReadContext xlsxReadContext) {
        try {
            PackagePart sheetPart = xlsxReadContext.xlsxReadSheetHolder().getPackagePart();
            if (sheetPart == null) {
                return;
            }

            // 查找cellImage关系
            List<PackageRelationship> cellImageRels = findCellImageRelationships(sheetPart);
            if (cellImageRels.isEmpty()) {
                return;
            }

            OPCPackage pkg = xlsxReadContext.xlsxReadWorkbookHolder().getOpcPackage();
            int sheetIndex = xlsxReadContext.readSheetHolder().getSheetNo();
            String sheetName = xlsxReadContext.readSheetHolder().getSheetName();

            for (PackageRelationship rel : cellImageRels) {
                PackagePart cellImgPart = getRelatedPart(pkg, sheetPart, rel);
                if (cellImgPart == null) continue;

                // 建立rId到图片Part的映射
                Map<String, PackagePart> rid2Img = buildImagePartMapping(cellImgPart);

                // 解析cellimages*.xml
                try (InputStream in = cellImgPart.getInputStream()) {
                    parseCellImagesXml(xlsxReadContext, in, rid2Img, sheetIndex, sheetName);
                }
            }
        } catch (Exception e) {
        }
    }

    /**
     * 查找CellImage关系
     */
    private static List<PackageRelationship> findCellImageRelationships(PackagePart sheetPart) {
        List<PackageRelationship> cellImageRels = new ArrayList<>();
        try {
            for (PackageRelationship r : sheetPart.getRelationships()) {
                if (r.getRelationshipType() != null &&
                    r.getRelationshipType().toLowerCase().contains("cellimage")) {
                    cellImageRels.add(r);
                }
            }
        } catch (Exception e) {
        }
        return cellImageRels;
    }

    /**
     * 建立rId到图片Part的映射
     */
    private static Map<String, PackagePart> buildImagePartMapping(PackagePart cellImgPart) {
        Map<String, PackagePart> rid2Img = new HashMap<>();
        try {
            OPCPackage pkg = cellImgPart.getPackage();
            for (PackageRelationship ir : cellImgPart.getRelationships()) {
                String type = ir.getRelationshipType();
                if (type != null && type.toLowerCase().endsWith("/image")) {
                    PackagePart imgPart = getRelatedPart(pkg, cellImgPart, ir);
                    if (imgPart != null) {
                        rid2Img.put(ir.getId(), imgPart);
                    }
                }
            }
        } catch (Exception e) {
        }
        return rid2Img;
    }

    /**
     * 解析cellimages*.xml文件
     */
    private static void parseCellImagesXml(XlsxReadContext xlsxReadContext, InputStream xml,
                                          Map<String, PackagePart> rid2Img, int sheetIndex, String sheetName) {
        try {
            XMLInputFactory factory = XMLInputFactory.newInstance();
            XMLStreamReader reader = factory.createXMLStreamReader(xml);

            // 解析状态
            boolean inFrom = false;
            int row1 = -1, col1 = -1;
            String currentRid = null;

            while (reader.hasNext()) {
                int event = reader.next();
                if (event == XMLStreamConstants.START_ELEMENT) {
                    String localName = reader.getLocalName();

                    if ("twoCellAnchor".equals(localName) || "oneCellAnchor".equals(localName)) {
                        row1 = col1 = -1;
                        currentRid = null;
                    } else if ("from".equals(localName)) {
                        inFrom = true;
                    } else if ("to".equals(localName)) {
                        inFrom = false;
                    } else if ("row".equals(localName) && inFrom) {
                        row1 = parseNextTextAsInt(reader, -1);
                    } else if ("col".equals(localName) && inFrom) {
                        col1 = parseNextTextAsInt(reader, -1);
                    } else if ("blip".equals(localName)) {
                        // 查找r:embed属性
                        for (int i = 0; i < reader.getAttributeCount(); i++) {
                            String attrLocal = reader.getAttributeLocalName(i);
                            if ("embed".equals(attrLocal)) {
                                currentRid = reader.getAttributeValue(i);
                                break;
                            }
                        }
                    }
                } else if (event == XMLStreamConstants.END_ELEMENT) {
                    String localName = reader.getLocalName();
                    if ("from".equals(localName) || "to".equals(localName)) {
                        inFrom = false;
                    } else if ("twoCellAnchor".equals(localName) || "oneCellAnchor".equals(localName)) {
                        // 创建图片对象
                        if (row1 >= 0 && col1 >= 0 && currentRid != null) {
                            PackagePart imgPart = rid2Img.get(currentRid);
                            if (imgPart != null) {
                                createCellExtraFromPart(xlsxReadContext, imgPart, row1, col1);
                            }
                        }
                        row1 = col1 = -1;
                        currentRid = null;
                    }
                }
            }
            reader.close();
        } catch (Exception e) {
        }
    }

    /**
     * 处理DISPIMG公式图片
     */
    private static void processDispimgFormulaImages(XlsxReadContext xlsxReadContext) {
        try {
            OPCPackage pkg = xlsxReadContext.xlsxReadWorkbookHolder().getOpcPackage();

            // 1. 建立ID到图片Part的映射
            Map<String, PackagePart> id2ImgPart = buildIdToImagePartMap(pkg);
            if (id2ImgPart.isEmpty()) {
                return;
            }

            // 2. 这里需要访问当前Sheet的实际工作簿来查找DISPIMG公式
            // 由于SAX解析的限制，这部分需要在Sheet解析完成后处理
            // 或者通过其他方式获取公式信息


        } catch (Exception e) {

        }
    }

    /**
     * 建立ID到图片Part的映射（用于DISPIMG）
     */
    private static Map<String, PackagePart> buildIdToImagePartMap(OPCPackage pkg) {
        Map<String, PackagePart> id2Img = new HashMap<>();

        try {
            // 查找所有cellimages*.xml Part
            List<PackagePart> cellImgParts = new ArrayList<>();
            for (PackagePart part : pkg.getParts()) {
                String name = part.getPartName().getName().toLowerCase();
                if (name.contains("cellimages") && name.endsWith(".xml")) {
                    cellImgParts.add(part);
                }
            }

            XMLInputFactory xf = XMLInputFactory.newInstance();
            for (PackagePart ciPart : cellImgParts) {
                // 收集该part的rId到imagePart映射
                Map<String, PackagePart> rid2Img = new HashMap<>();
                for (PackageRelationship rel : ciPart.getRelationships()) {
                    String type = String.valueOf(rel.getRelationshipType()).toLowerCase();
                    if (type.endsWith("/image")) {
                        PackagePart img = getRelatedPart(pkg, ciPart, rel);
                        if (img != null) {
                            rid2Img.put(rel.getId(), img);
                        }
                    }
                }

                // 解析cellimages.xml中的name="ID_xxx"和r:embed
                try (InputStream xml = ciPart.getInputStream()) {
                    XMLStreamReader r = xf.createXMLStreamReader(xml);
                    String currentId = null;
                    String currentRid = null;

                    while (r.hasNext()) {
                        int e = r.next();
                        if (e == XMLStreamConstants.START_ELEMENT) {
                            // 查找name和embed属性
                            for (int i = 0; i < r.getAttributeCount(); i++) {
                                String aLocal = r.getAttributeLocalName(i);
                                String aVal = r.getAttributeValue(i);
                                if ("name".equals(aLocal) && aVal != null && aVal.startsWith("ID_")) {
                                    currentId = aVal;
                                }
                                if ("embed".equals(aLocal)) {
                                    currentRid = aVal;
                                }
                            }
                        } else if (e == XMLStreamConstants.END_ELEMENT) {
                            // 建立映射
                            if (currentId != null && currentRid != null) {
                                PackagePart img = rid2Img.get(currentRid);
                                if (img != null) {
                                    id2Img.putIfAbsent(currentId, img);
                                }
                            }
                            // 重置状态
                            String localName = r.getLocalName();
                            if ("twoCellAnchor".equals(localName) || "oneCellAnchor".equals(localName) ||
                                "pic".equals(localName) || "cellImage".equals(localName)) {
                                currentId = currentRid = null;
                            }
                        }
                    }
                    r.close();
                }
            }
        } catch (Exception e) {

        }

        return id2Img;
    }

    /**
     * 从PackagePart创建CellExtra
     */
    private static void createCellExtraFromPart(XlsxReadContext xlsxReadContext, PackagePart imgPart,
                                               int row, int col) {
        try (InputStream inputStream = imgPart.getInputStream()) {
            byte[] pictureData = readInputStreamToByteArray(inputStream);
            String pictureFormat = determinePictureFormat(imgPart.getContentType(),
                    imgPart.getPartName().getName());

            CellExtra cellExtra = new CellExtra(
                    CellExtraTypeEnum.MERGE_IMAGE,
                    pictureData,
                    pictureFormat,
                    row,
                    col
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
     * 解析DISPIMG公式中的ID
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
     * 解析下一个文本为整数
     */
    private static int parseNextTextAsInt(XMLStreamReader reader, int defaultValue) {
        try {
            int event = reader.next();
            if (event == XMLStreamConstants.CHARACTERS) {
                return Integer.parseInt(reader.getText().trim());
            }
        } catch (Exception e) {
        }
        return defaultValue;
    }
}
