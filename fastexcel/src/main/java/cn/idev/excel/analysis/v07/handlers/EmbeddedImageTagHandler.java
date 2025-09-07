package cn.idev.excel.analysis.v07.handlers;

import cn.idev.excel.context.xlsx.XlsxReadContext;
import cn.idev.excel.enums.CellExtraTypeEnum;
import cn.idev.excel.metadata.CellExtra;
import cn.idev.excel.read.metadata.holder.xlsx.XlsxReadSheetHolder;
import cn.idev.excel.util.StringUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackageRelationshipCollection;
import org.xml.sax.Attributes;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.InputStream;

import org.w3c.dom.Document;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

/**
 * 处理标准 OOXML DrawingML 图片：<drawing r:id="…"> → drawing#.xml → a:blip@r:embed → drawing#.rels → /xl/media/imageX.ext
 */
@Slf4j
public class EmbeddedImageTagHandler extends AbstractXlsxTagHandler {

    @Override
    public void startElement(XlsxReadContext context, String name, Attributes attributes) {
        // 仅处理 <drawing r:id="...">
        String drawingRid = attributes.getValue("r:id");
        if (StringUtils.isEmpty(drawingRid)) {
            drawingRid = attributes.getValue("rId");
        }
        if (StringUtils.isEmpty(drawingRid)) {
            return;
        }

        try {
            PackageRelationshipCollection sheetRels = context.xlsxReadSheetHolder().getPackageRelationshipCollection();
            if (sheetRels == null) {
                return;
            }
            PackageRelationship drawingRel = sheetRels.getRelationshipByID(drawingRid);
            if (drawingRel == null) {
                return;
            }
            PackagePart drawingPart = context.xlsxReadSheetHolder().getPackagePart().getRelatedPart(drawingRel);
            if (drawingPart == null) {
                return;
            }

            // 解析 drawing#.xml，找 a:blip@r:embed
            String embedRid = extractEmbedRidFromDrawing(drawingPart);
            if (StringUtils.isEmpty(embedRid)) {
                return;
            }

            // 用 drawing 的 .rels 定位图片部件
            PackageRelationshipCollection drawingRels = drawingPart.getRelationships();
            if (drawingRels == null) {
                return;
            }
            PackageRelationship picRel = drawingRels.getRelationshipByID(embedRid);
            if (picRel == null) {
                return;
            }
            PackagePart picturePart = drawingPart.getRelatedPart(picRel);
            if (picturePart == null) {
                return;
            }

            // 读取图片为字节数组
            byte[] pictureData = readAllBytes(picturePart);
            String format = determinePictureFormat(picturePart.getContentType(),
                    picturePart.getPartName() == null ? null : picturePart.getPartName().getName());

            XlsxReadSheetHolder sheetHolder = context.xlsxReadSheetHolder();
            CellExtra cellExtra = new CellExtra(
                    CellExtraTypeEnum.MERGE_IMAGE,
                    pictureData,
                    format,
                    sheetHolder.getRowIndex(),
                    sheetHolder.getColumnIndex()
            );
            context.readSheetHolder().setCellExtra(cellExtra);
            context.analysisEventProcessor().extra(context);
        } catch (Exception e) {
            log.debug("处理drawing图片失败: {}", e.getMessage());
        }
    }

    private String extractEmbedRidFromDrawing(PackagePart drawingPart) {
        try (InputStream is = drawingPart.getInputStream()) {
            DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
            dbf.setNamespaceAware(true);
            DocumentBuilder db = dbf.newDocumentBuilder();
            Document doc = db.parse(is);
            NodeList all = doc.getElementsByTagName("*");
            for (int i = 0; i < all.getLength(); i++) {
                Node node = all.item(i);
                if (!"blip".equalsIgnoreCase(localName(node))) {
                    continue;
                }
                NamedNodeMap attrs = node.getAttributes();
                if (attrs == null) {
                    continue;
                }
                String embed = getAttr(attrs, "r:embed");
                if (embed == null) {
                    embed = getAttrEndsWith(attrs, ":embed");
                    if (embed == null) {
                        embed = getAttr(attrs, "embed");
                    }
                }
                if (!StringUtils.isEmpty(embed)) {
                    return embed;
                }
            }
        } catch (Exception e) {
            log.debug("解析 drawing.xml 获取 embed 失败: {}", e.getMessage());
        }
        return null;
    }

    private byte[] readAllBytes(PackagePart picturePart) throws Exception {
        try (InputStream in = picturePart.getInputStream()) {
            java.io.ByteArrayOutputStream buffer = new java.io.ByteArrayOutputStream();
            byte[] data = new byte[4096];
            int n;
            while ((n = in.read(data)) != -1) {
                buffer.write(data, 0, n);
            }
            return buffer.toByteArray();
        }
    }

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
            int dot = partName.lastIndexOf('.');
            if (dot > 0 && dot < partName.length() - 1) {
                return partName.substring(dot + 1).toLowerCase();
            }
        }
        return "bin";
    }

    private String localName(Node node) {
        if (node == null) return null;
        String ln = node.getLocalName();
        if (ln != null) return ln;
        String nm = node.getNodeName();
        if (nm == null) return null;
        int idx = nm.indexOf(':');
        return idx >= 0 ? nm.substring(idx + 1) : nm;
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
}
