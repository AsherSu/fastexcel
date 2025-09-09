package cn.idev.excel.analysis.v07;

import cn.idev.excel.analysis.ExcelReadExecutor;
import cn.idev.excel.analysis.v07.handlers.sax.SharedStringsTableHandler;
import cn.idev.excel.analysis.v07.handlers.sax.XlsxRowHandler;
import cn.idev.excel.read.metadata.holder.xlsx.XlsxReadSheetHolder;
import cn.idev.excel.util.AdvancedImageUtil;
import cn.idev.excel.cache.ReadCache;
import cn.idev.excel.context.xlsx.XlsxReadContext;
import cn.idev.excel.enums.CellExtraTypeEnum;
import cn.idev.excel.exception.ExcelAnalysisException;
import cn.idev.excel.exception.ExcelAnalysisStopSheetException;
import cn.idev.excel.metadata.CellExtra;
import cn.idev.excel.read.metadata.ReadSheet;
import cn.idev.excel.read.metadata.holder.xlsx.XlsxReadWorkbookHolder;
import cn.idev.excel.util.FileUtils;
import cn.idev.excel.util.MapUtils;
import cn.idev.excel.util.SheetUtils;
import cn.idev.excel.util.StringUtils;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.UUID;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.Comments;
import org.apache.poi.xssf.model.CommentsTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTOneCellAnchor;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTPicture;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTTwoCellAnchor;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbookPr;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STSheetState;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.WorkbookDocument;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

/**
 *
 */
@Slf4j
public class XlsxSaxAnalyser implements ExcelReadExecutor {

    /**
     * Storage sheet SharedStrings
     */
    public static final PackagePartName SHARED_STRINGS_PART_NAME;

    static {
        try {
            SHARED_STRINGS_PART_NAME = PackagingURIHelper.createPartName("/xl/sharedStrings.xml");
        } catch (InvalidFormatException e) {
            log.error("Initialize the XlsxSaxAnalyser failure", e);
            throw new ExcelAnalysisException("Initialize the XlsxSaxAnalyser failure", e);
        }
    }

    private final XlsxReadContext xlsxReadContext;
    private final List<ReadSheet> sheetList;
    private final Map<Integer, InputStream> sheetMap;
    private final Map<String, CTSheet> ctSheetMap;
    /**
     * excel comments key: sheetNo value: CommentsTable
     */
    private final Map<Integer, CommentsTable> commentsTableMap;

    public XlsxSaxAnalyser(XlsxReadContext xlsxReadContext, InputStream decryptedStream) throws Exception {
        this.xlsxReadContext = xlsxReadContext;
        // Initialize cache
        XlsxReadWorkbookHolder xlsxReadWorkbookHolder = xlsxReadContext.xlsxReadWorkbookHolder();

        OPCPackage pkg = readOpcPackage(xlsxReadWorkbookHolder, decryptedStream);
        xlsxReadWorkbookHolder.setOpcPackage(pkg);

        // Read the Shared information Strings
        PackagePart sharedStringsTablePackagePart = pkg.getPart(SHARED_STRINGS_PART_NAME);
        if (sharedStringsTablePackagePart != null) {
            // Specify default cache
            defaultReadCache(xlsxReadWorkbookHolder, sharedStringsTablePackagePart);

            // Analysis sharedStringsTable.xml
            analysisSharedStringsTable(sharedStringsTablePackagePart.getInputStream(), xlsxReadWorkbookHolder);
        }

        XSSFReader xssfReader = new XSSFReader(pkg);
        analysisUse1904WindowDate(xssfReader, xlsxReadWorkbookHolder);
        // set style table
        setStylesTable(xlsxReadWorkbookHolder, xssfReader);

        sheetList = new ArrayList<>();
        sheetMap = new HashMap<>();
        commentsTableMap = new HashMap<>();
        ctSheetMap = new HashMap<>();
        Map<Integer, PackageRelationshipCollection> packageRelationshipCollectionMap = MapUtils.newHashMap();
        xlsxReadWorkbookHolder.setPackageRelationshipCollectionMap(packageRelationshipCollectionMap);
        Map<Integer, PackagePart> packagePartMap = MapUtils.newHashMap();
        xlsxReadWorkbookHolder.setPackagePartMap(packagePartMap);
        // analysis CTSheet
        analysisCtSheetMap(xssfReader, xlsxReadWorkbookHolder);

        XSSFReader.SheetIterator ite = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
        int index = 0;
        if (!ite.hasNext()) {
            throw new ExcelAnalysisException("Can not find any sheet!");
        }
        while (ite.hasNext()) {
            InputStream inputStream = ite.next();
            String sheetName = ite.getSheetName();
            CTSheet ctSheet = ctSheetMap.get(sheetName);
            if (ctSheet == null) {
                continue;
            }
            ReadSheet readSheet = new ReadSheet(index, sheetName);
            readSheet.setHidden(ctSheet.getState() == STSheetState.HIDDEN);
            readSheet.setVeryHidden(ctSheet.getState() == STSheetState.VERY_HIDDEN);
            sheetList.add(readSheet);
            sheetMap.put(index, inputStream);
            if (xlsxReadContext.readWorkbookHolder().getExtraReadSet().contains(CellExtraTypeEnum.COMMENT)) {
                Comments comments = ite.getSheetComments();
                if (comments instanceof CommentsTable) {
                    commentsTableMap.put(index, (CommentsTable) comments);
                }
            }
            if (xlsxReadContext.readWorkbookHolder().getExtraReadSet().contains(CellExtraTypeEnum.HYPERLINK)) {
                PackageRelationshipCollection packageRelationshipCollection = Optional.ofNullable(ite.getSheetPart())
                        .map(packagePart -> {
                            try {
                                return packagePart.getRelationships();
                            } catch (InvalidFormatException e) {
                                log.warn("Reading the Relationship failed", e);
                                return null;
                            }
                        })
                        .orElse(null);
                if (packageRelationshipCollection != null) {
                    packageRelationshipCollectionMap.put(index, packageRelationshipCollection);
                }
            }
            if (xlsxReadContext.readWorkbookHolder().getExtraReadSet().contains(CellExtraTypeEnum.MERGE_IMAGE)) {
                PackageRelationshipCollection packageRelationshipCollection = Optional.ofNullable(ite.getSheetPart())
                        .map(packagePart -> {
                            try {
                                return packagePart.getRelationships();
                            } catch (InvalidFormatException e) {
                                log.warn("Reading the Relationship failed", e);
                                return null;
                            }
                        })
                        .orElse(null);
                PackagePart packagePart = ite.getSheetPart();
                if (packageRelationshipCollection != null) {
                    packageRelationshipCollectionMap.put(index, packageRelationshipCollection);
                }
                if (packagePart != null) {
                    packagePartMap.put(index, packagePart);
                }
            }
            index++;
        }
    }

    private void setStylesTable(XlsxReadWorkbookHolder xlsxReadWorkbookHolder, XSSFReader xssfReader) {
        try {
            xlsxReadWorkbookHolder.setStylesTable(xssfReader.getStylesTable());
        } catch (Exception e) {
            log.warn(
                    "Currently excel cannot get style information, but it doesn't affect the data analysis.You can try to"
                            + " save the file with office again or ignore the current error.",
                    e);
        }
    }

    private void defaultReadCache(
            XlsxReadWorkbookHolder xlsxReadWorkbookHolder, PackagePart sharedStringsTablePackagePart) {
        ReadCache readCache = xlsxReadWorkbookHolder.getReadCacheSelector().readCache(sharedStringsTablePackagePart);
        xlsxReadWorkbookHolder.setReadCache(readCache);
        readCache.init(xlsxReadContext);
    }

    private void analysisUse1904WindowDate(XSSFReader xssfReader, XlsxReadWorkbookHolder xlsxReadWorkbookHolder)
            throws Exception {
        if (xlsxReadWorkbookHolder.getReadWorkbook().getUse1904windowing() != null) {
            return;
        }
        InputStream workbookXml = xssfReader.getWorkbookData();
        WorkbookDocument ctWorkbook = WorkbookDocument.Factory.parse(workbookXml);
        CTWorkbook wb = ctWorkbook.getWorkbook();
        CTWorkbookPr prefix = wb.getWorkbookPr();
        if (prefix != null && prefix.getDate1904()) {
            xlsxReadWorkbookHolder.getGlobalConfiguration().setUse1904windowing(Boolean.TRUE);
        } else {
            xlsxReadWorkbookHolder.getGlobalConfiguration().setUse1904windowing(Boolean.FALSE);
        }
    }

    private void analysisSharedStringsTable(
            InputStream sharedStringsTableInputStream, XlsxReadWorkbookHolder xlsxReadWorkbookHolder) {
        ContentHandler handler = new SharedStringsTableHandler(xlsxReadWorkbookHolder.getReadCache());
        parseXmlSource(sharedStringsTableInputStream, handler);
        xlsxReadWorkbookHolder.getReadCache().putFinished();
    }

    private void analysisCtSheetMap(XSSFReader xssfReader, XlsxReadWorkbookHolder xlsxReadWorkbookHolder)
            throws Exception {
        CTWorkbook wb =
                WorkbookDocument.Factory.parse(xssfReader.getWorkbookData()).getWorkbook();
        for (CTSheet ctSheet : wb.getSheets().getSheetList()) {
            boolean isHidden =
                    (ctSheet.getState() == STSheetState.HIDDEN) || (ctSheet.getState() == STSheetState.VERY_HIDDEN);
            if (Boolean.FALSE.equals(xlsxReadWorkbookHolder.getIgnoreHiddenSheet()) || !isHidden) {
                ctSheetMap.put(ctSheet.getName(), ctSheet);
            }
        }
    }

    private OPCPackage readOpcPackage(XlsxReadWorkbookHolder xlsxReadWorkbookHolder, InputStream decryptedStream)
            throws Exception {
        if (decryptedStream == null && xlsxReadWorkbookHolder.getFile() != null) {
            return OPCPackage.open(xlsxReadWorkbookHolder.getFile());
        }
        if (xlsxReadWorkbookHolder.getMandatoryUseInputStream()) {
            if (decryptedStream != null) {
                return OPCPackage.open(decryptedStream);
            } else {
                return OPCPackage.open(xlsxReadWorkbookHolder.getInputStream());
            }
        }
        File readTempFile = FileUtils.createCacheTmpFile();
        xlsxReadWorkbookHolder.setTempFile(readTempFile);
        File tempFile = new File(readTempFile.getPath(), UUID.randomUUID() + ".xlsx");
        if (decryptedStream != null) {
            FileUtils.writeToFile(tempFile, decryptedStream, false);
        } else {
            FileUtils.writeToFile(
                    tempFile, xlsxReadWorkbookHolder.getInputStream(), xlsxReadWorkbookHolder.getAutoCloseStream());
        }
        return OPCPackage.open(tempFile, PackageAccess.READ);
    }

    @Override
    public List<ReadSheet> sheetList() {
        return sheetList;
    }

    private void parseXmlSource(InputStream inputStream, ContentHandler handler) {
        InputSource inputSource = new InputSource(inputStream);
        try {
            SAXParserFactory saxFactory;
            String xlsxSAXParserFactoryName =
                    xlsxReadContext.xlsxReadWorkbookHolder().getSaxParserFactoryName();
            if (StringUtils.isEmpty(xlsxSAXParserFactoryName)) {
                saxFactory = SAXParserFactory.newInstance();
            } else {
                saxFactory = SAXParserFactory.newInstance(xlsxSAXParserFactoryName, null);
            }
            try {
                saxFactory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
            } catch (Throwable ignore) {
            }
            try {
                saxFactory.setFeature("http://xml.org/sax/features/external-general-entities", false);
            } catch (Throwable ignore) {
            }
            try {
                saxFactory.setFeature("http://xml.org/sax/features/external-parameter-entities", false);
            } catch (Throwable ignore) {
            }
            SAXParser saxParser = saxFactory.newSAXParser();
            XMLReader xmlReader = saxParser.getXMLReader();
            xmlReader.setContentHandler(handler);
            xmlReader.parse(inputSource);
            inputStream.close();
        } catch (IOException | ParserConfigurationException | SAXException e) {
            throw new ExcelAnalysisException(e);
        } finally {
            if (inputStream != null) {
                try {
                    inputStream.close();
                } catch (IOException e) {
                    throw new ExcelAnalysisException("Can not close 'inputStream'!");
                }
            }
        }
    }

    @Override
    public void execute() {
        for (ReadSheet readSheet : sheetList) {
            readSheet = SheetUtils.match(readSheet, xlsxReadContext);
            if (readSheet != null) {
                try {
                    xlsxReadContext.currentSheet(readSheet);

                    XlsxReadSheetHolder xlsxReadSheetHolder = xlsxReadContext.xlsxReadSheetHolder();
                    PackagePart packagePart = xlsxReadSheetHolder.getPackagePart();
                    List<EmbeddedImage> images = extractEmbeddedImagesFromSheetPart(packagePart);
                    Map<Integer, Map<Integer, EmbeddedImage>> imageMap = buildImageMap(images);
                    xlsxReadSheetHolder.setImageMap(imageMap);

                    parseXmlSource(sheetMap.get(readSheet.getSheetNo()), new XlsxRowHandler(xlsxReadContext));
                    // Read comments
                    readComments(readSheet);
                    // Process advanced images (CellImages and DISPIMG)
                    if (xlsxReadContext.readWorkbookHolder().getExtraReadSet().contains(CellExtraTypeEnum.MERGE_IMAGE)) {
                        processAdvancedImages(readSheet);
                    }
                } catch (ExcelAnalysisStopSheetException e) {
                    if (log.isDebugEnabled()) {
                        log.debug("Custom stop!", e);
                    }
                } catch (XmlException e) {
                    throw new RuntimeException(e);
                } catch (IOException e) {
                    throw new RuntimeException(e);
                } catch (InvalidFormatException e) {
                    throw new RuntimeException(e);
                }
                // The last sheet is read
                xlsxReadContext.analysisEventProcessor().endSheet(xlsxReadContext);
            }
        }
    }

    // Java
    private static final String REL_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
    private static final String REL_IMAGE   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

    public static class EmbeddedImage {
        public final String rId;
        public final String contentType;
        public final String fileName;
        public final byte[] data;
        public final int row;   // -1 表示未解析到位置
        public final int col;   // -1 表示未解析到位置

        public EmbeddedImage(String rId, String contentType, String fileName, byte[] data, int row, int col) {
            this.rId = rId;
            this.contentType = contentType;
            this.fileName = fileName;
            this.data = data;
            this.row = row;
            this.col = col;
        }
    }

    /**
     * 通过 Sheet 的 PackagePart 提取所有内嵌图片（含起始单元格位置）
     */
    public static List<EmbeddedImage> extractEmbeddedImagesFromSheetPart(org.apache.poi.openxml4j.opc.PackagePart sheetPart)
            throws java.io.IOException, org.apache.poi.openxml4j.exceptions.InvalidFormatException, org.apache.xmlbeans.XmlException {
        List<EmbeddedImage> result = new java.util.ArrayList<>();
        if (sheetPart == null) return result;

        org.apache.poi.openxml4j.opc.PackageRelationshipCollection drawings =
                sheetPart.getRelationshipsByType(REL_DRAWING);
        if (drawings == null) return result;

        for (org.apache.poi.openxml4j.opc.PackageRelationship drRel : drawings) {
            org.apache.poi.openxml4j.opc.PackagePart drawingPart = sheetPart.getRelatedPart(drRel);

            // 解析 drawing.xml，建立 rId -> 单元格位置 的映射
            java.util.Map<String, org.apache.poi.ss.util.CellAddress> posMap = parseDrawingAnchors(drawingPart);

            // 遍历图片关系并读取二进制
            org.apache.poi.openxml4j.opc.PackageRelationshipCollection images =
                    drawingPart.getRelationshipsByType(REL_IMAGE);
            if (images == null) continue;

            for (org.apache.poi.openxml4j.opc.PackageRelationship imgRel : images) {
                org.apache.poi.openxml4j.opc.PackagePart imgPart = drawingPart.getRelatedPart(imgRel);
                byte[] bytes;
                try (java.io.InputStream is = imgPart.getInputStream()) {
                    bytes = org.apache.poi.util.IOUtils.toByteArray(is);
                }
                String fileName = fileNameOfPart(imgPart);
                String contentType = imgPart.getContentType();

                org.apache.poi.ss.util.CellAddress addr = posMap.get(imgRel.getId());
                int row = addr != null ? addr.getRow() : -1;
                int col = addr != null ? addr.getColumn() : -1;

                result.add(new EmbeddedImage(imgRel.getId(), contentType, fileName, bytes, row, col));
            }
        }
        return result;
    }

    /**
     * 解析 \`drawing.xml\`，提取图片 rId 对应的锚点起始单元格（row、col）
     */
    private static java.util.Map<String, org.apache.poi.ss.util.CellAddress> parseDrawingAnchors(
            org.apache.poi.openxml4j.opc.PackagePart drawingPart)
            throws java.io.IOException, org.apache.xmlbeans.XmlException {
        java.util.Map<String, org.apache.poi.ss.util.CellAddress> map = new java.util.HashMap<>();
        if (drawingPart == null) return map;

        try (java.io.InputStream is = drawingPart.getInputStream()) {
            org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTDrawing ct =
                    org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTDrawing.Factory.parse(is);

            // twoCellAnchor// todo 解析不了
            for (CTTwoCellAnchor a : ct.getTwoCellAnchorList()) {
                CTPicture pic = a.getPic();
                if (pic != null && pic.getBlipFill() != null && pic.getBlipFill().getBlip() != null) {
                    String rId = pic.getBlipFill().getBlip().getEmbed();
                    CTMarker from = a.getFrom();
                    if (rId != null && from != null) {
                        map.put(rId, new CellAddress(from.getRow(), from.getCol()));
                    }
                }
            }
            // oneCellAnchor
            for (CTOneCellAnchor a : ct.getOneCellAnchorList()) {
                CTPicture pic = a.getPic();
                if (pic != null && pic.getBlipFill() != null && pic.getBlipFill().getBlip() != null) {
                    String rId = pic.getBlipFill().getBlip().getEmbed();
                    CTMarker from = a.getFrom();
                    if (rId != null && from != null) {
                        map.put(rId, new CellAddress(from.getRow(), from.getCol()));
                    }
                }
            }
        }
        return map;
    }

    /**
     * 将图片列表转换为按行列号组织的Map结构
     * @param images 图片列表
     * @return Map<行号, Map<列号, EmbeddedImage>>
     */
    private static Map<Integer, Map<Integer, EmbeddedImage>> buildImageMap(List<EmbeddedImage> images) {
        Map<Integer, Map<Integer, EmbeddedImage>> imageMap = new HashMap<>();
        if (images == null || images.isEmpty()) {
            return imageMap;
        }

        for (EmbeddedImage image : images) {
            if (image.row >= 0 && image.col >= 0) { // 只处理有效位置的图片
                imageMap.computeIfAbsent(image.row, k -> new HashMap<>()).put(image.col, image);
            }
        }

        return imageMap;
    }

    private static String fileNameOfPart(org.apache.poi.openxml4j.opc.PackagePart part) {
        String path = part.getPartName().getName(); // 形如 /xl/media/image1.png
        int idx = path.lastIndexOf('/');
        return idx >= 0 ? path.substring(idx + 1) : path;
    }

    private void readComments(ReadSheet readSheet) {
        if (!xlsxReadContext.readWorkbookHolder().getExtraReadSet().contains(CellExtraTypeEnum.COMMENT)) {
            return;
        }
        CommentsTable commentsTable = commentsTableMap.get(readSheet.getSheetNo());
        if (commentsTable == null) {
            return;
        }
        Iterator<CellAddress> cellAddresses = commentsTable.getCellAddresses();
        while (cellAddresses.hasNext()) {
            CellAddress cellAddress = cellAddresses.next();
            XSSFComment cellComment = commentsTable.findCellComment(cellAddress);
            CellExtra cellExtra = new CellExtra(
                    CellExtraTypeEnum.COMMENT,
                    cellComment.getString().toString(),
                    cellAddress.getRow(),
                    cellAddress.getColumn());
            xlsxReadContext.readSheetHolder().setCellExtra(cellExtra);
            xlsxReadContext.analysisEventProcessor().extra(xlsxReadContext);
        }
    }

    /**
     * 处理高级图片类型（CellImages和DISPIMG公式图片）
     * Process advanced image types (CellImages and DISPIMG formula pictures)
     */
    private void processAdvancedImages(ReadSheet readSheet) {
        try {
            log.debug("开始处理Sheet[{}]的高级图片", readSheet.getSheetName());
            AdvancedImageUtil.processAdvancedImages(xlsxReadContext);
        } catch (Exception e) {
            log.warn("处理Sheet[{}]的高级图片时发生错误: {}", readSheet.getSheetName(), e.getMessage());
            // 不中断正常的读取流程，只记录警告
        }
    }
}
