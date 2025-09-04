package cn.idev.excel.constant;

/**
 *
 */
public class ExcelXmlConstants {
    public static final String DIMENSION_TAG = "dimension";
    public static final String ROW_TAG = "row";
    public static final String CELL_FORMULA_TAG = "f";
    public static final String CELL_VALUE_TAG = "v";
    /**
     * When the data is "inlineStr" his tag is "t"
     */
    public static final String CELL_INLINE_STRING_VALUE_TAG = "t";

    public static final String CELL_TAG = "c";
    public static final String MERGE_CELL_TAG = "mergeCell";
    public static final String HYPERLINK_TAG = "hyperlink";

    public static final String X_DIMENSION_TAG = "x:dimension";
    public static final String NS2_DIMENSION_TAG = "ns2:dimension";

    public static final String X_ROW_TAG = "x:row";
    public static final String NS2_ROW_TAG = "ns2:row";

    public static final String X_CELL_FORMULA_TAG = "x:f";
    public static final String NS2_CELL_FORMULA_TAG = "ns2:f";
    public static final String X_CELL_VALUE_TAG = "x:v";
    public static final String NS2_CELL_VALUE_TAG = "ns2:v";

    /**
     * When the data is "inlineStr" his tag is "t"
     */
    public static final String X_CELL_INLINE_STRING_VALUE_TAG = "x:t";

    public static final String NS2_CELL_INLINE_STRING_VALUE_TAG = "ns2:t";

    public static final String X_CELL_TAG = "x:c";
    public static final String NS2_CELL_TAG = "ns2:c";
    public static final String X_MERGE_CELL_TAG = "x:mergeCell";
    public static final String NS2_MERGE_CELL_TAG = "ns2:mergeCell";
    public static final String X_HYPERLINK_TAG = "x:hyperlink";
    public static final String NS2_HYPERLINK_TAG = "ns2:hyperlink";

    public static final String IMAGE_TAG = "pic";
    public static final String X_IMAGE_TAG = "x:pic";
    public static final String NS2_IMAGE_TAG = "ns2:pic";
    
    // 图片相关的额外标签
    public static final String DRAWING_TAG = "drawing";
    public static final String X_DRAWING_TAG = "x:drawing";
    public static final String NS2_DRAWING_TAG = "ns2:drawing";
    
    public static final String TWO_CELL_ANCHOR_TAG = "twoCellAnchor";
    public static final String X_TWO_CELL_ANCHOR_TAG = "x:twoCellAnchor";
    public static final String NS2_TWO_CELL_ANCHOR_TAG = "ns2:twoCellAnchor";
    
    public static final String ONE_CELL_ANCHOR_TAG = "oneCellAnchor";
    public static final String X_ONE_CELL_ANCHOR_TAG = "x:oneCellAnchor";
    public static final String NS2_ONE_CELL_ANCHOR_TAG = "ns2:oneCellAnchor";
    
    public static final String BLIP_TAG = "blip";
    public static final String X_BLIP_TAG = "x:blip";
    public static final String NS2_BLIP_TAG = "ns2:blip";
    public static final String A_BLIP_TAG = "a:blip";
    
    public static final String FROM_TAG = "from";
    public static final String X_FROM_TAG = "x:from";
    public static final String NS2_FROM_TAG = "ns2:from";
    
    public static final String TO_TAG = "to";
    public static final String X_TO_TAG = "x:to";
    public static final String NS2_TO_TAG = "ns2:to";
    
    public static final String ROW_OFF_TAG = "row";
    public static final String COL_OFF_TAG = "col";
    public static final String X_ROW_OFF_TAG = "x:row";
    public static final String X_COL_OFF_TAG = "x:col";
    public static final String NS2_ROW_OFF_TAG = "ns2:row";
    public static final String NS2_COL_OFF_TAG = "ns2:col";

    /**
     * s attribute
     */
    public static final String ATTRIBUTE_S = "s";
    /**
     * ref attribute
     */
    public static final String ATTRIBUTE_REF = "ref";
    /**
     * r attribute
     */
    public static final String ATTRIBUTE_R = "r";
    /**
     * t attribute
     */
    public static final String ATTRIBUTE_T = "t";
    /**
     * location attribute
     */
    public static final String ATTRIBUTE_LOCATION = "location";

    /**
     * rId attribute
     */
    public static final String ATTRIBUTE_RID = "r:id";

    /**
     * Cell range split
     */
    public static final String CELL_RANGE_SPLIT = ":";

    // The following is a constant read the `SharedStrings.xml`

    /**
     * text
     */
    public static final String SHAREDSTRINGS_T_TAG = "t";

    public static final String SHAREDSTRINGS_X_T_TAG = "x:t";
    public static final String SHAREDSTRINGS_NS2_T_TAG = "ns2:t";

    /**
     * SharedStringItem
     */
    public static final String SHAREDSTRINGS_SI_TAG = "si";

    public static final String SHAREDSTRINGS_X_SI_TAG = "x:si";
    public static final String SHAREDSTRINGS_NS2_SI_TAG = "ns2:si";

    /**
     * Mac 2016 2017 will have this extra field to ignore
     */
    public static final String SHAREDSTRINGS_RPH_TAG = "rPh";

    public static final String SHAREDSTRINGS_X_RPH_TAG = "x:rPh";
    public static final String SHAREDSTRINGS_NS2_RPH_TAG = "ns2:rPh";
}
