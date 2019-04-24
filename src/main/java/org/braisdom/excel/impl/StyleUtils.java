package org.braisdom.excel.impl;

import com.helger.css.ECSSVersion;
import com.helger.css.decl.CSSDeclaration;
import com.helger.css.decl.CSSDeclarationList;
import com.helger.css.decl.shorthand.CSSShortHandDescriptor;
import com.helger.css.decl.shorthand.CSSShortHandRegistry;
import com.helger.css.property.CCSSProperties;
import com.helger.css.property.ECSSProperty;
import com.helger.css.propertyvalue.CCSSValue;
import com.helger.css.reader.CSSReaderDeclarationList;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.awt.*;
import java.util.List;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class StyleUtils {

    private static final Pattern SIZE_PATTERN = Pattern.compile("(\\d+)\\w*");
    private static final Stack<Integer> HSSF_COLORS = new Stack();

    static {
        Set<Integer> colorIndexSet = HSSFColor.getIndexHash().keySet();
        HSSF_COLORS.addAll(colorIndexSet);

        CCSSProperties.BORDER_STYLE.newValue("thin");
    }

    public static Map<String, List<CSSDeclaration>> getStyle(String rawStyle) {
        Map<String, List<CSSDeclaration>> declarationMap = new HashMap<>();
        CSSDeclarationList cssDeclarations = CSSReaderDeclarationList.readFromString(rawStyle,
                ECSSVersion.CSS30);
        List<CSSDeclaration> declarations = cssDeclarations.getAllDeclarations();

        for (CSSDeclaration declaration : declarations) {
            ECSSProperty ecssProperty = ECSSProperty.getFromNameOrNull(declaration.getProperty());
            if (CSSShortHandRegistry.isShortHandProperty(ecssProperty)) {
                CSSShortHandDescriptor descriptor = CSSShortHandRegistry.getShortHandDescriptor(ecssProperty);
                declarationMap.put(declaration.getProperty(), new CSSShortHandDescriptorWrapper(descriptor)
                        .getSplitIntoPieces(declaration));
            } else
                declarationMap.put(declaration.getProperty(), Collections.singletonList(declaration));
        }

        return declarationMap;
    }

    public static void setBackgroundColorStyle(HSSFPalette palette, HSSFCellStyle hssfCellStyle, List<CSSDeclaration> cssDeclarations) {
        if (hasCssDeclarations(cssDeclarations)) {
            CSSDeclaration cssDeclaration = cssDeclarations.get(0);
            hssfCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            hssfCellStyle.setFillForegroundColor(getHssfColorIndex(palette,
                    cssDeclaration.getExpressionAsCSSString()));
        }
    }

    public static void setTextColor(HSSFPalette palette, HSSFFont hssfFont, List<CSSDeclaration> cssDeclarations) {
        if (hasCssDeclarations(cssDeclarations)) {
            CSSDeclaration cssDeclaration = cssDeclarations.get(0);
            hssfFont.setColor(getHssfColorIndex(palette, cssDeclaration.getExpressionAsCSSString()));
        }
    }

    public static void setHorizontalAlignment(HSSFCellStyle hssfCellStyle, List<CSSDeclaration> cssDeclarations) {
        if (hasCssDeclarations(cssDeclarations)) {
            CSSDeclaration cssDeclaration = cssDeclarations.get(0);
            String expression = cssDeclaration.getExpressionAsCSSString();
            hssfCellStyle.setAlignment(HorizontalAlignment.valueOf(expression.toUpperCase()));
        }
    }

    public static void setVerticalAlignment(HSSFCellStyle hssfCellStyle, List<CSSDeclaration> cssDeclarations) {
        if (hasCssDeclarations(cssDeclarations)) {
            CSSDeclaration cssDeclaration = cssDeclarations.get(0);
            String expression = cssDeclaration.getExpressionAsCSSString();
            hssfCellStyle.setVerticalAlignment(VerticalAlignment.valueOf(expression.toUpperCase()));
        }
    }

    public static void setBorder(HSSFPalette palette, HSSFCellStyle hssfCellStyle, List<CSSDeclaration> cssDeclarations) {
        if (hasCssDeclarations(cssDeclarations)) {
            Map<String, CSSDeclaration> cssDeclarationMap = cssDeclarations.stream().collect(HashMap::new, (m, v) ->
                    m.put(v.getProperty(), v), HashMap::putAll);
            CSSDeclaration borderStyleDecl = cssDeclarationMap.get(ECSSProperty.BORDER_STYLE.getName());
            CSSDeclaration borderColorDecl = cssDeclarationMap.get(ECSSProperty.BORDER_COLOR.getName());

            if(borderStyleDecl != null) {
                setLeftBorderStyle(hssfCellStyle, Collections.singletonList(borderStyleDecl));
                setRightBorderStyle(hssfCellStyle, Collections.singletonList(borderStyleDecl));
                setTopBorderStyle(hssfCellStyle, Collections.singletonList(borderStyleDecl));
                setBottomBorderStyle(hssfCellStyle, Collections.singletonList(borderStyleDecl));
            }

            if(borderColorDecl != null) {
                setLeftBorderColor(palette, hssfCellStyle, Collections.singletonList(borderColorDecl));
                setRightBorderColor(palette, hssfCellStyle, Collections.singletonList(borderColorDecl));
                setTopBorderColor(palette, hssfCellStyle, Collections.singletonList(borderColorDecl));
                setBottomBorderColor(palette, hssfCellStyle, Collections.singletonList(borderColorDecl));
            }
        }
    }

    public static void setLeftBorderColor(HSSFPalette palette, HSSFCellStyle hssfCellStyle, List<CSSDeclaration> cssDeclaration) {
        if (hasCssDeclarations(cssDeclaration))
            hssfCellStyle.setLeftBorderColor(getHssfColorIndex(palette, cssDeclaration.get(0).getExpressionAsCSSString()));
    }

    public static void setRightBorderColor(HSSFPalette palette, HSSFCellStyle hssfCellStyle, List<CSSDeclaration> cssDeclaration) {
        if (hasCssDeclarations(cssDeclaration))
            hssfCellStyle.setRightBorderColor(getHssfColorIndex(palette, cssDeclaration.get(0).getExpressionAsCSSString()));
    }

    public static void setTopBorderColor(HSSFPalette palette, HSSFCellStyle hssfCellStyle, List<CSSDeclaration> cssDeclaration) {
        if (hasCssDeclarations(cssDeclaration))
            hssfCellStyle.setTopBorderColor(getHssfColorIndex(palette, cssDeclaration.get(0).getExpressionAsCSSString()));
    }

    public static void setBottomBorderColor(HSSFPalette palette, HSSFCellStyle hssfCellStyle, List<CSSDeclaration> cssDeclaration) {
        if (hasCssDeclarations(cssDeclaration))
            hssfCellStyle.setBottomBorderColor(getHssfColorIndex(palette, cssDeclaration.get(0).getExpressionAsCSSString()));
    }

    public static void setLeftBorderStyle(HSSFCellStyle hssfCellStyle, List<CSSDeclaration> cssDeclaration) {
        if (hasCssDeclarations(cssDeclaration))
            hssfCellStyle.setBorderLeft(BorderStyle.valueOf(cssDeclaration.get(0).getExpressionAsCSSString().toUpperCase()));
    }

    public static void setRightBorderStyle(HSSFCellStyle hssfCellStyle, List<CSSDeclaration> cssDeclaration) {
        if (hasCssDeclarations(cssDeclaration))
            hssfCellStyle.setBorderRight(BorderStyle.valueOf(cssDeclaration.get(0).getExpressionAsCSSString().toUpperCase()));
    }

    public static void setTopBorderStyle(HSSFCellStyle hssfCellStyle, List<CSSDeclaration> cssDeclaration) {
        if (hasCssDeclarations(cssDeclaration))
            hssfCellStyle.setBorderTop(BorderStyle.valueOf(cssDeclaration.get(0).getExpressionAsCSSString().toUpperCase()));
    }

    public static void setBottomBorderStyle(HSSFCellStyle hssfCellStyle, List<CSSDeclaration> cssDeclaration) {
        if (hasCssDeclarations(cssDeclaration))
            hssfCellStyle.setBorderBottom(BorderStyle.valueOf(cssDeclaration.get(0).getExpressionAsCSSString().toUpperCase()));
    }

    public static void setWhiteSpaceStyle(HSSFCellStyle hssfCellStyle, List<CSSDeclaration> cssDeclaration) {
        if (hasCssDeclarations(cssDeclaration)) {
            String value = cssDeclaration.get(0).getExpressionAsCSSString();
            if (CCSSValue.NOWRAP.equals(value))
                hssfCellStyle.setWrapText(false);
            else
                hssfCellStyle.setWrapText(true);
        }
    }

    public static void setFontStyle(HSSFFont hssfFont, List<CSSDeclaration> cssDeclarations) {
        if (hasCssDeclarations(cssDeclarations)) {
            Map<String, CSSDeclaration> cssDeclarationMap = cssDeclarations.stream().collect(HashMap::new, (m, v) ->
                    m.put(v.getProperty(), v), HashMap::putAll);
            CSSDeclaration fontStyleDecl = cssDeclarationMap.get(ECSSProperty.FONT_STYLE.getName());
            CSSDeclaration fontFamilyDecl = cssDeclarationMap.get(ECSSProperty.FONT_FAMILY.getName());
            CSSDeclaration fontWeightDecl = cssDeclarationMap.get(ECSSProperty.FONT_WEIGHT.getName());
            CSSDeclaration fontSizeDecl = cssDeclarationMap.get(ECSSProperty.FONT_SIZE.getName());

            if(fontSizeDecl != null)
                setFontSize(hssfFont, Collections.singletonList(fontSizeDecl));

            if(fontStyleDecl != null)
                setFontStyle2(hssfFont, Collections.singletonList(fontStyleDecl));

            if(fontFamilyDecl != null)
                setFontFamily(hssfFont, Collections.singletonList(fontFamilyDecl));

            if(fontWeightDecl != null)
                setFontWeight(hssfFont, Collections.singletonList(fontWeightDecl));
        }
    }

    public static void setFontStyle2(HSSFFont hssfFont, List<CSSDeclaration> cssDeclarations) {
        if (hasCssDeclarations(cssDeclarations)) {
            String rawFontStyle = cssDeclarations.get(0).getExpressionAsCSSString();
            hssfFont.setItalic(CCSSValue.ITALIC.equals(rawFontStyle));
        }
    }

    public static void setFontWeight(HSSFFont hssfFont, List<CSSDeclaration> cssDeclarations) {
        if (hasCssDeclarations(cssDeclarations)) {
            String rawFontWeight = cssDeclarations.get(0).getExpressionAsCSSString();
            hssfFont.setBold(CCSSValue.BOLD.equals(rawFontWeight));
        }
    }

    public static void setTextDecoration(HSSFFont hssfFont, List<CSSDeclaration> cssDeclarations) {
        if (hasCssDeclarations(cssDeclarations)) {
            String textDecorationWeight = cssDeclarations.get(0).getExpressionAsCSSString();
            if(CCSSValue.UNDERLINE.equals(textDecorationWeight))
                hssfFont.setUnderline(HSSFFont.U_SINGLE);
            else
                hssfFont.setUnderline(HSSFFont.U_NONE);
        }
    }

    public static void setFontFamily(HSSFFont hssfFont, List<CSSDeclaration> cssDeclarations) {
        if (hasCssDeclarations(cssDeclarations)) {
            hssfFont.setFontName(cssDeclarations.get(0).getExpressionAsCSSString());
        }
    }

    public static void setFontSize(HSSFFont hssfFont, List<CSSDeclaration> cssDeclarations) {
        if (hasCssDeclarations(cssDeclarations)) {
            String rawFontSize = cssDeclarations.get(0).getExpressionAsCSSString();
            hssfFont.setFontHeightInPoints(getSizeValue(rawFontSize));
        }
    }

    public static short getHssfColorIndex(HSSFPalette palette, String hexString) {
        Color color = Color.decode(hexString);
        if (HSSF_COLORS.empty())
            throw new IllegalStateException("The template has too many colors");
        short colorIndex = HSSF_COLORS.pop().shortValue();
        palette.setColorAtIndex(colorIndex,
                (byte) color.getRed(), (byte) color.getGreen(), (byte) color.getBlue());
        return colorIndex;
    }

    public static boolean hasCssDeclarations(List<CSSDeclaration> cssDeclarations) {
        return cssDeclarations != null && cssDeclarations.size() > 0;
    }

    public static int safeInteger(String s) {
        if (s == null)
            return 0;
        try {
            return Integer.valueOf(s);
        } catch (NumberFormatException ex) {
            return 0;
        }
    }

    public static short getSizeValue(String rawSize) {
        Matcher matcher = SIZE_PATTERN.matcher(rawSize);
        if (matcher.matches())
            return Short.parseShort(matcher.group(1));
        else
            return -1;
    }
}
