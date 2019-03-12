package util;

import org.apache.poi.sl.usermodel.PaintStyle;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

import java.util.HashMap;
import java.util.Map;

/**
 * created by FingerZhu on 2019/3/7.
 */

public class PlanUtil {

    public static Map<FontStyle, Object> getFontStyle(XSLFTextShape shape) {
        Map<FontStyle, Object> ret = new HashMap();

        XSLFTextRun text = shape.getTextParagraphs().get(0).getTextRuns().get(0);
        ret.put(FontStyle.FontSize, text.getFontSize());
        ret.put(FontStyle.FontColor, text.getFontColor());
        ret.put(FontStyle.FontFamily, text.getFontFamily());
        return ret;
    }

    public static void setFontStyle(XSLFTextRun run, Map<FontStyle, Object> fontStyle) {
        if (fontStyle.containsKey(FontStyle.FontColor)) {
            run.setFontColor((PaintStyle) fontStyle.get(FontStyle.FontColor));
        }
        if (fontStyle.containsKey(FontStyle.FontFamily)) {
            run.setFontFamily((String) fontStyle.get(FontStyle.FontFamily));
        }
        if (fontStyle.containsKey(FontStyle.FontSize)) {
            run.setFontSize((Double) fontStyle.get(FontStyle.FontSize));
        }
    }

    public static void replaceTextShape(XSLFTextShape shape, String newStr) {
        replaceTextShape(shape, null, newStr);
    }

    public static void replaceTextShape(XSLFTextShape shape, String oldStr, String newStr, TextAlign align) {
        String text = shape.getText();
        Map<FontStyle, Object> fontStyle = PlanUtil.getFontStyle(shape);
        shape.clearText();
        XSLFTextParagraph paragraph = shape.addNewTextParagraph();
        paragraph.setTextAlign(align);
        XSLFTextRun run = paragraph.addNewTextRun();
        PlanUtil.setFontStyle(run, fontStyle);
        if (oldStr == null || oldStr.isEmpty()) {
            run.setText(newStr);
        } else {
            run.setText(text.replace(oldStr, newStr));
        }
    }

    public static void replaceTextShape(XSLFTextShape shape, String oldStr, String newStr) {
        replaceTextShape(shape, oldStr, newStr, TextAlign.CENTER);
    }

}
