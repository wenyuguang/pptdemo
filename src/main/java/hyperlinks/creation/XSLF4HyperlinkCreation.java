package hyperlinks.creation;

import org.apache.poi.xslf.usermodel.*;

import java.awt.*;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * XSLF构建超链接示例
 * Created by cx on 2015/7/7.
 */
public class XSLF4HyperlinkCreation {
    public static void main(String[] args) throws IOException {
        XMLSlideShow ppt = new XMLSlideShow();
        XSLFSlide slide = ppt.createSlide();

        XSLFTextBox textBox = slide.createTextBox();
        textBox.setAnchor(new Rectangle(50, 50, 200, 50));
        XSLFTextRun textRun = textBox.addNewTextParagraph().addNewTextRun();
        XSLFHyperlink link = textRun.createHyperlink();

        textRun.setText("百度");
        link.setAddress("http://www.baidu.com");

        FileOutputStream out = new FileOutputStream("output/hyperlink_creation.pptx");
        ppt.write(out);
        out.close();
    }
}
