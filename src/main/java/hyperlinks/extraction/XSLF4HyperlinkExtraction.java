package hyperlinks.extraction;

import org.apache.poi.xslf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

/**
 * XSLF抽取超链接示例
 * Created by cx on 2015/7/7.
 */
public class XSLF4HyperlinkExtraction {
    private static final Logger logger = LoggerFactory.getLogger(XSLF4HyperlinkExtraction.class);

    public static void main(String[] args) throws IOException {
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("data/hyperlink_extraction.pptx"));
        XSLFSlide slideOne = ppt.getSlides()[0];

        //读取文本超链接
        XSLFTextBox textBox = (XSLFTextBox) slideOne.getShapes()[0];
        List<XSLFTextParagraph> textParagraphs = textBox.getTextParagraphs();

        for (XSLFTextParagraph textParagraph : textParagraphs) {
            List<XSLFTextRun> textRuns = textParagraph.getTextRuns();

            for (XSLFTextRun textRun : textRuns) {
                XSLFHyperlink link = textRun.getHyperlink();

                logger.debug("the text link address is {}", link.getTargetURI());
            }
        }
    }
}
