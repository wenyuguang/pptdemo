package hyperlinks.extraction;

import org.apache.poi.hslf.HSLFSlideShow;
import org.apache.poi.hslf.model.Hyperlink;
import org.apache.poi.hslf.model.Shape;
import org.apache.poi.hslf.model.Slide;
import org.apache.poi.hslf.model.TextRun;
import org.apache.poi.hslf.usermodel.SlideShow;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;

/**
 * HSLF抽取超链接示例
 * Created by cx on 2015/7/7.
 */
public class HSLF4HyperlinkExtraction {
    private static final Logger logger = LoggerFactory.getLogger(HSLF4HyperlinkExtraction.class);

    public static void main(String[] args) throws IOException {
        SlideShow ppt = new SlideShow(new HSLFSlideShow("data/hyperlink_extraction.ppt"));

        Slide slideOne = ppt.getSlides()[0];
        Slide slideTwo = ppt.getSlides()[1];

        //读取文本超链接
        TextRun[] textRuns = slideOne.getTextRuns();

        for (TextRun textRun : textRuns) {
            Hyperlink[] links = textRun.getHyperlinks();

            if (links != null) {
                for (Hyperlink link : links) {
                    logger.debug("the text link title is {}", link.getTitle());
                    logger.debug("the text link address is {}", link.getAddress());
                }
            }
        }

        //读取自选图形的超链接
        Shape[] shapes = slideTwo.getShapes();

        for (Shape shape : shapes) {
            Hyperlink link = shape.getHyperlink();
            if (link != null) {
                logger.debug("the shape link address is {}", link.getAddress());
            }
        }
    }
}
