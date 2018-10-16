package text.titleExtraction;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.IOException;

/**
 * XSLF抽取文本示例
 * Created by cx on 2015/6/30.
 */
public class XSLF4TitleExtraction {
    private static final Logger logger = LoggerFactory.getLogger(XSLF4TitleExtraction.class);

    public static void main(String[] args) throws IOException {
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("data/text_extraction.pptx"));

        XSLFSlide[] slides = ppt.getSlides();
        for (int i = 0; i < slides.length; i++) {
            XSLFSlide slide = slides[i];

            //XSLFSlide的getTile方法有些问题，当ppt第一张幻灯片为“标题幻灯片”版式时，读不到title
            logger.debug("Slide {}: the title is {}", i + 1, slide.getTitle());
        }
    }
}
