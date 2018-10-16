package text.titleExtraction;

import org.apache.poi.hslf.HSLFSlideShow;
import org.apache.poi.hslf.model.Slide;
import org.apache.poi.hslf.usermodel.SlideShow;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;

/**
 * HSLF抽取幻灯片标题示例
 * Created by cx on 2015/6/30.
 */
public class HSLF4TitleExtraction {
    private static final Logger logger = LoggerFactory.getLogger(HSLF4TitleExtraction.class);

    public static void main(String[] args) throws IOException {
        SlideShow ppt = new SlideShow(new HSLFSlideShow("data/text_extraction.ppt"));

        Slide[] slides = ppt.getSlides();
        for (int i = 0; i < slides.length; i++) {
            Slide slide = slides[i];

            logger.debug("Slide {}: the title is {}", i + 1, slide.getTitle());
        }
    }
}
