package text.extraction;

import org.apache.poi.hslf.HSLFSlideShow;
import org.apache.poi.hslf.model.Slide;
import org.apache.poi.hslf.model.TextRun;
import org.apache.poi.hslf.usermodel.SlideShow;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;

/**
 * HSLF抽取文本示例
 * Created by cx on 2015/6/30.
 */
public class HSLF4TextExtraction {
    private static final Logger logger = LoggerFactory.getLogger(HSLF4TextExtraction.class);

    public static void main(String[] args) throws IOException {
        SlideShow ppt = new SlideShow(new HSLFSlideShow("data/text_extraction.ppt"));

        Slide[] slides = ppt.getSlides();
        for (int i = 0; i < slides.length; i++) {
            logger.debug("Slide {}:", i + 1);

            Slide slide = slides[i];
            TextRun[] textRuns = slide.getTextRuns();
            StringBuilder stringBuilder = new StringBuilder();

            for (int j = 0; j < textRuns.length; j++) {
                TextRun textRun = textRuns[j];
                stringBuilder.append(textRun.getText()).append("\n");
            }

            logger.debug("The content of slide is \n {}", stringBuilder.toString());
        }
    }
}
