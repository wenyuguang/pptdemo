package text.extraction;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.IOException;

/**
 * XSLF抽取文本示例
 * Created by cx on 2015/6/30.
 */
public class XSLF4TextExtraction {
    private static final Logger logger = LoggerFactory.getLogger(XSLF4TextExtraction.class);

    public static void main(String[] args) throws IOException {
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("data/text_extraction.pptx"));

        XSLFSlide[] slides = ppt.getSlides();
        for (int i = 0; i < slides.length; i++) {
            logger.debug("Slide {}:", i + 1);

            XSLFSlide slide = slides[i];
            XSLFTextShape[] textShapes = slide.getPlaceholders();
            StringBuilder stringBuilder = new StringBuilder();

            for (int j = 0; j < textShapes.length; j++) {
                XSLFTextShape textShape = textShapes[j];

                stringBuilder.append(textShape.getText()).append("\n");
            }

            logger.debug("The content of slide is \n {}", stringBuilder.toString());
        }
    }
}
