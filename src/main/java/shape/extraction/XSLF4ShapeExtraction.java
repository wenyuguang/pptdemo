package shape.extraction;

import org.apache.poi.xslf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.IOException;

/**
 * HSLF抽取图形示例
 * Created by cx on 2015/6/30.
 */
public class XSLF4ShapeExtraction {
    private static final Logger logger = LoggerFactory.getLogger(XSLF4ShapeExtraction.class);

    public static void main(String[] args) throws IOException {
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("data/shape_extraction.pptx"));
        XSLFSlide slide = ppt.getSlides()[0];
        XSLFShape[] shapes = slide.getShapes();

        for (int i = 0; i < shapes.length; i++) {
            XSLFShape shape = shapes[i];
            String type;

            if (shape instanceof XSLFConnectorShape) {
                type = "连接线";
            }
            else if (shape instanceof XSLFTextShape) {
                XSLFTextShape textShape = (XSLFTextShape) shape;
                type = "文本框:" + textShape.getText();
            }
            else if (shape instanceof XSLFPictureShape) {
                XSLFPictureShape pictureShape = (XSLFPictureShape) shape;
                type = "图片:" + pictureShape.getPictureData().getFileName();
            }
            else if (shape instanceof XSLFGroupShape) {
                type = "组合图形";
            }
            else {
                type = "未知";
            }

            logger.debug("shape{}:type->{}", i+1, type);
        }
    }
}
