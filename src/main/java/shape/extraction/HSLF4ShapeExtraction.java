package shape.extraction;

import org.apache.poi.hslf.HSLFSlideShow;
import org.apache.poi.hslf.model.*;
import org.apache.poi.hslf.usermodel.SlideShow;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;

/**
 * HSLF抽取图形示例
 * Created by cx on 2015/6/30.
 */
public class HSLF4ShapeExtraction {
    private static final Logger logger = LoggerFactory.getLogger(HSLF4ShapeExtraction.class);

    public static void main(String[] args) throws IOException {
        SlideShow ppt = new SlideShow(new HSLFSlideShow("data/shape_extraction.ppt"));
        Slide slide = ppt.getSlides()[0];
        Shape[] shapes = slide.getShapes();

        for (int i = 0; i < shapes.length; i++) {
            Shape shape = shapes[i];
            String type;

            if (shape instanceof AutoShape) {
                AutoShape autoShape = (AutoShape) shape;
                type = "自选图形:" + autoShape.getText();
            }
            else if (shape instanceof Line) {
                type = "直线";
            }
            else if (shape instanceof TextBox) {
                TextBox textBox = (TextBox) shape;
                type = "文本框:" + textBox.getText();
            }
            else if (shape instanceof Picture) {
                type = "图片";
            }
            else if (shape instanceof ShapeGroup) {
                type = "组合图形";
            }
            else {
                type = "未知";
            }

            logger.debug("shape{}:type->{}", i+1, type);
        }
    }
}
