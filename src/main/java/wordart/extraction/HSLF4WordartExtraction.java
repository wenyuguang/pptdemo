package wordart.extraction;

import org.apache.poi.hslf.HSLFSlideShow;
import org.apache.poi.hslf.model.Picture;
import org.apache.poi.hslf.model.Shape;
import org.apache.poi.hslf.model.Slide;
import org.apache.poi.hslf.usermodel.PictureData;
import org.apache.poi.hslf.usermodel.SlideShow;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * HSLF抽取艺术字示例
 * Created by cx on 2015/6/30.
 */
public class HSLF4WordartExtraction {
    private static final Logger logger = LoggerFactory.getLogger(HSLF4WordartExtraction.class);

    public static void main(String[] args) throws IOException {
        SlideShow ppt = new SlideShow(new HSLFSlideShow("data/wordart_extraction.ppt"));
        Slide slide = ppt.getSlides()[0];
        Shape[] shapes = slide.getShapes();

        for (int i = 0; i < shapes.length; i++) {
            Shape shape = shapes[i];
            if (shape instanceof Picture) {
                logger.debug("艺术字被识别为图片");

                PictureData pictureData = ((Picture) shape).getPictureData();
                byte[] data = pictureData.getData();
                int type = pictureData.getType();
                String ext;
                switch (type) {
                    case Picture.JPEG: ext=".jpg"; break;
                    case Picture.PNG: ext=".png"; break;
                    case Picture.WMF: ext=".wmf"; break;
                    case Picture.EMF: ext=".emf"; break;
                    case Picture.PICT: ext=".pict"; break;
                    default: continue;
                }

                FileOutputStream out = new FileOutputStream("output\\image\\wordart" + ext);
                out.write(data);
                out.close();
            }
        }
    }
}
