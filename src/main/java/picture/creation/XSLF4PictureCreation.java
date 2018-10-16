package picture.creation;

import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import java.awt.geom.Rectangle2D;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * XSLF构建图片示例
 * Created by cx on 2015/7/8.
 */
public class XSLF4PictureCreation {
    public static void main(String[] args) throws IOException {
        XMLSlideShow ppt = new XMLSlideShow();
        XSLFSlide slide = ppt.createSlide();

        File image = new File("data/girl.jpg");
        byte[] data = IOUtils.toByteArray(new FileInputStream(image));
        int pictureIndex = ppt.addPicture(data, XSLFPictureData.PICTURE_TYPE_JPEG);

        XSLFPictureShape picture = slide.createPicture(pictureIndex);
        picture.setAnchor(new Rectangle2D.Double(50, 50, 150, 200));

        FileOutputStream out = new FileOutputStream("output/picture_creation.pptx");
        ppt.write(out);
        out.close();
    }
}
