package picture.creation;

import org.apache.poi.hslf.model.Picture;
import org.apache.poi.hslf.model.Slide;
import org.apache.poi.hslf.usermodel.SlideShow;

import java.awt.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * HSLF构建图片示例
 * Created by cx on 2015/7/8.
 */
public class HSLF4PictureCreation {
    public static void main(String[] args) throws IOException {
        SlideShow ppt = new SlideShow();
        Slide slide = ppt.createSlide();

        int idx = ppt.addPicture(new File("data/girl.jpg"), Picture.JPEG);
        Picture picture = new Picture(idx);

        //set position
        picture.setAnchor(new Rectangle(100, 100, 100, 80));

        slide.addShape(picture);

        FileOutputStream out = new FileOutputStream("output/picture_creation.ppt");
        ppt.write(out);
        out.close();
    }
}
