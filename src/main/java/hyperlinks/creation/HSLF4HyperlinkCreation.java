package hyperlinks.creation;

import org.apache.poi.hslf.model.AutoShape;
import org.apache.poi.hslf.model.Hyperlink;
import org.apache.poi.hslf.model.ShapeTypes;
import org.apache.poi.hslf.model.Slide;
import org.apache.poi.hslf.usermodel.SlideShow;

import java.awt.*;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * HSLF构建超链接示例
 * Created by cx on 2015/7/7.
 */
public class HSLF4HyperlinkCreation {
    public static void main(String[] args) throws IOException {
        SlideShow ppt = new SlideShow();
        Slide slide = ppt.createSlide();

        AutoShape shape = new AutoShape(ShapeTypes.Rectangle);
        shape.setAnchor(new Rectangle(50, 100, 50, 100));

        Hyperlink link = new Hyperlink();
        link.setTitle("百度");
        link.setAddress("http://www.baidu.com");

        ppt.addHyperlink(link);
        shape.setHyperlink(link);

        slide.addShape(shape);

        FileOutputStream out = new FileOutputStream("output/hyperlink_creation.ppt");
        ppt.write(out);
        out.close();
    }
}
