package video.extraction;

import org.apache.poi.hslf.HSLFSlideShow;
import org.apache.poi.hslf.model.MovieShape;
import org.apache.poi.hslf.model.Shape;
import org.apache.poi.hslf.model.Slide;
import org.apache.poi.hslf.usermodel.SlideShow;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * HSLF抽取视频示例
 * Created by cx on 2015/7/7.
 */
public class HSLF4VideoExtraction {
    public static void main(String[] args) throws IOException {
        SlideShow ppt = new SlideShow(new HSLFSlideShow("data/video_extraction.ppt"));
        Slide slide = ppt.getSlides()[0];
        Shape[] shapes = slide.getShapes();

        for (Shape shape : shapes) {
            if (shape instanceof MovieShape) {
                MovieShape movie = (MovieShape) shape;

                //读取ppt中视频的绝对路径，输出到其他目录
                File file = new File(movie.getPath());
                FileOutputStream out = new FileOutputStream("output/" + file.getName());
                FileInputStream in = new FileInputStream(file);
                byte[] data = new byte[1024];

                while (in.read(data) != -1) {
                    out.write(data);
                }

                out.flush();
                in.close();
                out.close();
            }
        }
    }
}
