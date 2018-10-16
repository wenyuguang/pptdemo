package chart.extraction;

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
 * HSLF抽取图表示例
 * Created by cx on 2015/7/8.
 */
public class HSLF4ChartExtraction {
    private static final Logger logger = LoggerFactory.getLogger(HSLF4ChartExtraction.class);

    public static void main(String[] args) throws IOException {
        SlideShow ppt = new SlideShow(new HSLFSlideShow("data/chart_extraction.ppt"));
        Slide slide = ppt.getSlides()[0];
        Shape shape = slide.getShapes()[0];

        logger.debug("the shape of chart is a picture:{}", shape instanceof Picture);

        Picture picture = (Picture) shape;
        PictureData pictureData = picture.getPictureData();
        byte[] data = pictureData.getData();
        int type = pictureData.getType();
        String ext = getExt(type);

        if (!ext.isEmpty()) {
            //读取幻灯片中的图表，输出到指定目录
            outputFile("chart" + ext, data);
        }
    }

    private static String getExt(int type) {
        String ext = "";

        switch (type) {
            case Picture.JPEG: ext=".jpg"; break;
            case Picture.PNG: ext=".png"; break;
            case Picture.WMF: ext=".wmf"; break;
            case Picture.EMF: ext=".emf"; break;
            case Picture.PICT: ext=".pict"; break;
        }

        return ext;
    }

    private static void outputFile(String pictureName, byte[] data) throws IOException {
        FileOutputStream out = new FileOutputStream("output/image/" + pictureName);
        out.write(data);
        out.close();
    }
}
