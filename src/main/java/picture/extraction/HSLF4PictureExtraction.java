package picture.extraction;

import org.apache.poi.hslf.HSLFSlideShow;
import org.apache.poi.hslf.model.Picture;
import org.apache.poi.hslf.model.Shape;
import org.apache.poi.hslf.model.Slide;
import org.apache.poi.hslf.usermodel.PictureData;
import org.apache.poi.hslf.usermodel.SlideShow;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * HSLF抽取图片示例
 * Created by cx on 2015/7/8.
 */
public class HSLF4PictureExtraction {
    public static void main(String[] args) throws IOException {
        //抽取幻灯片中的图片
        extractPictrueFromSlide();
        //抽取PPT中的所有图片
        extractPictruesFromPPT();
    }

    private static void extractPictrueFromSlide() throws IOException {
        SlideShow ppt = new SlideShow(new HSLFSlideShow("data/picture_extraction.ppt"));
        Slide slide = ppt.getSlides()[0];
        Shape shape = slide.getShapes()[0];

        Picture picture = (Picture) shape;
        PictureData pictureData = picture.getPictureData();
        byte[] data = pictureData.getData();
        int type = pictureData.getType();
        String ext = getExt(type);

        if (!ext.isEmpty()) {
            //读取幻灯片中的图片，输出到指定目录
            outputFile("picture" + ext, data);
        }
    }

    private static void extractPictruesFromPPT() throws IOException {
        SlideShow ppt = new SlideShow(new HSLFSlideShow("data/picture_extraction.ppt"));
        PictureData[] pictureDatas = ppt.getPictureData();

        for (int i = 0; i < pictureDatas.length; i++) {
            PictureData pictureData = pictureDatas[i];

            byte[] data = pictureData.getData();
            int type = pictureData.getType();
            String ext = getExt(type);

            if (!ext.isEmpty()) {
                //读取幻灯片中的图片，输出到指定目录
                outputFile("picture_" + i + ext, data);
            }
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
