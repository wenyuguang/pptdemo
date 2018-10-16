package picture.extraction;

import org.apache.poi.xslf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

/**
 * XSLF抽取图片示例
 * Created by cx on 2015/7/8.
 */
public class XSLF4PictureExtraction {
    public static void main(String[] args) throws IOException {
        //抽取幻灯片中的图片
        extractPictrueFromSlide();
        //抽取PPT中的所有图片
        extractPictruesFromPPT();
    }

    private static void extractPictrueFromSlide() throws IOException {
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("data/picture_extraction.pptx"));
        XSLFSlide slide = ppt.getSlides()[0];
        XSLFShape shape = slide.getShapes()[0];
        XSLFPictureShape picture = (XSLFPictureShape) shape;

        XSLFPictureData pictureData = picture.getPictureData();
        byte[] data = pictureData.getData();
        int type = pictureData.getPictureType();
        String ext = getExt(type);

        if (!ext.isEmpty()) {
            //读取幻灯片中的图片，输出到指定目录
            outputFile("picture_x" + ext, data);
        }
    }

    private static void extractPictruesFromPPT() throws IOException {
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("data/picture_extraction.pptx"));
        List<XSLFPictureData> pictureDatas = ppt.getAllPictures();
        int index = 0;

        for (XSLFPictureData pictureData : pictureDatas) {
            byte[] data = pictureData.getData();
            int type = pictureData.getPictureType();
            String ext = getExt(type);

            if (!ext.isEmpty()) {
                //读取幻灯片中的图片，输出到指定目录
                outputFile("picture_x" + index + ext, data);
            }

            index++;
        }
    }

    private static String getExt(int type) {
        String ext = "";

        switch (type) {
            case XSLFPictureData.PICTURE_TYPE_BMP: ext=".bmp"; break;
            case XSLFPictureData.PICTURE_TYPE_DIB: ext=".dib"; break;
            case XSLFPictureData.PICTURE_TYPE_EMF: ext=".emf"; break;
            case XSLFPictureData.PICTURE_TYPE_EPS: ext=".eps"; break;
            case XSLFPictureData.PICTURE_TYPE_GIF: ext=".gif"; break;
            case XSLFPictureData.PICTURE_TYPE_JPEG: ext=".jpeg"; break;
            case XSLFPictureData.PICTURE_TYPE_PICT: ext=".pict"; break;
            case XSLFPictureData.PICTURE_TYPE_PNG: ext=".png"; break;
            case XSLFPictureData.PICTURE_TYPE_TIFF: ext=".tiff"; break;
            case XSLFPictureData.PICTURE_TYPE_WDP: ext=".wdp"; break;
            case XSLFPictureData.PICTURE_TYPE_WMF: ext=".wmf"; break;
            case XSLFPictureData.PICTURE_TYPE_WPG: ext=".wpg"; break;
        }

        return ext;
    }

    private static void outputFile(String pictureName, byte[] data) throws IOException {
        FileOutputStream out = new FileOutputStream("output/image/" + pictureName);
        out.write(data);
        out.close();
    }
}
