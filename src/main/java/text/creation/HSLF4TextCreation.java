package text.creation;

import org.apache.poi.hslf.model.Slide;
import org.apache.poi.hslf.model.TextBox;
import org.apache.poi.hslf.usermodel.RichTextRun;
import org.apache.poi.hslf.usermodel.SlideShow;

import java.awt.*;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * HSLF构建文本示例
 * Created by cx on 2015/6/30.
 */
public class HSLF4TextCreation {
    public static void main(String[] args) throws IOException {
        SlideShow ppt = new SlideShow();
        Slide slide = ppt.createSlide();

        //构建幻灯片标题
        TextBox title = slide.addTitle();
        title.setText("Hello");

        //设置标题格式
        RichTextRun titleStyle = title.getTextRun().getRichTextRuns()[0];
        titleStyle.setFontColor(Color.red);
        titleStyle.setBold(true);

        //构建一个文本框
        TextBox textBox = new TextBox();
        textBox.setText("new TextBox");
        textBox.setAnchor(new Rectangle(100, 100, 300, 50));

        //设置文本框格式
        RichTextRun textBoxStyle = textBox.getTextRun().getRichTextRuns()[0];
        textBoxStyle.setFontSize(32);
        textBoxStyle.setUnderlined(true);
        textBoxStyle.setFontColor(Color.blue);

        slide.addShape(textBox);

        FileOutputStream out = new FileOutputStream("output/text_creation.ppt");
        ppt.write(out);
        out.close();
    }
}
