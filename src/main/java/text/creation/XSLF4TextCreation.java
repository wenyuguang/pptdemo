package text.creation;

import org.apache.poi.xslf.usermodel.*;

import java.awt.*;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * XSLF构建文本示例
 * Created by cx on 2015/6/30.
 */
public class XSLF4TextCreation {
    public static void main(String[] args) throws IOException {
        XMLSlideShow ppt = new XMLSlideShow();
        XSLFSlide slide = ppt.createSlide();

        //构建一个标题
        XSLFTextShape title = slide.createTextBox();
        title.setPlaceholder(Placeholder.TITLE);
        //如果需要设置样式，使用XSLFTextRun的setTitle方法；使用下面注释的代码，打开生成的ppt时会报错
//        title.setText("Hello");
        title.setAnchor(new Rectangle(50, 50, 400, 100));

        //设置标题格式
        XSLFTextParagraph titleParagraph = title.addNewTextParagraph();
        XSLFTextRun titleStyle = titleParagraph.addNewTextRun();
        titleStyle.setText("Hello");
        titleStyle.setFontColor(Color.red);
        titleStyle.setBold(true);

        //构建一个文本框
        XSLFTextBox textBox = slide.createTextBox();
        //如果需要设置样式，使用XSLFTextRun的setTitle方法；使用下面注释的代码，打开生成的ppt时会报错
//        textBox.setText("new textBox");
        textBox.setAnchor(new Rectangle(100, 100, 300, 50));

        //设置文本框格式
        XSLFTextParagraph textBoxParagraph = textBox.addNewTextParagraph();
        XSLFTextRun textBoxStyle = textBoxParagraph.addNewTextRun();
        textBoxStyle.setText("new textBox");
        textBoxStyle.setFontSize(32);
        textBoxStyle.setUnderline(true);
        textBoxStyle.setFontColor(Color.blue);

        FileOutputStream out = new FileOutputStream("output/text_creation.pptx");
        ppt.write(out);
        out.close();
    }
}
