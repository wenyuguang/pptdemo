package shape.creation;

import org.apache.poi.xslf.usermodel.*;

import java.awt.*;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * XSLF构建图形示例
 * Created by cx on 2015/6/30.
 */
public class XSLF4ShapeCreation {
    public static void main(String[] args) throws IOException {
        XMLSlideShow ppt = new XMLSlideShow();
        XSLFSlide slideForLine = ppt.createSlide();
        XSLFSlide slideForTextBox = ppt.createSlide();
        XSLFSlide slideForTrapezoid = ppt.createSlide();
        XSLFSlide slideForShapeGroup = ppt.createSlide();

        //构建直线
        XSLFConnectorShape line = slideForLine.createConnector();
        line.setLineColor(Color.red);
        line.setAnchor(new Rectangle(200, 200, 200, 100));

        //构建一个文本框
        XSLFTextBox textBox = slideForTextBox.createTextBox();
        textBox.setAnchor(new Rectangle(100, 100, 300, 50));
        //设置文本框格式
        XSLFTextParagraph textBoxParagraph = textBox.addNewTextParagraph();
        XSLFTextRun textBoxStyle = textBoxParagraph.addNewTextRun();
        textBoxStyle.setText("new textBox");
        textBoxStyle.setFontSize(32);
        textBoxStyle.setUnderline(true);
        textBoxStyle.setFontColor(Color.blue);

        //构建梯形
        XSLFAutoShape trapezoid = slideForTrapezoid.createAutoShape();
        trapezoid.setShapeType(XSLFShapeType.TRAPEZOID);
        trapezoid.setAnchor(new java.awt.Rectangle(150, 150, 100, 200));
        trapezoid.setFillColor(Color.blue);

        /**
         * 构建组合图形，下面的方式生成的图形不是组合在一起的，目前没找到关于XSLF操作组合图形的例子
         */
        XSLFGroupShape groupShape = slideForShapeGroup.createGroup();
        groupShape.setAnchor(new Rectangle(100, 100, 300, 300));

        XSLFAutoShape autoShape = slideForShapeGroup.createAutoShape();
        autoShape.setAnchor(new Rectangle(100, 150, 200, 200));
        autoShape.setLineWidth(5);
        autoShape.setLineColor(Color.black);

        XSLFConnectorShape lineOfGroup = slideForShapeGroup.createConnector();
        lineOfGroup.setAnchor(new Rectangle(100, 150, 200, 200));
        lineOfGroup.setLineColor(Color.red);
        lineOfGroup.setLineWidth(1);

        FileOutputStream out = new FileOutputStream("output/shape_creation.pptx");
        ppt.write(out);
        out.close();
    }
}
