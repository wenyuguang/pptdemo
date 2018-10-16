package shape.creation;

import org.apache.poi.hslf.model.*;
import org.apache.poi.hslf.usermodel.RichTextRun;
import org.apache.poi.hslf.usermodel.SlideShow;

import java.awt.*;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * HSLF构建图形示例
 * Created by cx on 2015/6/30.
 */
public class HSLF4ShapeCreation {
    public static void main(String[] args) throws IOException {
        SlideShow ppt = new SlideShow();
        Slide slideForLine = ppt.createSlide();
        Slide slideForTextBox = ppt.createSlide();
        Slide slideForTrapezoid = ppt.createSlide();
        Slide slideForShapeGroup = ppt.createSlide();

        //构建直线图形
        Line line = new Line();
        line.setAnchor(new java.awt.Rectangle(100, 100, 200, 200));
        line.setLineColor(Color.blue);
        line.setLineStyle(Line.LINE_SIMPLE);
        slideForLine.addShape(line);

        //构建文本框图形
        TextBox txt = new TextBox();
        txt.setText("横向文本");
        txt.setAnchor(new java.awt.Rectangle(300, 100, 300, 50));
        //设置文本框中文字样式
        RichTextRun rt = txt.getTextRun().getRichTextRuns()[0];
        rt.setFontSize(32);
        rt.setFontName("宋体");
        rt.setBold(true);
        rt.setItalic(true);
        rt.setUnderlined(true);
        rt.setFontColor(Color.red);
        rt.setAlignment(TextBox.AlignLeft);
        slideForTextBox.addShape(txt);

        //构建梯形图形
        AutoShape trapezoid = new AutoShape(ShapeTypes.Trapezoid);
        trapezoid.setAnchor(new java.awt.Rectangle(150, 150, 100, 200));
        trapezoid.setFillColor(Color.blue);
        slideForTrapezoid.addShape(trapezoid);

        /**
         * 构建组合图形
         *
         * 主观认为可以直接调用ShapeGroup的addShape方法进行简单图形的组合，直接输出在Slide上，但真正执行起来不可以。
         * 目前还没有发现原因；个人推测ShapGroup需要结合Graphics2D才能发挥输出的作用。
         * 官网及网上搜索的关于ShapeGroup的用法均是结合Graphics2D进行渲染的。
         */
        ShapeGroup shapeGroup = new ShapeGroup();
        shapeGroup.setAnchor(new Rectangle(100, 100, 300, 300));
        slideForShapeGroup.addShape(shapeGroup);

        Graphics2D graphics = new PPGraphics2D(shapeGroup);

        //画出一个矩形
        graphics.setColor(Color.red);
        graphics.fillRect(120, 120, 20, 30);

        //画出一个直线
        graphics.setColor(Color.black);
        graphics.drawLine(140, 135, 160, 135);

        //再画出一个矩形
        graphics.setColor(Color.red);
        graphics.fillRect(160, 120, 20, 30);

        FileOutputStream out = new FileOutputStream("output/shape_creation.ppt");
        ppt.write(out);
        out.close();
    }
}
