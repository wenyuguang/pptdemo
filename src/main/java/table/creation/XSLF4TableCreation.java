package table.creation;

import org.apache.poi.xslf.usermodel.*;

import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * XSLF构建表格示例
 * Created by cx on 2015/7/8.
 */
public class XSLF4TableCreation {
    public static void main(String[] args) throws IOException {
        XMLSlideShow ppt = new XMLSlideShow();
        XSLFSlide slide = ppt.createSlide();

        int COLUMN_NUM = 3;
        int ROW_NUM = 3;
        XSLFTable table = slide.createTable();
        table.setAnchor(new Rectangle2D.Double(50, 50, 450, 300));

        XSLFTableRow headerRow = table.addRow();
        headerRow.setHeight(50);

        for (int i = 0; i < COLUMN_NUM; i++) {
            XSLFTableCell th = headerRow.addCell();

            XSLFTextParagraph textParagraph = th.addNewTextParagraph();
            textParagraph.setTextAlign(TextAlign.CENTER);

            XSLFTextRun textRun = textParagraph.addNewTextRun();
            textRun.setText("Header" + (i+1));
            textRun.setBold(true);

            th.setFillColor(new Color(79, 129, 189));
            th.setBorderBottom(2);
            th.setBorderBottomColor(Color.white);

            table.setColumnWidth(i, 150);
        }

        for(int rowNum = 0; rowNum < ROW_NUM; rowNum ++){
            XSLFTableRow tr = table.addRow();
            tr.setHeight(50);

            for(int i = 0; i < COLUMN_NUM; i++) {
                XSLFTableCell cell = tr.addCell();
                XSLFTextParagraph textParagraph = cell.addNewTextParagraph();

                XSLFTextRun textRun = textParagraph.addNewTextRun();
                textRun.setText("Cell " + (i+1));

                if(rowNum % 2 == 0)
                    cell.setFillColor(new Color(208, 216, 232));
                else
                    cell.setFillColor(new Color(233, 247, 244));
            }
        }

        FileOutputStream out = new FileOutputStream("output/table_creation.pptx");
        ppt.write(out);
        out.close();
    }
}
