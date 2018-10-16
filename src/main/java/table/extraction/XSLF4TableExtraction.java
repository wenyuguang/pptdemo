package table.extraction;

import org.apache.poi.xslf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

/**
 * XSLF抽取表格示例
 * Created by cx on 2015/7/8.
 */
public class XSLF4TableExtraction {
    private static final Logger logger = LoggerFactory.getLogger(XSLF4TableExtraction.class);

    public static void main(String[] args) throws IOException {
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("data/table_extraction.pptx"));
        XSLFSlide slide = ppt.getSlides()[0];
        XSLFShape shape = slide.getShapes()[0];

        XSLFTable table = (XSLFTable) shape;
        List<XSLFTableRow> rows = table.getRows();
        StringBuilder tableContent = new StringBuilder();

        for (XSLFTableRow row : rows) {
            List<XSLFTableCell> cells = row.getCells();

            for (XSLFTableCell cell : cells) {
                tableContent.append("|").append(cell.getText());
            }

            tableContent.append("|").append("\n");
        }

        logger.debug("the table is \n {}", tableContent.toString());
    }
}
