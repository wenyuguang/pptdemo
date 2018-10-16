package table.extraction;

import org.apache.poi.hslf.HSLFSlideShow;
import org.apache.poi.hslf.model.Shape;
import org.apache.poi.hslf.model.Slide;
import org.apache.poi.hslf.model.Table;
import org.apache.poi.hslf.model.TableCell;
import org.apache.poi.hslf.usermodel.SlideShow;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;

/**
 * HSLF抽取表格示例
 * Created by cx on 2015/7/8.
 */
public class HSLF4TableExtraction {
    private static final Logger logger = LoggerFactory.getLogger(HSLF4TableExtraction.class);

    public static void main(String[] args) throws IOException {
        SlideShow ppt = new SlideShow(new HSLFSlideShow("data/table_extraction.ppt"));
        Slide slide = ppt.getSlides()[0];
        Shape shape = slide.getShapes()[0];

        Table table = (Table) shape;
        int columnNum = table.getNumberOfColumns();
        int rowNum = table.getNumberOfRows();
        StringBuilder tableContent = new StringBuilder();

        for (int i = 0; i < rowNum; i++) {
            for (int j = 0; j < columnNum; j++) {
                TableCell cell = table.getCell(i, j);

                if (cell != null) { //避免合并单元格造成的空指针异常
                    tableContent.append("|").append(cell.getText());
                }
            }

            tableContent.append("|").append("\n");
        }

        logger.debug("the table is \n {}", tableContent.toString());
    }
}
