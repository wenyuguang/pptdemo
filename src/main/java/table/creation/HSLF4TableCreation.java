package table.creation;

import org.apache.poi.hslf.model.Slide;
import org.apache.poi.hslf.model.Table;
import org.apache.poi.hslf.model.TableCell;
import org.apache.poi.hslf.usermodel.SlideShow;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * HSLF构建表格示例
 * Created by cx on 2015/7/8.
 */
public class HSLF4TableCreation {
    public static void main(String[] args) throws IOException {
        int COLUMN_NUM = 3;
        int ROW_NUM = 3;
        SlideShow ppt = new SlideShow();
        Slide slide = ppt.createSlide();

        Table table = new Table(ROW_NUM, COLUMN_NUM);
        for (int i = 0; i < ROW_NUM; i++) {
            for (int j = 0; j < COLUMN_NUM; j++) {
                TableCell cell = table.getCell(i, j);

                cell.setText("xxxxxx");
            }
        }
        slide.addShape(table);

        FileOutputStream out = new FileOutputStream("output/table_creation.ppt");
        ppt.write(out);
        out.close();
    }
}
