package chart.extraction;

import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xslf.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;

import java.io.*;

/**
 * XSLF抽取图表示例
 * Created by cx on 2015/7/9.
 */
public class XSLF4ChartExtraction {
    public static void main(String[] args) throws IOException {
        final String PPT_TEMPLATE = "data/pie-chart-template.pptx";
        final String DATA_FILE = "data/pie-chart-data.txt";

        //加载数据文件
        BufferedReader modelReader = new BufferedReader(new FileReader(DATA_FILE));
        try {
            //从数据文件读取数据标题
            String chartTitle = modelReader.readLine();  // first line is chart title

            //读取模版PPT
            XMLSlideShow pptx = new XMLSlideShow(new FileInputStream(PPT_TEMPLATE));
            XSLFSlide slide = pptx.getSlides()[0];

            // 读取模版PPT中的图表
            XSLFChart chart = null;
            for(POIXMLDocumentPart part : slide.getRelations()){
                if(part instanceof XSLFChart){
                    chart = (XSLFChart) part;
                    break;
                }
            }

            if(chart == null)
                throw new IllegalStateException("chart not found in the template");

            // 获取图表关联的excel数据源
            POIXMLDocumentPart xlsPart = chart.getRelations().get(0);
            XSSFWorkbook wb = new XSSFWorkbook();
            try {
                XSSFSheet sheet = wb.createSheet();

                CTChart ctChart = chart.getCTChart();
                CTPlotArea plotArea = ctChart.getPlotArea();

                CTPieChart pieChart = plotArea.getPieChartArray(0);
                //Pie Chart Series
                CTPieSer ser = pieChart.getSerArray(0);

                // Series Text
                CTSerTx tx = ser.getTx();
                tx.getStrRef().getStrCache().getPtArray(0).setV(chartTitle);
                sheet.createRow(0).createCell(1).setCellValue(chartTitle);
                String titleRef = new CellReference(sheet.getSheetName(), 0, 1, true, true).formatAsString();
                tx.getStrRef().setF(titleRef);

                // Category Axis Data
                CTAxDataSource cat = ser.getCat();
                CTStrData strData = cat.getStrRef().getStrCache();

                // Values
                CTNumDataSource val = ser.getVal();
                CTNumData numData = val.getNumRef().getNumCache();

                strData.setPtArray(null);  // unset old axis text
                numData.setPtArray(null);  // unset old values

                // set model
                int idx = 0;
                int rownum = 1;
                String ln;
                while((ln = modelReader.readLine()) != null){
                    String[] vals = ln.split("\\s+");
                    CTNumVal numVal = numData.addNewPt();
                    numVal.setIdx(idx);
                    numVal.setV(vals[1]);

                    CTStrVal sVal = strData.addNewPt();
                    sVal.setIdx(idx);
                    sVal.setV(vals[0]);

                    idx++;
                    XSSFRow row = sheet.createRow(rownum++);
                    row.createCell(0).setCellValue(vals[0]);
                    row.createCell(1).setCellValue(Double.valueOf(vals[1]));
                }
                numData.getPtCount().setVal(idx);
                strData.getPtCount().setVal(idx);

                String numDataRange = new CellRangeAddress(1, rownum-1, 1, 1).formatAsString(sheet.getSheetName(), true);
                val.getNumRef().setF(numDataRange);
                String axisDataRange = new CellRangeAddress(1, rownum-1, 0, 0).formatAsString(sheet.getSheetName(), true);
                cat.getStrRef().setF(axisDataRange);

                // 更新excel数据源
                OutputStream xlsOut = xlsPart.getPackagePart().getOutputStream();
                try {
                    wb.write(xlsOut);
                } finally {
                    xlsOut.close();
                }

                // 保存结果到新的PPT中
                OutputStream out = new FileOutputStream("output/pie-chart-demo-output.pptx");
                try {
                    pptx.write(out);
                } finally {
                    out.close();
                }
            } finally {
                wb.close();
            }
        } finally {
            modelReader.close();
        }
    }
}
