package office;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.xssf.usermodel.*;

public class ExcelOutput extends OfficeOutput
{
    @Override
    POIXMLDocument createDocument(String[][] data) {
        XSSFWorkbook book = new XSSFWorkbook();
        XSSFSheet sheet = book.createSheet();
        
        for (int i = 0; i < data.length; i++) {
            XSSFRow row = sheet.createRow(i);
            for (int j = 0; j < data[i].length; j++) {
                try {
                    double num = Double.parseDouble(data[i][j].replace(',', '.'));
                    row.createCell(j).setCellValue(num);
                    XSSFDataFormat format = book.createDataFormat();
                    XSSFCellStyle style = book.createCellStyle();
                    style.setDataFormat(format.getFormat("0.0"));
                    row.getCell(j).setCellStyle(style);
                } catch (NumberFormatException ex) {
                    row.createCell(j).setCellValue(data[i][j]);
                }
            }
        }
        return book;
    }
}
