package office;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class WordOutput extends OfficeOutput
{
    @Override
    POIXMLDocument createDocument(String[][] data) {
        XWPFDocument document = new XWPFDocument();
        XWPFTable table = document.createTable(data.length, data[0].length);
        for (int i = 0; i < data.length; i++) {
            XWPFTableRow row = table.getRow(i);
            for (int j = 0; j < data[i].length; j++) {
                row.getCell(j).setText(data[i][j]);
            }
        }
        return document;
    }
}
