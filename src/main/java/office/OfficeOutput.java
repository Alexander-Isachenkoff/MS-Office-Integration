package office;

import org.apache.poi.POIXMLDocument;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public abstract class OfficeOutput
{
    public void writeTable(String[][] data, File file, Action onSave) {
        POIXMLDocument document = createDocument(data);
        try {
            write(document, file);
            onSave.performAction();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    
    abstract POIXMLDocument createDocument(String[][] data);
    
    private void write(POIXMLDocument document, File file) throws IOException {
        try (OutputStream os = new FileOutputStream(file)) {
            document.write(os);
        }
    }
    
    public interface Action
    {
        void performAction();
    }
}
