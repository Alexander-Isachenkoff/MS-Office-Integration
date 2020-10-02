package office;

import javax.swing.*;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.io.File;

public class MainForm extends JFrame
{
    private final JTable table = new JTable();
    private final DefaultTableModel tableModel = new DefaultTableModel();
    private final JFileChooser fileChooser = new JFileChooser();
    
    
    public MainForm() {
        fileChooser.setCurrentDirectory(new File("").getAbsoluteFile());
        initTable();
        tableModel.setRowCount(3);
        tableModel.setColumnCount(4);
        
        JToolBar toolBar = new JToolBar();
        toolBar.setFloatable(false);
        JButton saveAsExcelButton = new JButton("Сохранить в Excel");
        toolBar.add(saveAsExcelButton);
        JButton saveAsWordButton = new JButton("Сохранить в Word");
        toolBar.add(saveAsWordButton);
        add(table, BorderLayout.CENTER);
        add(toolBar, BorderLayout.NORTH);
        
        saveAsWordButton.addActionListener(e -> {
            File file = getFileFromFileDialog(".docx");
            if (file != null) {
                new WordOutput().writeTable(getTableData(), file,
                        () -> showMessage("Файл " + file.getAbsolutePath() + " сохранён!"));
            }
            
        });
        saveAsExcelButton.addActionListener(e -> {
            File file = getFileFromFileDialog(".xlsx");
            if (file != null) {
                new ExcelOutput().writeTable(getTableData(), file,
                        () -> showMessage("Файл " + file.getAbsolutePath() + " сохранён!"));
            }
        });
        
        setPreferredSize(new Dimension(640, 480));
        setDefaultCloseOperation(DISPOSE_ON_CLOSE);
    }
    
    private File getFileFromFileDialog(String ext) {
        fileChooser.setSelectedFile(null);
        fileChooser.showSaveDialog(this);
        
        File file = fileChooser.getSelectedFile();
        if (file != null) {
            String path = fileChooser.getSelectedFile().getAbsolutePath();
            if (!path.endsWith(ext)) {
                path += ext;
            }
            return new File(path);
        } else {
            return null;
        }
    }
    
    private void showMessage(String message) {
        JOptionPane.showMessageDialog(this, message);
    }
    
    private void initTable() {
        table.setSelectionBackground(new Color(166, 223, 236));
        table.setRowHeight(24);
        table.setFont(table.getFont().deriveFont(14.721377f));
        table.setSelectionForeground(Color.BLACK);
        table.setGridColor(new Color(48, 48, 48));
        DefaultTableCellRenderer centerRend = new DefaultTableCellRenderer();
        centerRend.setHorizontalAlignment(JLabel.CENTER);
        table.setDefaultRenderer(Object.class, centerRend);
        table.setModel(tableModel);
    }
    
    private String[][] getTableData() {
        String[][] data = new String[tableModel.getRowCount()][];
        for (int i = 0; i < tableModel.getRowCount(); i++) {
            data[i] = new String[tableModel.getColumnCount()];
            for (int j = 0; j < tableModel.getColumnCount(); j++) {
                Object value = tableModel.getValueAt(i, j);
                data[i][j] = (value != null) ? value.toString() : "";
            }
        }
        return data;
    }
}
