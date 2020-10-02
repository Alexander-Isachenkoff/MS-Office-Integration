package office;

import javax.swing.*;

public class Main
{
    public static void main(String[] args) {
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (ClassNotFoundException | InstantiationException | IllegalAccessException | UnsupportedLookAndFeelException e) {
            e.printStackTrace();
        }
        
        JFrame frame = new MainForm();
        frame.pack();
        frame.setLocationRelativeTo(null);
        frame.setVisible(true);
    }
}