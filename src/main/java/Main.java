import jxl.read.biff.BiffException;
import org.apache.xmlbeans.XmlException;
import javax.swing.*;
import java.io.IOException;

public class Main {
    public static void main(String []args) throws IOException, BiffException, XmlException {
        try {
            Gui gui = new Gui();
            gui.show();
        } catch (Exception e) {
            JOptionPane.showMessageDialog(null, "Error", "InfoBox: Close Excel files", JOptionPane.INFORMATION_MESSAGE);
        }
    }
}
