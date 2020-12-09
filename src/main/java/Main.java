import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.WriteException;
import org.apache.xmlbeans.XmlException;

import javax.sound.midi.Track;
import javax.swing.*;
import javax.swing.filechooser.FileSystemView;
import javax.swing.text.SimpleAttributeSet;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.io.PrintStream;
import javax.swing.text.StyleConstants;

public class Main {
    public static void main(String []args) throws IOException, BiffException, XmlException {
        Gui gui = new Gui();
        gui.show();
    }
}
