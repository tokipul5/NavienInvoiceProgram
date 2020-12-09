import jxl.read.biff.BiffException;
import jxl.write.WriteException;
import org.apache.xmlbeans.XmlException;
import sun.management.snmp.jvminstr.JvmThreadInstanceEntryImpl;

import javax.swing.*;
import javax.swing.filechooser.FileSystemView;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;
import java.io.OutputStream;
import java.io.PrintStream;

public class Gui {

    private JTextArea consoleOutput =new JTextArea();
    private OutlookEmail email;
    private CreateInvoice createInvoice;
    public JTextArea getConsoleOutput() {
        return consoleOutput;
    }
    public void show() {
        JFrame f=new JFrame("Navien: Sending invoice program");//creating instance of JFrame
        f.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        //Subject
        final JTextArea subject =new JTextArea("Subject");
        JScrollPane scrollSubject = new JScrollPane(subject);
        scrollSubject.setBounds(400,20, 550,20);
        f.getContentPane().add(scrollSubject);
        f.setLayout(null);
        f.setVisible(true);

        //Email message
        final JTextArea message =new JTextArea("Type your message here.");
        JScrollPane scrollMessage = new JScrollPane(message);
        scrollMessage.setBounds(400,50, 550,200);
        f.getContentPane().add(scrollMessage);
        f.setLayout(null);
        f.setVisible(true);

        //Console output
        JScrollPane scrollConsole = new JScrollPane(this.consoleOutput);
        scrollConsole.setBounds(400,270, 550,350);
        f.getContentPane().add(scrollConsole);
        f.setLayout(null);
        f.setVisible(true);
        consoleOutput.append("testing\n");

        JLabel id = new JLabel("Outlook ID");
        id.setBounds(20, 20, 100, 20);
        JLabel password = new JLabel("Password");
        password.setBounds(20, 40, 100, 20);
        f.add(id);
        f.add(password);

        final JTextField textId = new JTextField();
        final JPasswordField textPassword = new JPasswordField();
        textId.setBounds(110, 20, 170, 20);
        textPassword.setBounds(110, 40, 170, 20);
        f.add(textId);
        f.add(textPassword);

        JButton loginButton = new JButton("Login");//creating instance of JButton
        loginButton.setBounds(290,20,70, 40);//x axis, y axis, width, height
        loginButton.addActionListener(new ActionListener(){
            public void actionPerformed(ActionEvent e){
                String strId = textId.getText();
                String strPassword = String.valueOf(textPassword.getPassword());
                email = new OutlookEmail(strId, strPassword, consoleOutput);
                email.login();
            }
        });
        f.add(loginButton);

        //Data file browser
        final JTextField dataFile = new JTextField("No file selected");
        dataFile.setBounds(20, 110, 340, 20);
        f.add(dataFile);

        JLabel dataFileText = new JLabel("Select the data file");
        dataFileText.setBounds(20, 80, 250, 20);
        f.add(dataFileText);

        JButton browseDataFile = new JButton("browse");
        browseDataFile.setBounds(260, 80, 100, 20);
        browseDataFile.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                // create an object of JFileChooser class
                JFileChooser j = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());

                // invoke the showsSaveDialog function to show the save dialog
                int r = j.showSaveDialog(null);

                if (r == JFileChooser.APPROVE_OPTION) {
                    // set the label to the path of the selected directory
                    dataFile.setText(j.getSelectedFile().getAbsolutePath());
                }
                // if the user cancelled the operation
                else
                    dataFile.setText("the user cancelled the operation");
            }
        });
        f.add(browseDataFile);
        // ------------------------------------------------------------------

        //Directory browser
        final JTextField directory = new JTextField("No file selected");
        directory.setBounds(20, 180, 340, 20);
        f.add(directory);

        JLabel directoryText = new JLabel("Select the directory you want to save");
        directoryText.setBounds(20, 150, 250, 20);
        f.add(directoryText);

        JButton browseDirectory = new JButton("browse");
        browseDirectory.setBounds(260, 150, 100, 20);
        browseDirectory.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                // create an object of JFileChooser class
                JFileChooser j = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());

                // set the selection mode to directories only
                j.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

                // invoke the showsSaveDialog function to show the save dialog
                int r = j.showSaveDialog(null);

                if (r == JFileChooser.APPROVE_OPTION) {
                    // set the label to the path of the selected directory
                    directory.setText(j.getSelectedFile().getAbsolutePath());
                }
                // if the user cancelled the operation
                else
                    directory.setText("the user cancelled the operation");
            }
        });
        f.add(browseDirectory);
        // ------------------------------------------------------------------

        JButton b=new JButton("Create and send invoices");//creating instance of JButton
        b.setBounds(20,210,200, 40);//x axis, y axis, width, height
        b.addActionListener(new ActionListener(){
            public void actionPerformed(ActionEvent e){
                String strId = textId.getText();
                String strPassword = String.valueOf(textPassword.getPassword());
                String dataPath = dataFile.getText();
                String savePath = directory.getText();
                String strSubject = subject.getText();
                String strBodyMessage = message.getText();
                try {
                    runProgram(dataPath, savePath, strId, strPassword, strSubject, strBodyMessage);
                } catch (BiffException biffException) {
                    biffException.printStackTrace();
                } catch (IOException ioException) {
                    ioException.printStackTrace();
                } catch (WriteException writeException) {
                    writeException.printStackTrace();
                } catch (XmlException xmlException) {
                    xmlException.printStackTrace();
                }
            }

            public void runProgram(String dataPath, String savePath, String id, String pw, String subject, String message) throws BiffException, IOException, WriteException, XmlException {
                long start = System.currentTimeMillis();

                //Create invoices
                createInvoice = new CreateInvoice(dataPath, savePath, id, pw);
                createInvoice.generateInvoice(consoleOutput);

                //Send emails
                System.out.println("Email");
                email.setSubject(subject);
                email.setBodyMessage(message);
                email.sendEmail(savePath);

                //Result
                consoleOutput.append("Created " + createInvoice.getCount() + " invoices.\n");
                consoleOutput.append("Sent " + email.getCount() + " emails.\n");

                long elapsedTime = System.currentTimeMillis() - start;
                System.out.print(elapsedTime/(60*1000F));
                System.out.println(" mins");
                consoleOutput.append(elapsedTime/(60*1000F) + " mins.");
            }
        });
        f.add(b);//create and send invoice button


        //Set size and format of JFrame.
        f.setSize(1000,700);//400 width and 500 height
        f.setLayout(null);//using no layout managers
        f.setVisible(true);//making the frame visible
    }
}
