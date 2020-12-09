import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Properties;
import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.*;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import javax.swing.*;

public class OutlookEmail {
    private String id;
    private String pw;
    private Session session;
    private String subject;
    private String bodyMessage;
    private String pathTrack;
    private JTextArea textArea;
    private int count = 0;
    public OutlookEmail(String id, String pw, JTextArea textArea) {
        this.id = id;
        this.pw = pw;
        this.textArea = textArea;
        final String username = id;
        final String password = pw;
        Properties props = new Properties();
        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.starttls.enable", "true");
        props.put("mail.smtp.host", "outlook.office365.com");
        props.put("mail.smtp.port", "587");

        session = Session.getInstance(props,
                new javax.mail.Authenticator() {
                    protected PasswordAuthentication getPasswordAuthentication() {
                        return new PasswordAuthentication(username, password);
                    }
                });
    }

    public void setSubject(String subject) {
        this.subject = subject;
    }

    public void setBodyMessage(String bodyMessage) {
        this.bodyMessage = bodyMessage;
    }

    public String getSubject() {
        return subject;
    }

    public String getBodyMessage() {
        return bodyMessage;
    }

    public int getCount() {
        return count;
    }

    public void sendEmail(String pathSave) throws IOException, BiffException {
        this.pathTrack = pathSave + "\\" + "trackMail.xlsx";
        //obtaining input bytes from a file
        FileInputStream input = new FileInputStream(new File(pathTrack));
        //creating workbook instance that refers to .xls file
        XSSFWorkbook workbook = new XSSFWorkbook(input);
        //creating a Sheet object to retrieve the object
        XSSFSheet sheet = workbook.getSheetAt(0);
        for(int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
            Row row = sheet.getRow(i);
            if (row.getCell(0) == null) {
                break;
            } else if (row.getCell(4) != null) {
                continue;
            }
            String po = row.getCell(0).getStringCellValue();
            String email = row.getCell(1).getStringCellValue();
            String filePath = row.getCell(2).getStringCellValue();
            String fileName = row.getCell(3).getStringCellValue();
            System.out.println(po + " " + email + " " + filePath + " " + fileName);
            try {
                if (email.equals(" ")) {
                    throw new Exception();
                }
                sendToBuyer(email, filePath, fileName);
                Cell cellSent = row.createCell(4);
                cellSent.setCellValue("Yes");
                count++;
            } catch (Exception e) {
                textArea.append(po + " " + "has no email of the buyer.\n");
            }
        }
        FileOutputStream fileOutputStream = new FileOutputStream(pathTrack);
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        workbook.close();
    }

    public void login() {
        try {
            Message message = new MimeMessage(session);
            message.setFrom(new InternetAddress(id));
            message.setRecipients(Message.RecipientType.TO,
                    InternetAddress.parse(id));
            message.setSubject("Login Successfully in Navien Invoice Program");
            //message.setText("HI");

            // Create the message part
            BodyPart messageBodyPart = new MimeBodyPart();

            // Now set the actual message
            messageBodyPart.setText("Please change your password if you did not use the program.");

            // Create a multipart message
            Multipart multipart = new MimeMultipart();

            // Set text message part
            multipart.addBodyPart(messageBodyPart);

            // Send the complete message parts
            message.setContent(multipart);

            Transport.send(message);
        } catch (MessagingException e) {
            textArea.append("Please check your username and password.\n");
            throw new RuntimeException("Wrong username or password");
        }
        textArea.append("Login in successfully!\n");
    }

    public void sendToBuyer(String buyerEmail, String pathFile, String pdfName) {
        try {
            Message message = new MimeMessage(session);
            message.setFrom(new InternetAddress(id));
            message.setRecipients(Message.RecipientType.TO,
                    InternetAddress.parse(buyerEmail));
            message.setSubject(subject);
            //message.setText("HI");

            // Create the message part
            BodyPart messageBodyPart = new MimeBodyPart();

            // Now set the actual message
            messageBodyPart.setText(bodyMessage);

            // Create a multipart message
            Multipart multipart = new MimeMultipart();

            // Set text message part
            multipart.addBodyPart(messageBodyPart);

            // Part two is attachment
            messageBodyPart = new MimeBodyPart();
            String filename = pathFile;
            DataSource source = new FileDataSource(filename);
            messageBodyPart.setDataHandler(new DataHandler(source));
            messageBodyPart.setFileName(pdfName);
            multipart.addBodyPart(messageBodyPart);

            // Send the complete message parts
            message.setContent(multipart);

            Transport.send(message);

            textArea.append("Sent " + pdfName + "\n");

        } catch (MessagingException e) {
            throw new RuntimeException("Failed to send an email");
        }
    }
}
