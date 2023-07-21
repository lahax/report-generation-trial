package emailService;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Properties;
import javax.mail.Authenticator;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class EmailWithExcelAttachment {
    public static void main(String[] args) {
        // Crea il file Excel
        Workbook workbook = new XSSFWorkbook();
        // Aggiungi dati al file Excel (opzionale)
        // ...

        // Salva il file Excel all'interno del container
        String fileName = "file.xlsx";
        try (FileOutputStream outputStream = new FileOutputStream(fileName)) {
            workbook.write(outputStream);
        } catch (Exception e) {
            e.printStackTrace();
        }

        // Configura le proprietà per la sessione di JavaMail
        Properties properties = new Properties();
        properties.put("mail.smtp.host", "your_smtp_host");
        properties.put("mail.smtp.port", "your_smtp_port");
        properties.put("mail.smtp.auth", "true");
        properties.put("mail.smtp.starttls.enable", "true");

        // Credenziali di accesso al tuo account email
        final String username = "your_email@example.com";
        final String password = "your_email_password";

        // Configura l'autenticazione per la sessione di JavaMail
        Authenticator authenticator = new Authenticator() {
            @Override
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(username, password);
            }
        };

        // Crea una nuova sessione di JavaMail
        Session session = Session.getInstance(properties, authenticator);

        try {
            // Crea un nuovo messaggio di posta
            Message message = new MimeMessage(session);

            // Imposta il mittente e i destinatari
            message.setFrom(new InternetAddress("your_email@example.com"));
            message.setRecipients(Message.RecipientType.TO, InternetAddress.parse("recipient@example.com"));

            // Imposta l'oggetto del messaggio
            message.setSubject("Oggetto della email");

            // Crea il corpo del messaggio
            MimeBodyPart messageBodyPart = new MimeBodyPart();
            messageBodyPart.setText("Contenuto del messaggio");

            // Crea il multipart per allegare il file Excel
            MimeMultipart multipart = new MimeMultipart();
            multipart.addBodyPart(messageBodyPart);

            // Aggiungi il file Excel come allegato
            MimeBodyPart attachPart = new MimeBodyPart();
            attachPart.attachFile(new File(fileName));
            multipart.addBodyPart(attachPart);

            // Imposta il multipart come contenuto del messaggio
            message.setContent(multipart);

            // Invia il messaggio
            Transport.send(message);

            System.out.println("Email inviata con successo!");
        } catch (MessagingException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}