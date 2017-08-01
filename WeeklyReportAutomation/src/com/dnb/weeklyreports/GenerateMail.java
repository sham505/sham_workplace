package com.dnb.weeklyreports;

import java.util.Properties;

import javax.mail.Address;
import javax.mail.Message;
import javax.mail.Multipart;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import org.apache.log4j.Logger;

public class GenerateMail {
	
	public static final Logger log = Logger.getLogger(GenerateMail.class);
	
	public static void sendMail(String mailBody,String subject,String filePath, String toAddress){
		
		try{
		log.info("sending report k"+subject);	
	 Properties Props = new Properties();
     Props.setProperty("mail.transport.protocol", "smtp");
     Props.setProperty("mail.smtp.host", "smtp-gw.us.dnb.com");
     
     // Initializing...
     Session MailSession = Session.getDefaultInstance(Props,null );
     Transport MailTransport = MailSession.getTransport();
     MimeMessage MailMessage = new MimeMessage(MailSession);                      
    
     // Getting Ready...
     
     //toAddress = "shamsheerc@dnb.com";
     mailBody  = mailBody + "<FONT FACE=ARIAL SIZE=2><BR><BR><B>NOTE: </B> This is an autogenerated mail. Please do not reply.</FONT>";
    
     
     MailMessage.setSubject(subject);
     MailMessage.setFrom(new InternetAddress("no-reply@dnb.com", "no-reply@dnb.com"));
     
     
     
     MimeBodyPart messageBodyPart = new MimeBodyPart();
     messageBodyPart.setContent(mailBody, "text/html");
     
     Multipart multipart = new MimeMultipart();
     multipart.addBodyPart(messageBodyPart);
     
     MimeBodyPart attachPart = new MimeBodyPart();
     
     attachPart.attachFile(filePath);

     multipart.addBodyPart(attachPart);
     
     MailMessage.setContent(multipart);
     
     String[] AddressList = toAddress.split(";",-1);
     Address[] ToAddresses = new InternetAddress[AddressList.length];
     
     for(int iTemp = 0;iTemp <=AddressList.length-1; iTemp++)
     {
                 String[] ToList = AddressList[iTemp].split(":",-1);
                 if(iTemp ==0)
                 {
                             //ToAddresses[iTemp] = new InternetAddress(ToList[0], ToList[1] + "(Primary Contact)");
                 	ToAddresses[iTemp] = new InternetAddress(ToList[0]);
                             MailMessage.addRecipient(Message.RecipientType.TO, ToAddresses[iTemp]);
                 }
                 else
                 {
                             ToAddresses[iTemp] = new InternetAddress(ToList[0]);
                             MailMessage.addRecipient(Message.RecipientType.TO, ToAddresses[iTemp]);
                             
                 }
     }
   //  SystemLogger.debug("Adding Receipents...", ToAddresses);
     // Sending Mail...
     MailTransport.connect();
    // SystemLogger.debug("Connected.", ToAddresses);
     MailTransport.sendMessage(MailMessage, MailMessage.getRecipients(Message.RecipientType.TO));
     MailTransport.close();
     log.info("report sent successfully");
     
}
catch(Exception Err)
{
	log.error("Exception in sending mail", Err);
Err.printStackTrace();

}
}	
}

