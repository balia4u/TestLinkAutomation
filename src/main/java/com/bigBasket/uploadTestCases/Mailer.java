package com.bigBasket.uploadTestCases;


import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.BodyPart;
import javax.mail.Message.RecipientType;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.Session;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow.Builder;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.gmail.Gmail;
import com.google.api.services.gmail.GmailScopes;
import com.google.api.services.gmail.model.Message;

public class Mailer {
	/** Application name. */
	private static final String APPLICATION_NAME = "Gmail API Java Quickstart";

	/** Directory to store user credentials for this application. */
	private static final File DATA_STORE_DIR = new File(System.getProperty("user.dir")+"//src//main//resources",
			".credentials//gmail-java-quickstart");

	/** Global instance of the {@link FileDataStoreFactory}. */
	private static FileDataStoreFactory DATA_STORE_FACTORY;

	/** Global instance of the JSON factory. */
	private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();

	/** Global instance of the HTTP transport. */
	private static HttpTransport HTTP_TRANSPORT;

	/**
	 * Global instance of the scopes required by this quickstart.
	 *
	 * If modifying these scopes, delete your previously saved credentials at
	 * ~/.credentials/gmail-java-quickstart
	 */
	private static final List<String> SCOPES = Arrays.asList(GmailScopes.GMAIL_SEND, GmailScopes.GMAIL_COMPOSE);

	static {
		try {
			HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
			DATA_STORE_FACTORY = new FileDataStoreFactory(DATA_STORE_DIR);
		} catch (Throwable t) {
			t.printStackTrace();
			System.exit(1);
		}
	}

	/**
	 * Creates an authorized Credential object.
	 * 
	 * @return an authorized Credential object.
	 * @throws IOException
	 * @author rakesh
	 */
	static Credential authorize() throws IOException {
		InputStream inputStream = new FileInputStream(System.getProperty("user.dir") + "//src/main/resources//client_secret.json");
		GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(inputStream));

		GoogleAuthorizationCodeFlow flow = new Builder(HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
				.setDataStoreFactory(DATA_STORE_FACTORY).setAccessType("offline").build();
		Credential credential = new AuthorizationCodeInstalledApp(flow, new LocalServerReceiver()).authorize("user");
		System.out.println("Credentials saved to " + DATA_STORE_DIR.getAbsolutePath());
		return credential;
	}

	/**
	 * Build and return an authorized Gmail client service.
	 * 
	 * @return an authorized Gmail client service
	 * @throws IOException
	 * @author rakesh
	 */
	static Gmail getGmailService() throws IOException {
		Credential credential = authorize();
		return new Gmail.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential).setApplicationName(APPLICATION_NAME).build();
	}

	/**
	 * Create a MimeMessage using the parameters provided.
	 *
	 * @param to
	 *            email address of the receiver
	 * @param from
	 *            email address of the sender, the mailbox account
	 * @param subject
	 *            subject of the email
	 * @param bodyText
	 *            body text of the email
	 * @return the MimeMessage to be used to send email
	 * @throws javax.mail.MessagingException
	 *             throws whatever the exception is.
	 * @author rakesh
	 */
	static MimeMessage createEmail(String[] to, String from, String subject, ArrayList<MimeBodyPart> mimeBodyList)
			throws MessagingException {
		Properties props = new Properties();
		Session session = Session.getDefaultInstance(props, null);

		MimeMessage email = new MimeMessage(session);

		email.setFrom(new InternetAddress(from));
		for (int i = 0; i < to.length; i++) {
			email.addRecipients(RecipientType.TO, InternetAddress.parse(to[i]));
		}
		email.setSubject(subject);

		Multipart multipart = new MimeMultipart();
		for (MimeBodyPart part : mimeBodyList) {
			part.setDisposition(MimeBodyPart.INLINE);
			multipart.addBodyPart(part);
		}
		email.setContent(multipart);
		return email;
	}

	/**
	 * Create a message from an email.
	 *
	 * @param emailContent
	 *            Email to be set to raw of message
	 * @return a message containing a base64url encoded email
	 * @throws IOException
	 * @throws MessagingException
	 * @author rakesh
	 */
	static Message createMessageWithEmail(MimeMessage emailContent) throws MessagingException, IOException {
		ByteArrayOutputStream buffer = new ByteArrayOutputStream();
		emailContent.writeTo(buffer);
		byte[] bytes = buffer.toByteArray();
		String encodedEmail = com.google.api.client.util.Base64.encodeBase64URLSafeString(bytes);
		Message message = new Message();
		message.setRaw(encodedEmail);
		return message;
	}

	/**
	 * To add attachment with a mail.
	 * @param multipart
	 * @param filename
	 * @throws MessagingException
	 * @author rakesh
	 */
	private static void addAttachment(Multipart multipart, String filename) throws MessagingException {
		DataSource source = new FileDataSource(filename);
		BodyPart messageBodyPart = new MimeBodyPart();
		messageBodyPart.setDataHandler(new DataHandler(source));
		messageBodyPart.setFileName(filename);
		multipart.addBodyPart(messageBodyPart);
	}

	/**
	 * Used to send emails.
	 * @param to
	 * @param subject
	 * @param mimeBodyParts
	 * @throws MessagingException
	 * @throws IOException
	 */
	public void sendMail(String[] to, String subject, ArrayList<MimeBodyPart> mimeBodyParts)
			throws MessagingException, IOException {
		System.out.println("line 1");
		Gmail service = Mailer.getGmailService();
		System.out.println("line 2");
		final String emailAddress = "bbqabot@bigbasket.com";
		Message email = Mailer.createMessageWithEmail(Mailer.createEmail(to, emailAddress, subject, mimeBodyParts));
		System.out.println("line 3");
		service.users().messages().send(emailAddress, email).execute();
		System.out.println("line 4");
	}
}
