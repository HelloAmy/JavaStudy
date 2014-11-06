package zjr.amy.emil.test;

import java.util.Date;
import java.util.Properties;

import javax.mail.Authenticator;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.MimeMessage;

/**
 * 发邮件测试
 * @author zhujinrong
 *
 */
public class MailTest {
	
	static Authenticator auth = new Authenticator(){
		@Override
		protected PasswordAuthentication getPasswordAuthentication() {
			// TODO Auto-generated method stub
			return new PasswordAuthentication("1251759009@qq.com", "XY147258369");
		}
	};

	/**
	 * 主函数
	 * @param args
	 */
	public static void main(String[] args) throws MessagingException {
		// TODO Auto-generated method stub
		Properties props = new Properties();
		props.put("mail.smtp.host", "smtp.qq.com");
		props.put("mail.smtp.auth", "true");
		props.put("mail.from", "1251759009@qq.com");
		Session session = Session.getInstance(props, auth);
		MimeMessage msg = new MimeMessage(session);
		msg.setFrom();
		msg.setRecipients(Message.RecipientType.TO, "2804163771@qq.com");
		msg.setSubject("JavaMail hello world example");
		msg.setSentDate(new Date());
		msg.setText("<html><body><span style='color:red;'>Hello world!</span></body></html>", "utf-8", "html");
		Transport.send(msg);
	}
}
