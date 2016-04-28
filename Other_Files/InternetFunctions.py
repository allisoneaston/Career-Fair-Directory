'''Has three functions: setup_internet, send_mail, kill internet'''

import smtplib
import getpass
class Mail_Connection:
	server = None 
	'''Sets up the umich smtp mail server to send auto email messages. Prompts for user login info.'''
	def setup_internet(self):
		self.server = smtplib.SMTP('smtp.mail.umich.edu')
		username = raw_input('Uniqname: ')
		password = getpass.getpass('Password: ')
		try:
			self.server.set_debuglevel(False)
			self.server.ehlo()
			if self.server.has_extn('STARTTLS'):
				self.server.starttls()
				self.server.ehlo()
			self.server.login(username,password)
		finally:
			return
	'''Quits the email server'''		
	def kill_internet(self):
		self.server.quit()
		return

	'''Sends an email from the setup server with names for to and from, addresses for to and from, subject and body text.
	Body text supports html'''	
	def send_mail(self,to_name,to_address,from_name,from_address,subject,body):
		message = """From: %s <%s>
MIME-Version: 1.0
Content-type: text/html
To: %s <%s>
Subject: %s

%s""" %(from_name,from_address,to_name,to_address,subject,body)
		self.server.sendmail(from_address,[to_address],message)
		
