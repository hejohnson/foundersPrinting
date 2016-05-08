import email, imaplib, os, subprocess, smtplib

detach_dir = "/home/ubuntu/foundersPrinting/files"

f = open('userConfig.txt', 'r')

user = f.readline().strip()
pwd = f.readline().strip()

def printFile(file, copies, sided):
	parts = file.split(".")
	print parts	
	if parts[1] == "doc" or parts[1] == "docx":
		os.system("libreoffice --headless --convert-to pdf --outdir "+detach_dir+" \""+detach_dir+"/"+file+"\"")
		os.system("rm \""+detach_dir+"/"+file+"\"")

	os.system("lp -d MG2500-series -o media=letter -o fit-to-page -n "+str(copies)+sided+" \""+detach_dir+"/"+parts[0]+".pdf\"")
	
	os.system("rm \""+detach_dir+"/"+parts[0]+".pdf\"")


smtp = smtplib.SMTP_SSL("smtp.gmail.com", 465)
smtp.login(user, pwd)
m = imaplib.IMAP4_SSL('imap.gmail.com')
m.login(user,pwd)
m.select('Inbox')
resp, items = m.search(None, "(BODY 'confirm')", "(UNSEEN)")
items = items[0].split()

for emailid in items: # Iterates through email messages
	copies = 1
	doubleOption = ""
	foundOptions = False
	resp, data = m.fetch(emailid, "(RFC822)")
	email_body = data[0][1]
	mail = email.message_from_string(email_body)
	origSender = mail["From"]
	if mail.get_content_maintype() != 'multipart': # If the message doesn't have multiple parts, it can't have an attachment
		continue
	
	for part in mail.walk(): # Step through each part of the message
		
		if part.get_content_maintype() == 'text' and not foundOptions: # If it's the plaintext part, the main body, check for any options
			print "body"
			foundOptions = True
			body = part.get_payload(decode=True).lower()
			if "copies" in body:
				words = body.split()
				copies = words[words.index("copies")-1]
				print copies
			if "double" in body:
				doubleOption = " -o sides=two-sided-long-edge"
			print doubleOption
			continue
		
		if part.get_content_maintype() == 'multipart': # if it's still a multipart, ignore it
			continue
		
		if part.get('Content-Disposition') is None:
			continue
	
		filename = part.get_filename()

		att_path = os.path.join(detach_dir, filename)

		if not os.path.isfile(att_path):
			fp = open(att_path, 'wb')
			fp.write(part.get_payload(decode=True))
			fp.close()
			smtp.sendmail("founders401printing@gmail.com", origSender, """From: Founders 401 Printing <founders401printing@gmail.com
										      To: You <"""+origSender+""">
										      Subject: File Received
										      Your file has been received and will be printed shortly""")
			printFile(filename, copies, doubleOption)

	m.store(emailid, '+FLAGS', '\Seen')
m.logout()

