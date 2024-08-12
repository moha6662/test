from __future__ import print_function
import sqlite3
import os, smtplib, ssl, email, re, csv
import publicdomain, adstaff
from email.mime.application import MIMEApplication
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from os.path import basename
from getpass import getpass
from os.path import exists
from time import sleep
from datetime import date


thisday = date.today().strftime("%d-%m-%Y")
domainlist = publicdomain.domainlist
staff= adstaff.staff
domain_users_emailkey = {}
domain_users_e0key = {}
main_manager = 'E08430' # for those who don't have manager 
with open("Domain_users.csv" , "r") as reader:
	reader_ = csv.reader(reader)
	for email in reader_:
		domain_users_emailkey[email[0].upper()] = email[2].upper()
		domain_users_e0key[email[2].upper()] = email[0].upper()
			
whitelist = ['adib.ae', 'adib.com',  'adib.eg', 'hima205088@gmail.com', 'ahmed-habashi@hotmail.com', 'abdelrhmen10@hotmail.com', 'khedr.group@gmail.com', 'tamerkhedr52@gmail.com','mohamed.fathy.alfa@gmail.com', 
        'generaltopspeed2010@yahoo.com', 'radiint@gmail.com', 'a.radi@radi-international.com', 'amr@radi-international.com', 'm.sakr@radi-international.com',
        'intoffic2011@yahoo.com', 'almoshreqhr@gmail.com', 'dr.osamaegyptfarm@gmail.com', 'operations_tiba2000@yahoo.com', 'intadn@yahoo.com', 'moh_gr2d@hotmail.com',
        'Leaders-co@hotmail.com', 'adibleaders.co@gmail.com', 'Alexleaders.co@gmail.com', 'khoshala11@hotmail.com', 'waleedgomaa@gmail.com',
        'EG_', 'gebreilco.com','ahmedsaadabomomen@gmail.com', 'innovation-dev.com', 'ahmedshokry33@yahoo.com', 'm.elhandasia@gmail.com',
        'acc.tharwat.plast@gmail.com', 'mohd.fouda89@gmail.com', 'i.elhadary@oilex.com', 'ibrahim_alhadary@yahoo.com', 'OC00002', 'E04907', 'E04341', '902344', 'E04063',
        'E04386', 'E05667', 'A003', 'E03657', 'E05664', 'Imran Ibrahim', 'E01118', 'E03721', 'Amr adel halim samy', 'E04806', 'E02850', 'Wafaa el Dars',
        'seiflimousine@yahoo.com', 'akhbarart@gmail.com', 'nakib.adv@gmail.com', 'a.yacoub30@gmail.com', 'mohamed.salama0@gmail.com', 'islamemam8310@gmail.com', 
        'emanaliyones@gmail.com', 'yasmin.hamdi@gmail.com', 'medhat.yaacoub@gmail.com', 'mohamedkhamees79@gmail.com','akhbar_art@yahoo.com',
        'ahmedgamal8687@gmail.com', 'ngz11_ahram@hotmail.com', 'adhm77@hotmail.com', 'gazali-wf@hotmail.com', 'afma67@hotmail.com', 'medhat333@hotmail.com',
        'aleqaria@yahoo.com', 'cebdaa@yahoo.com', 'mhmd_khamees@yahoo.com', 'akhbar_art@yahoo.com', 'magedmjaher@yahoo.com', 'Ashraf_elleithy2002@yahoo.com',
        'almal.computer@yahoo.com', 'm.hendy75@yahoo.com', 'shere-mohamed@hotmail.com', 'mohamedaboelfotoh35@gmail.com', 'doaaselim.ds@gmail.com',
        'reemabdelmeguid@gmail.com ', 'edugateegypt@gmail.com ', 'sarah.iskandar@hotmail.com', 'abughuddah@gmail.com', 'drmhmdomar42@yahoo.com', 'egyptadib@gmail.com',
        'hanan2raouf@gmail.com', 'reham.h.fadl@gmail.com', 'alarakyr@yahoo.com', 'aer.ahmed@gmail.com']

def creat_manager_csvs(holder):
	master_db = sqlite3.connect('.\\emails.db')
	master_db.isolation_level = None
	nave = master_db.cursor()
	finished = nave.execute('select csv_finished from Configuration')
	for  f in finished:
		try:
			if int(f[0]) == 1:
				return
		except Exception as e:
			print(str(e))
	print('[+] Creating managers CSV files')
	result = nave.execute('select id,sent_date,username,subject,file_name,email_address,recipient,manager from Emails where id > '+str(holder)+' order by id ASC')
	last_id = 0
	try:
		for rr in result:
			if os.path.isfile('.\\managers\\'+rr[7]+'.csv') == False:
				with open('.\\managers\\'+rr[7]+'.csv', 'ab+') as mancsv:
					mancsv.write(u'\ufeff'.encode('utf8'))
					mancsvr = csv.writer(mancsv)
					mancsvr.writerow(["EndTime", "User Name", "Subject", "File Name", "Sender", "recipient"])
					mancsvr.writerow([unicode(s).encode("utf-8") for s in rr[1:-1]])
					last_id = rr[0]
			else:
				with open('.\\managers\\'+rr[7]+'.csv', 'ab+') as mancsv:
					mancsvr = csv.writer(mancsv)
					mancsvr.writerow([unicode(s).encode("utf-8") for s in rr[1:-1]])
					master_db = sqlite3.connect('.\\emails.db')
					last_id = rr[0]
		nave.execute('update Configuration set holder_id = '+str(rr[0]))
		nave.execute('update Configuration set csv_finished = 1')
		master_db.commit()
		master_db.close()
	except Exception as e:
		nave.execute('update Configuration set holder_id = '+str(last_id))
		master_db.commit()
		master_db.close()
		print(str(e))
		
def create_main_table():
	holder = 0
	if os.path.isfile('.\\emails.db') == False:
		print('[+] Creating master database')
		master_db = sqlite3.connect('.\\emails.db')
		nave = master_db.cursor()
		nave.execute('create table Emails (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT NOT NULL, email_address TEXT NOT NULL, sent_date TEXT NOT NULL, subject TEXT NOT NULL, file_name TEXT, recipient TEXT NOT NULL, manager TEXT)')
		nave.execute('create table Configuration (holder_id INTEGER, emails_finished INTEGER, csv_finished INTEGER, last_manager_sent TEXT)')
		master_db.commit()
		master_db.close()
		master_db = sqlite3.connect('.\\emails.db')
		nave = master_db.cursor()
		nave.execute('insert into Configuration (emails_finished) values(0)')
		nave.execute('update Configuration set csv_finished = 0')
		master_db.commit()
		master_db.close()
		maincsv = open('Forcepoint - Email Protection.csv', 'r')
		lencsv = sum(1 for line in maincsv)
		maincsv.close()
		maincsv = open('Forcepoint - Email Protection.csv', 'r')
		progress = 0
		mainreader = csv.reader(maincsv)
		mainreader.next()
		master_db = sqlite3.connect('.\\emails.db')
		master_db.text_factory = str
		nave = master_db.cursor()
		for row in mainreader:
			public = False
			for wl in whitelist:
				if wl.upper() in row[4].upper():
					continue
			for public_domain in domainlist:
				if public_domain.upper() in row[4].upper():
					public = True
					#print(row[4].upper())
					break
			if public == False:
				continue
			try:
				username = domain_users_emailkey[row[1].upper()]
			except:
				username = 'Group_Email'
			if username == 'Group_Email':
				manager = main_manager
			else:
				for k,v in staff.items():
					if username in v:
						manager = k
						break
					manager = main_manager 
			if row[3] == '||||||':
				filename = ''
			else:
				filename = row[3]
			sqlq = u'insert into Emails (username,email_address,sent_date,subject,file_name,recipient,manager) values(?,?,?,?,?,?,?)'
			params = [username,row[1],row[0],row[2],filename,row[4],manager]
			nave.execute(sqlq,params)
			progress2 = (float(progress) / float(lencsv)) * 1000
			if progress2 >= 100.0:
				progress2 = 100.0
			print("Progress: {:.1f}%".format(progress2), end='\r')
			progress += 1
		nave.execute('update Configuration set holder_id = 0')
		nave.execute('update Configuration set emails_finished = 1')
		master_db.commit()
		master_db.close()
		maincsv.close()
		creat_manager_csvs(holder)
	else:
		master2_db = sqlite3.connect('.\\emails.db')
		nave2 = master2_db.cursor()
		diditfinish = nave2.execute('select holder_id, emails_finished from Configuration')
		for ff in diditfinish:
			if int(ff[1]) == 1:
				creat_manager_csvs(ff[0])
				master2_db.close()
				return
			else:
				master2_db.close()
				print("[!] Script was interrupted in the last run!")
				os.remove(".\\emails.db")
				create_main_table()
		master2_db.close()
		
def send_csv_by_mail():
	managers_csv_list = os.listdir('.\\managers')
	port = 587
	username = raw_input("[+]Enter your username: ")
	emailuser = "ADIBEG\\"+username
	password = getpass()
	smtp_server = "10.240.48.36"
	for i in managers_csv_list:
		man = i.split('.')[0].upper()
		try:
			man_name_mail = domain_users_e0key[man]
			man_name = man_name_mail.split('.')[0]
		except:
			if "L" in man.upper() and len(man) == 3:
				man2 = man.upper().split('L')
				man2.insert(0, 'L0')
				man2.remove('')
				man2 = ''.join(man2)
				man_name_mail = domain_users_e0key[man2]
				man_name = man_name_mail.split('.')[0]
        	try:
			if len(man_name) == 1:
				try:
					man_name = man_name_mail.split('.')[1].split('@')[0]
				except:
					man_name = man_name_mail('.')[1].split('@')[0]
		except Exception as e:
			#print(e)
			print("[!]Error, Kindly check: "+man+" it maybe deleted from the AD")
			os.close(0)
		if man == main_manager:
			print("[+]Sending Users without Managers to Mai and CC:[SNOC Manager, Eng Barakat]")
			body = "Dear Mai,<br><br>Please find the attached report for users with no managers list of mails<br><br>Thanks<br>Best of luck<br>Security Operation Center Team"
			to = man_name_mail
			ccc = "egthreatintel@adib.eg,Ahmed.Barakat@adib.eg"
		else:
			body = "Dear "+man_name+",<br><br>Kindly note that as per ADIB Egypt Policy all emails sent externally are monitored.<br><br>Being a partner to our success in safeguarding ADIB Informational assets, kindly find attached all mails sent externally from your team for the last week.<br><br>Would you please review it and for any comments please refer to Information Security team (Eg_SecurityIncident@adib.eg)<br><br>Thank you in advance for your understanding and cooperation.\n\nFor any enquiries please do not hesitate to contact information security team.\n\n"+ """<style> {font-family: Calibri, 'Times New Roman', Times, serif;}
                               .sig1 {color: #1F497D;}
                               .sig2 {color: #7F7F7F;}</style>
                               <font class='sig1'><br><br>Best Regards<br>Security Operation Center<br></font>
                               <font class='sig2'>Information Security Department</font></body>"""
			if man == "E04458":
				print("[+]Eng Barakat - CC [EG.ADIB.InfoSecMonitoring@adib.eg] only")
				to = man_name_mail
				ccc = "EG.ADIB.InfoSecMonitoring@adib.eg,EG.ADIB.InfoSecMonitoring@adib.eg"
			else:
				to = man_name_mail
				ccc = "EG.ADIB.InfoSecMonitoring@adib.eg,Eg_SecurityIncident@adib.eg"
                		#to = 'mohamed.nagy@adib.eg'
				#ccc = 'mohamed.nagy@adib.eg,mohamed.nagy@adib.eg'
		rcpt = ccc.split(",") + [to]
		message = MIMEMultipart()
		message["From"] = "EG.ADIB.InfoSecMonitoring@adib.eg"
		#message["From"] = "mohamed.nagy@adib.eg"
		#message["To"] = "mohamed.nagy@adib.eg"
		message["To"] = to
		message["Subject"] = man + " - Mails Sent to External Mail by your team for review - "+ thisday
		message["Cc"] = ccc
		message.attach(MIMEText(body, "html"))
		attachmentfile = ".\\managers\\"+man+'.csv'
		with open(attachmentfile, 'rb') as attachment:
			part = MIMEApplication(attachment.read(), Name=basename(attachmentfile))
			part.add_header("Content-Disposition","attachment; filename=%s" % basename(attachmentfile))
			message.attach(part)
			text = message.as_string()
		retry = True
		while retry:
			try:
				server = smtplib.SMTP(smtp_server,port)
				server.starttls() # Secure the connection
				server.login(emailuser, password)
				server.sendmail(message["From"], rcpt, text)
				print('[+]Mail sent to: '+man)
				retry = False
				os.remove('.\\managers\\'+man+'.csv')
				server.quit()
				sleep(2)
			except Exception as e:
				if "Authentication unsuccessful" in str(e):
					print("[!]Wrong username or password!")
					exit()
				#print(str(e))
				print("[!]Connection Error, trying again... please wait!")
				sleep(30)

def delete_after_finish():
	os.remove('.\\emails.db')
	
create_main_table()
send_csv_by_mail()
delete_after_finish()