#!/usr/bin/python
#encoding: utf-8
from datetime import date, timedelta
#from time import sleep
import os
import re
import getopt
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

from smb.SMBConnection import SMBConnection
import imaplib, poplib
import email

def prev_weekday(adate):
	_offsets = (3, 1, 1, 1, 1, 1, 2)
	return adate - timedelta(days=_offsets[adate.weekday()])

def smb_copy(smb, tfile, sfile, wd):
	tar_path = tfile.split(":")[0]
	tar_file = tfile.split(":")[1]
	conn=SMBConnection(smb['username'], smb['password'],"", smb['host'], domain=smb['domain'], use_ntlm_v2 = True, is_direct_tcp=True)
	try:
#		print(u"登陆SMB服务器")
		result = conn.connect(smb['host'], port=445, timeout=30) #smb协议默认端口445
		if result :
			print(u"%s登陆SMB成功：%s" %(smb['host'],result))
	except Exception as e:
		print(u"!!! %s登陆SMB报错：%s !!!" % (smb['host'],e))
		sys.exit(1)
	try:
		local_file=open(sfile, "r")
		conn.storeFile(tar_path , tar_file , local_file)
		print("上传《%s》完成" % tfile)
		local_file.close()
	except Exception as e :
		print(u"\n!!! 拷贝文件时出错 !!!\n!!! %s\n" % (str(e).split("\n")[0]))
	return 

def get_config(cfg_file, ddate):
	cfg = dict()
	#处理日期
	#date_YYYY=datetime.date.today().strftime("%Y")
	#date_MM=datetime.date.today().strftime("%m")
	#date_DD=datetime.date.today().strftime("%d")
	date_YYYY=ddate[0:4]
	date_MM=ddate[4:6]
	date_DD=ddate[6:8]
	for line in open(cfg_file,"r") :
		line = line.decode("utf-8").replace(" ", "").replace("\n", "").strip()
		if len(line) > 0 :
			if line[0] != "#" and line[0] != "'":
				if line[0]=="[":
					target = line.replace("[","").replace("]","")
					cfg.update({target:{}})
				else:
					detail=line.split("#")[0].strip()
					detail=detail.split("=")
					if len(detail) > 1:
						detail[0] = detail[0].strip()
						detail[1] = detail[1].replace("${YYYYMMDD}", ddate).replace("${YYYY}", date_YYYY).replace("${MM}", date_MM).replace("${DD}", date_DD).strip()
						if detail[0] not in cfg[target] :
							cfg[target].update({detail[0]:detail[1]})
						else :
							cfg[target][detail[0]] = detail[1]
#	print("读取配置完成 %s" % cfg)
	return cfg
	
def savefile(filename, data, path):#保存文件方法（保存在path目录下）
	try:
		filepath = path + u'/'+ filename
		print(filepath)
		fn = open(filepath, 'wb')
	except:
		print('filename error')
		fn.close()
	fn.write(data)
	fn.close()

def get_email(ecfg):
	global local_path
	result = []
	try :
		emode = ecfg['email_mode'].lower()
	except :
		emode = "imap"
	if emode == "imap" :
		#处理imap模式
		try:
			a=imaplib.IMAP4(host = ecfg['email_server'], port = int(ecfg['email_port']) )
			a.login(ecfg['email_user'], ecfg['email_pwd'])
			a.utf8_enabled = True
			a.select(mailbox='INBOX', readonly=True)
			filter=[]
			if "email_sender" in ecfg:
				filter.append('FROM')
				filter.append(ecfg['email_sender'])
			if "email_subject" in ecfg:
				filter.append("SUBJECT")
				filter.append(ecfg['email_subject'])
			status, data = a.search("UTF-8", *filter)
			print(data)
			if status != 'OK':
				raise Exception('读取邮件发生错误')
			for num in data[0].split():
				type_, data = a.fetch(num,'(RFC822)')
				msg = email.message_from_string(data[0][1].decode('utf-8'))#传输邮件全部内容，用email解析
				From_mail = email.utils.parseaddr(msg.get('from'))[1]
				From_mail_name = From_mail.split('@')[0]
				mail_title , mail_charset = email.Header.decode_header(msg.get('Subject'))[0]
				mail_title = mail_title.decode(mail_charset)
				print(mail_title)
				for part in msg.walk():
					if not part.is_multipart():
						filename = part.get_filename() #如果是附件，这里就会取出附件的文件名
						if filename:
							fname,file_charset = email.Header.decode_header(filename)[0]
							if file_charset :
								fname = fname.decode(file_charset)
							attach_data = part.get_payload(decode=True)
							print('下载文件 %s (长度 %i)' % (fname, len(attach_data)))
							savefile(fname, attach_data, local_path)
							result.append(fname)
						else:
							print('不是附件')
							pass
				#下载一个文件之后把这个文件移动到新的邮件文件夹，以便后面遍历for少一些数据。内容在下面。print ('</br>')
#				pcount += 1
		except Exception as ex:
			print("IMAP error %s " % ex)
		else:
			# close
			a.close()
			a.logout()
		
	elif emode == "pop" :
		#处理pop模式
		try:
			a = poplib()
            #待开发
		except Exception as e:
			print(e)
	return result

def main(argv):
	global local_path
	#初始化默认参数

	#workdate = ""
	help_msg="get_email.py -c /home/tomcat/.jenkins/jobs/业务邮件收取/workspace/get_email.cfg -d 20200615 -u aaaa -f true -a true"
	#共享盘设置
	smb = {
		'host':"192.168.100.1", #ip或域名 192.168.100.1
		'username':"aaa",
		'password':"aaaa",
		'domain':"domain.local",
	}

	try:
		opts,args= getopt.getopt( argv, "c:d:u:f:a:",["config=", "date=", "unit=", "force=","autodate="])
	except getopt.GetoptError:
		print help_msg
		sys.exit(2)
	try :
		for opt, arg in opts:
			if opt == '-h':
				print help_msg
				sys.exit(1)
			elif opt in ("-d", "--date"):
				if len(arg)<> 8 :
					raise Exception(u"日期")
				workdate = arg
			elif opt in ("-u", "--unit"):
				target = arg.decode("utf-8").replace("[","").replace("]","").split(",")
				if len(arg)== 0 or len(arg) == 0:
					raise Exception(u"功能号")
			elif opt in ("-c", "--config"):
				if not os.path.exists(arg) :
					raise Exception(u"配置文件")
				config_arg = arg
			elif opt in ("-f", "--force"):
				if arg.lower() not in ("true", "false") :
					raise Exception(u"强制下载")
				forced = arg == "true"
			elif opt in ("-a", "--autodate") :
				if arg.lower() not in ("true", "false"):
					raise Exception(u"自动日期")
				autodate = arg == 'true'
		if autodate :
			workdate=prev_weekday(date.today()).strftime("%Y%m%d")
		config = get_config(config_arg, workdate)
		if len(config) == 0 :
			raise Exception(u"配置文件内容\n")
	except Exception as e:
		print(u"\n!!! %s参数有误 !!!\n" % e)
		sys.exit(1)
	
	#本机临时目录不存在则创建，目录存在则清空
	local_path = "/dev/shm/get_email_" + workdate
	if not os.path.isdir(local_path):
		os.makedirs(local_path)
	filelist=os.listdir(local_path)                #列出该目录下的所有文件名
	if forced :
		for f in filelist:
			#清空本地临时目录内的文件
			filepath = os.path.join( local_path, f )   #将文件名映射成绝对路劲
			os.remove(filepath)
			filelist.remove(f)
	else:
#		print("已下载文件：%s" % filelist)
		pass
	
	if target[0].upper() == "ALL" :
		j = config.keys()
		#try:
		#	j.remove("测试使用")
		#except :
		#	print("没有[测试使用]")
	else :
		j = target
	for i in j :
		#逐个unit需求进行处理
		print("开始处理[%s]模块" % i)
		if i in config :
			target_filename = i + ".ok"
			if target_filename not in filelist:		#先判断是否已完成下载
				files = get_email(config[i])			#再读取邮件
				#上传到smb服务器
				try :
					#避免配置缺少项
					attach_max = int(config[i]['attach_counts'])
				except :
					attach_max = 0
				if len(files) >= attach_max :		#判断下载数量是否符合
					savefile( target_filename, "OK", local_path)		#生成下载完成标志
					for k in files:			#按已下载文件逐个处理
						for m in range(1, attach_max + 1):
							try:
								if config[i].get("attach_name_" + str(m)) == k:
									smb_copy(smb, config[i].get('attach_filename_' + str(m)), os.path.join(local_path,k), workdate)
									break
							except Exception as e:
								print("上传文件或配置内容有误%s" % e)
				else:
					print("[%s] 已下载文件数%s与配置不符%s" %(i, files, config[i]['attach_counts']))
			else:
				print("[%s] 附件已下载" % i)

if __name__ == '__main__' :
	main(sys.argv[1:])
