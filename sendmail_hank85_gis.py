#! usr/bin/python
#coding=utf-8 
from email import encoders
from email.header import Header
from email.mime.text import MIMEText
from email.utils import parseaddr, formataddr
from bs4 import BeautifulSoup
import smtplib
import xdrlib,sys
import xlrd
import time

reload(sys)
sys.setdefaultencoding('utf-8')

def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr(( \
        Header(name, 'utf-8').encode(), \
        addr.encode('utf-8') if isinstance(addr, unicode) else addr))

def _send_email(to_addr,title,context,types):
	from_addr = 'wintel_support@hank85.com'
	smtp_server = 'smtpinternal.hank85.com'

	msg = MIMEText(context, types, 'utf-8')
	msg['From'] = _format_addr(from_addr)
	msg['To'] = _format_addr(to_addr)
	msg['Subject'] = Header(title, 'utf-8').encode()

	server = smtplib.SMTP(smtp_server, 25)
	server.set_debuglevel(1)
	server.sendmail(from_addr, [to_addr] , msg.as_string())
	server.quit()

def _open(file):
	try:
		return xlrd.open_workbook(file)
	except:
		print '12345'

def _getList(file,sheetName):
	dataFile = _open(file)
	table=dataFile.sheet_by_name(sheetName)
	row_list = []
	for d in range(0,table.nrows):
		row_list.append(table.row_values(d))
	return row_list


def _makeTable(Hostname,IP,ITcode,OS_Version):
	_h=''
	_h+='<tr>'
	_h+='<td style="padding:.75pt .75pt .75pt .75pt"><p class="MsoNormal"><span lang="EN-US">'+Hostname+'<o:p></o:p></span></p></td>'
	_h+='<td style="padding:.75pt .75pt .75pt .75pt"><p class="MsoNormal"><span lang="EN-US">'+IP+'<o:p></o:p></span></p></td>'
	_h+='<td width="152" style="width:91.0pt;padding:.75pt .75pt .75pt .75pt"><p class="MsoNormal"><span lang="EN-US">'+ITcode+'</a><o:p></o:p></span></p></td>'
	_h+='<td width="11" valign="top" style="width:6.6pt;padding:0cm 0cm 0cm 0cm"><p class="MsoNormal"><span lang="EN-US">'+OS_Version+'<o:p></o:p></span></p></td>'
	_h+='<td style="padding:.75pt .75pt .75pt .75pt"></td>'
	_h+='<td style="padding:.75pt .75pt .75pt .75pt"></td>'
	_h+='</tr>'
	
	return _h


def _makeHtml(table):
	return '''<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40">
 <head>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <meta name="Generator" content="Microsoft Word 14 (filtered medium)" />
  <style><!--
/* Font Definitions */
@font-face
	{font-family:宋体;
	panose-1:2 1 6 0 3 1 1 1 1 1;}
@font-face
	{font-family:宋体;
	panose-1:2 1 6 0 3 1 1 1 1 1;}
@font-face
	{font-family:Calibri;
	panose-1:2 15 5 2 2 2 4 3 2 4;}
@font-face
	{font-family:"\@宋体";
	panose-1:2 1 6 0 3 1 1 1 1 1;}
/* Style Definitions */
p.MsoNormal, li.MsoNormal, div.MsoNormal
	{margin:0cm;
	margin-bottom:.0001pt;
	font-size:12.0pt;
	font-family:宋体;}
a:link, span.MsoHyperlink
	{mso-style-priority:99;
	color:blue;
	text-decoration:underline;}
a:visited, span.MsoHyperlinkFollowed
	{mso-style-priority:99;
	color:purple;
	text-decoration:underline;}
p.MsoListParagraph, li.MsoListParagraph, div.MsoListParagraph
	{mso-style-priority:34;
	margin:0cm;
	margin-bottom:.0001pt;
	text-indent:21.0pt;
	font-size:12.0pt;
	font-family:宋体;}
span.EmailStyle18
	{mso-style-type:personal;}
span.EmailStyle19
	{mso-style-type:personal;
	font-family:"Calibri","sans-serif";
	color:#1F497D;}
span.EmailStyle20
	{mso-style-type:personal-reply;
	font-family:"Calibri","sans-serif";
	color:#1F497D;}
.MsoChpDefault
	{mso-style-type:export-only;
	font-size:10.0pt;}
@page WordSection1
	{size:612.0pt 792.0pt;
	margin:72.0pt 90.0pt 72.0pt 90.0pt;}
div.WordSection1
	{page:WordSection1;}
/* List Definitions */
@list l0
	{mso-list-id:1330251677;
	mso-list-type:hybrid;
	mso-list-template-ids:2114093470 -1901806214 67698713 67698715 67698703 67698713 67698715 67698703 67698713 67698715;}
@list l0:level1
	{mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:48.0pt;
	text-indent:-18.0pt;}
@list l0:level2
	{mso-level-number-format:alpha-lower;
	mso-level-text:"%2\)";
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	text-indent:-21.0pt;}
@list l0:level3
	{mso-level-number-format:roman-lower;
	mso-level-tab-stop:none;
	mso-level-number-position:right;
	margin-left:93.0pt;
	text-indent:-21.0pt;}
@list l0:level4
	{mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:114.0pt;
	text-indent:-21.0pt;}
@list l0:level5
	{mso-level-number-format:alpha-lower;
	mso-level-text:"%5\)";
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:135.0pt;
	text-indent:-21.0pt;}
@list l0:level6
	{mso-level-number-format:roman-lower;
	mso-level-tab-stop:none;
	mso-level-number-position:right;
	margin-left:156.0pt;
	text-indent:-21.0pt;}
@list l0:level7
	{mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:177.0pt;
	text-indent:-21.0pt;}
@list l0:level8
	{mso-level-number-format:alpha-lower;
	mso-level-text:"%8\)";
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:198.0pt;
	text-indent:-21.0pt;}
@list l0:level9
	{mso-level-number-format:roman-lower;
	mso-level-tab-stop:none;
	mso-level-number-position:right;
	margin-left:219.0pt;
	text-indent:-21.0pt;}
ol
	{margin-bottom:0cm;}
ul
	{margin-bottom:0cm;}
--></style>
  <!--[if gte mso 9]><xml>
<o:shapedefaults v:ext="edit" spidmax="1026" />
</xml><![endif]-->
  <!--[if gte mso 9]><xml>
<o:shapelayout v:ext="edit">
<o:idmap v:ext="edit" data="1" />
</o:shapelayout></xml><![endif]-->
 </head>
 <body lang="ZH-CN" link="blue" vlink="purple">
  <div class="WordSection1">
   <p class="MsoNormal"><span lang="EN-US">
     <o:p>
      &nbsp;
     </o:p></span></p>
   <p class="MsoNormal">亲爱的应用管理员，您好！ <span lang="EN-US">
     <o:p></o:p></span></p>
   <p class="MsoNormal" style="text-indent:24.0pt">根据安全管理规定，<span lang="EN-US">-ITS-IAAS-Computing</span>团队将开始<span lang="EN-US">Q3</span>季度<span lang="EN-US">Windows&amp;Linux</span>服务器打补丁及重启工作，初步定于<span lang="EN-US">2016</span>年<span lang="EN-US">10</span>月<span lang="EN-US">15</span>日及<span lang="EN-US">10</span>月<span lang="EN-US">16</span>日那个周末进行<span style="color:#1F497D">，</span>还请您反馈是否同意<span lang="EN-US">-ITS-IAAS-Computing</span>团队安排的时间？如有任何异议，还请您反馈合适的打补丁和重启时间。一共<span lang="EN-US">4</span>个时间选项，请在补丁窗口列填写。<span lang="EN-US" style="color:#1F497D">
     <o:p></o:p></span></p>
   <p class="MsoNormal"><span lang="EN-US">1 </span>正常工作时间（周一至周五）可操作。 <span lang="EN-US">
     <o:p></o:p></span></p>
   <p class="MsoNormal"><span lang="EN-US">2 </span>周末可操作，默认为<span lang="EN-US">2016</span>年<span lang="EN-US">10</span>月<span lang="EN-US">15</span>日及<span lang="EN-US">10</span>月<span lang="EN-US">16</span>日。 <span lang="EN-US">
     <o:p></o:p></span></p>
   <p class="MsoNormal"><span lang="EN-US">3 </span>战略平台<span lang="EN-US">outage</span>期间可操作。（仅限重要系统） <span lang="EN-US">
     <o:p></o:p></span></p>
   <p class="MsoNormal"><span lang="EN-US">4 </span>本季度不打补丁<span lang="EN-US"> ,</span>不打补丁需要提交<span lang="EN-US">Exception</span>申请并得到您部门负责人的审批。 <span lang="EN-US">
     <o:p></o:p></span></p>
   <p class="MsoNormal"><span lang="EN-US">
     <o:p>
      &nbsp;
     </o:p></span></p>
   <p class="MsoNormal"><span lang="EN-US">&nbsp;&nbsp;&nbsp;&nbsp;</span>注意事项列请填写需要的注意事项，比如需要补丁前电话确认已停止应用，群集启动后要在特殊节点上，应用管理员已经变更等 <span lang="EN-US">
     <o:p></o:p></span></p>
   <p class="MsoNormal"><span lang="EN-US">&nbsp;&nbsp;&nbsp;&nbsp;</span>如果在<span lang="EN-US">2016</span>年<span lang="EN-US">10</span>月<span lang="EN-US">12</span>日 前未得到您的反馈，我们将视为您同意<span lang="EN-US">-ITS-IAAS-Computing</span>安排计划时间。 <span lang="EN-US">
     <o:p></o:p></span></p>
   <p class="MsoNormal"><span lang="EN-US">&nbsp;&nbsp;&nbsp;&nbsp;</span>补丁成功升级后，我们会以电话和邮件的方式通知您，请您尽快测试应用并在<span lang="EN-US">24</span>小时内回复测试结果邮件； <span lang="EN-US">
     <o:p></o:p></span></p>
   <p class="MsoNormal"><span lang="EN-US">&nbsp;&nbsp; 
     <o:p></o:p></span></p>
   <p class="MsoNormal"><span lang="EN-US">English
     <o:p></o:p></span></p>
   <p class="MsoNormal"><span lang="EN-US">Dear Application Administrators,
     <o:p></o:p></span></p>
   <p class="MsoNormal"><span lang="EN-US">&nbsp;&nbsp;IT -ITS-IAAS-Computing team will soon begin OS patch installing and OS restarting task on Servers with Windows&amp;Linux OS for Q3 according to security compliance policy. We have planned to do this on the weekend(10/15/2016-10/16/2016).Therefore, if you are disagree on our plan or have any your own opinion on this, please fill your feedback on the highlight field below table(time referred below is Beijing Time, GMT +8):
     <o:p></o:p></span></p>
   <p class="MsoNormal"><span lang="EN-US">1.Normal work hours is available for above OS patch installing and OS restarting actions (Monday to Friday)
     <o:p></o:p></span></p>
   <p class="MsoNormal"><span lang="EN-US">2.Weekend is available for above OS patch installing and OS restarting actions;(default weekend: 2016.10.15 and 2016.10.16)
     <o:p></o:p></span></p>
   <p class="MsoNormal"><span lang="EN-US">3.Outage window for strategic platform(only limits to critical systems)
     <o:p></o:p></span></p>
   <p class="MsoNormal"><span lang="EN-US">4.Skip OS patch installing and OS restarting actions for this Q, but this requires your Exception Approval from your department lead(submit Exception application following security policy)
     <o:p></o:p></span></p>
   <p class="MsoNormal"><span lang="EN-US">
     <o:p>
      &nbsp;
     </o:p></span></p>
   <p class="MsoNormal"><span lang="EN-US">&nbsp;&nbsp; Please note down other related notes for this action that OS team needs to be careful for in the field </span>“<span lang="EN-US">Remarks</span>”<span lang="EN-US">. For example:
     <o:p></o:p></span></p>
   <p class="MsoListParagraph" style="margin-left:48.0pt;text-indent:-18.0pt;mso-list:l0 level1 lfo2">
    <!--[if !supportLists]--><span lang="EN-US"><span style="mso-list:Ignore">1.<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp; </span></span></span>
    <!--[endif]--><span lang="EN-US">OS admin need to confirm through phone that the application is fully stopped before the action with Application admin;
     <o:p></o:p></span></p>
   <p class="MsoListParagraph" style="margin-left:48.0pt;text-indent:-18.0pt;mso-list:l0 level1 lfo2">
    <!--[if !supportLists]--><span lang="EN-US"><span style="mso-list:Ignore">2.<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp; </span></span></span>
    <!--[endif]--><span lang="EN-US">Cluster node needs to be set on a specific node;
     <o:p></o:p></span></p>
   <p class="MsoListParagraph" style="margin-left:48.0pt;text-indent:-18.0pt;mso-list:l0 level1 lfo2">
    <!--[if !supportLists]--><span lang="EN-US"><span style="mso-list:Ignore">3.<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp; </span></span></span>
    <!--[endif]--><span lang="EN-US">Application admin is now changed and current contact is not available;
     <o:p></o:p></span></p>
   <p class="MsoNormal" style="margin-left:30.0pt"><span lang="EN-US" style="color:#1F497D">
     <o:p>
      &nbsp;
     </o:p></span></p>
   <p class="MsoNormal" style="margin-left:30.0pt"><span lang="EN-US">NOTES:
     <o:p></o:p></span></p>
   <p class="MsoListParagraph" style="margin-left:48.0pt;text-indent:0cm"><span lang="EN-US">&nbsp;&nbsp; if there is no email feedback from you before 2016.10.12, the OS patch installing and OS restarting actions for your application server will be managed and set as -ITS-IAAS-Computing</span>’<span lang="EN-US">s plan. We will email/call you after the OS patch installing and OS restarting are successfully completed and please test your application in 24hours and let us know your test results.
     <o:p></o:p></span></p>
   <p class="MsoListParagraph" style="margin-left:48.0pt;text-indent:0cm"><span lang="EN-US">
     <o:p>
      &nbsp;
     </o:p></span></p>
   <table class="MsoNormalTable" border="1" cellspacing="3" cellpadding="0" width="1539" style="width:923.55pt">
    <tbody>
     <tr>
      <td style="padding:.75pt .75pt .75pt .75pt"><p class="MsoNormal"><span lang="EN-US">Hostname
         <o:p></o:p></span></p></td>
      <td style="padding:.75pt .75pt .75pt .75pt"><p class="MsoNormal"><span lang="EN-US">IP
         <o:p></o:p></span></p></td>
      <td width="152" style="width:91.0pt;padding:.75pt .75pt .75pt .75pt"><p class="MsoNormal">应用管理员邮箱<span lang="EN-US"><br />app admin email
         <o:p></o:p></span></p></td>
      <td width="11" valign="top" style="width:6.6pt;padding:0cm 0cm 0cm 0cm"><p class="MsoNormal"><span lang="EN-US">OS Version
         <o:p></o:p></span></p></td>
      <td style="padding:.75pt .75pt .75pt .75pt"><p class="MsoNormal">补丁窗口<span lang="EN-US">(</span>请选<span lang="EN-US">1</span>，<span lang="EN-US">2</span>，<span lang="EN-US">3</span>，<span lang="EN-US">4</span>填写<span lang="EN-US">)<br />patch windows,Please select from 1 to 4 on above description
         <o:p></o:p></span></p></td>
      <td style="padding:.75pt .75pt .75pt .75pt"><p class="MsoNormal">注意事项<span lang="EN-US">(</span>请填写需要的注意事项<span lang="EN-US">)<br />Remarks(Please note down other related notes for this actions that OS team need to be careful for)
         <o:p></o:p></span></p></td>
     </tr>
     '''+table+'''
    </tbody>
   </table>
   <p class="MsoNormal"><span lang="EN-US">
     <o:p>
      &nbsp;
     </o:p></span></p>
  </div>
 </body>
</html>'''

def main(sheetName,file1,file2):
	dnsList=_getList(file1,sheetName)
	itcodeList=_getList(file2,sheetName)
	for l in itcodeList:
		_h=''
		itcode=l[0]
		for d in dnsList:
			if itcode==d[2]:
				_h+=_makeTable(d[0],d[1],d[2],d[3])
		htmlText=_makeHtml(_h)		
		_send_email(str(itcode)+'@hank85.com','请确认Windows&Linux系统打补丁及重启的时间 | Need your help to confirm windows&linux server 2016Q3 security patch window',htmlText,'html')

if __name__ == '__main__':
	main('Sheet1','Test.xlsx','itcode.xlsx')
	
