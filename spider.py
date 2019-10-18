import requests
import json
import time
import xlwt
import re
import math

# def xuanbao_get_reason():
# 	headers = {
# 		'Referer':'https://xuangubao.cn/dingpan/redian?tdsourcetag=s_pctim_aiomsg',
# 		'Origin':'https://xuangubao.cn',
# 		'Connection':'keep-alive',
# 		'Host':'flash-api.xuangubao.cn',
# 	}
# 	html = requests.get('https://flash-api.xuangubao.cn/api/surge_stock/plates',headers= headers)
# 	dic = json.loads(html.text)
# 	res = dict()
# 	#print(len(dic['data']['items']))
# 	for item in dic['data']['items']:
# 		name = item['name']
# 		description = ''
# 		if 'description' in item.keys():
# 			description = item['description']
# 		res[name] = description
# 	return res



# def xuan_gu_bao(ws):
# 	headers = {
# 		'Referer':'https://xuangubao.cn/dingpan/redian?tdsourcetag=s_pctim_aiomsg',
# 		'Origin':'https://xuangubao.cn',
# 		'Connection':'keep-alive',
# 		'Host':'flash-api.xuangubao.cn',
# 	}

# 	ws.write(0,0,'股票代码')
# 	ws.write(0,1,'股票名称')
# 	ws.write(0,2,'连扳数')
# 	ws.write(0,3,'价格')
# 	ws.write(0,4,'涨停时间')
# 	ws.write(0,5,'概念')
# 	ws.write(0,6,'涨停原因')
# 	ws.write(0,7,'板块原因')
# 	url = 'https://flash-api.xuangubao.cn/api/surge_stock/stocks?normal=true&uplimit=true'

# 	html = requests.get(url,headers=headers)
# 	row = 1

# 	dic = json.loads(html.text)
# 	for item in dic['data']['items']:
# 		reasons = xuanbao_get_reason()
# 		times = time.strftime('%Y/%m/%d %H:%M:%S',time.localtime(item[6]))
# 		ress = ''
# 		reason = ''
# 		if len(item[8]) != 0:
# 			for res in item[8]:
# 				ress += res['name'] + " "
# 				reason += reasons[res['name']] + " "
# 		ws.write(row,0,item[0])
# 		ws.write(row,1,item[1])
# 		ws.write(row,2,item[11])
# 		ws.write(row,3,str(item[2]))
# 		ws.write(row,4,times)
# 		ws.write(row,5,ress)
# 		ws.write(row,6,item[5])
# 		ws.write(row,7,reason)
# 		print('------选股宝---正在插入-----'+str(row))
# 		row  += 1
		
		# print("-----------------------")
		# print('股票代码:'+item[0])
		# print('股票名称:'+item[1])
		# print('连扳数:'+item[11])
		# print('价格:'+str(item[2]))
		# print('涨停时间:'+times)
		# print('概念:'+ress)
		# print('涨停原因:'+item[5])
		# print('板块原因:'+reason)
		
def get_tonghua_ids():
	ids = []
	session = requests.session()
	url = 'http://www.iwencai.com/stockpick/search?typed=1&preParams=&ts=1&f=1&qs=index_rewrite&selfsectsn=&querytype=stock&searchfilter=&tid=stockpick&w=%E6%B6%A8%E5%81%9C%E6%8F%AD%E7%A7%98++%E6%B6%A8%E5%81%9C'
	headers = {
		'Host': 'www.iwencai.com',
		'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36',
		'Connection': 'keep-alive',
		#'Referer': 'http://www.iwencai.com/stockpick/search?typed=1&preParams=&ts=1&f=1&qs=index_rewrite&selfsectsn=&querytype=stock&searchfilter=&tid=stockpick&w=%E6%B6%A8%E5%81%9C%E6%8F%AD%E7%A7%98++%E6%B6%A8%E5%81%9C',
	}
	with open('cookie.txt', 'r') as f:  
			cookie = f.read()
			cookieDict = {}
			cookies = cookie.split("; ")
			for co in cookies:
				co = co.strip()
				p = co.split('=')
				cookieDict[p[0]] = p[1] 
	session.cookies.update(cookieDict)

	while True:
		html = session.get(url,headers=headers).text
		
		page_par = '"total":(.*?),'
		
		nums = re.compile(page_par,re.S).findall(html)
		if len(nums) != 0:
			nums = int(nums[0])
			break
		print("请更换您的cookie,具体做法看备注！")
		time.sleep(10)
	pages = math.ceil(nums/30)
	token_par = '"token":"(.*?)"'

	token = re.compile(token_par,re.S).findall(html)

	#print(token[0])
	#print(pages)
	for i in range(1,int(pages)+1):
		time.sleep(5)
		url = 'http://www.iwencai.com/stockpick/cache?token='+token[0]+'&p='+str(i)+'&perpage=30&showType=[%22%22,%22%22,%22onTable%22,%22onTable%22,%22onTable%22,%22onTable%22,%22onTable%22,%22onTable%22,%22onTable%22,%22onTable%22,%22onTable%22,%22onTable%22,%22onTable%22,%22onTable%22,%22onTable%22,%22onTable%22,%22onTable%22]'
		html2 = session.get(url,headers=headers).text
		nob_par = '"wccode2hq":({.*?})'
		nob=re.compile(nob_par,re.S).findall(html2)
		if len(nob) ==1:
			nobs = eval(nob[0])
			for key in nobs.keys():
				ids.append(nobs[key][0])
			
		else:
			print('--------------发生错误！！！！！---------')

		#html1 = session.get(url,headers=headers).text

	return ids


def tong_hua_shun_par(html,ids):
	name_par = '<a.*?tid="this".*?posid="r1c2".*?title="(.*?)"'
	name = re.compile(name_par,re.S).findall(html)
	if len(name) == 1:
		name = name[0].split(" ")
		name = name[0]

	title_par = '涨停原因类别：(.*?)<'
	title = re.compile(title_par,re.S).findall(html)
	if len(title) == 1:
		title = title[0].strip()
	else:
		title = ' '

	reason_par = '<span class="open_btn">涨停原因.*?<div class="check_else">(.*?)<'
	reason = re.compile(reason_par,re.S).findall(html)
	if len(reason) == 1:
		reason = reason[0].strip()
	else:
		reason = " "

	nob1_par = '<a.*?title="此概念在该股票中贴合度排名第一".*?>(.*?)<'
	nob2_par = '<a.*?title="此概念在该股票中贴合度排名第二".*?>(.*?)<'
	nob3_par = '<a.*?title="此概念在该股票中贴合度排名第三".*?>(.*?)<'
	nob1 = re.compile(nob1_par,re.S).findall(html)
	nob2 = re.compile(nob2_par,re.S).findall(html)
	nob3 = re.compile(nob3_par,re.S).findall(html)
	if len(nob1) == 1:
		nob1 = "1、"+nob1[0].strip()+" "
	else:
		nob1 =" "

	if len(nob2) == 1:
		nob2 = "2、"+nob2[0].strip()+" "
	else:
		nob2 =" "

	if len(nob3) == 1:
		nob3 = "3、"+nob3[0].strip()+" "
	else:
		nob3 =" "
	ids = str(ids)
	if ids[0] == '6':
		ids += '.sh'
	else:
		ids += '.ss'

	return ids,name,title,reason,nob1+nob2+nob3

def tong_hua_shun_par1(html):
	person1_par = '<span.*?class="hltip">控制股东.*?class="tip">(.*?)</td>'
	person2_par = '<span.*?class="hltip">实际控制人.*?class="tip">(.*?)</td>'
	person3_par = '<span.*?class="hltip">最终控制人.*?class="tip">(.*?)</td>'

	person1 = re.compile(person1_par,re.S).findall(html)
	person2 = re.compile(person2_par,re.S).findall(html)
	person3 = re.compile(person3_par,re.S).findall(html)
	if len(person1) == 1:
		person1 = person1[0].strip().replace("</span>","")
	else:
		person1 =" "
	if len(person2) == 1:
		person2 = person2[0].strip().replace("</span>","")
	else:
		person2 =" "
	if len(person3) == 1:
		person3 = person3[0].strip().replace("</span>","")
	else:
		person3 =" "

	return person1,person2,person3

def tong_hua_shun(wa):
	wa.write(0,0,'股票代码')
	wa.write(0,1,'股票名称')
	wa.write(0,2,'概念')
	wa.write(0,3,'涨停原因')
	wa.write(0,4,'概念强弱排名')
	wa.write(0,5,'控制股东')
	wa.write(0,6,'实际控制人')
	wa.write(0,7,'最终控制人')
	row = 1
	headers = {
		'Host':'basic.10jqka.com.cn',
		'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36',
		'Connection': 'keep-alive',
		'Accept-Language':'zh-CN,zh;q=0.9,en;q=0.8',
	
	}
	ids = get_tonghua_ids()
	
	for i in ids:
		time.sleep(0.1)
		url = 'http://basic.10jqka.com.cn/'+str(i)+'/'
		url1 = 'http://basic.10jqka.com.cn/'+str(i)+'/holder.html'

		html = requests.get(url,headers=headers).text
		html=html.encode('ISO 8859-1').decode('gbk')
		num,name,title,reason,nob = tong_hua_shun_par(html,i)
		
		html1 = requests.get(url1,headers=headers).text
		html1=html1.encode('ISO 8859-1').decode('gbk')
		
		person1,person2,person3 = tong_hua_shun_par1(html1)
		print('---------同花顺--------正在插入'+str(row))
		wa.write(row,0,str(num))
		wa.write(row,1,name)
		wa.write(row,2,title)
		wa.write(row,3,reason)
		wa.write(row,4,nob)
		wa.write(row,5,person1)
		wa.write(row,6,person2)
		wa.write(row,7,person3)
		row += 1
		'''
		print('------------------------------')
		print('股票代码:'+str(num))
		print('股票名称:'+name)
		print('概念:'+title)
		print('涨停原因:'+reason)
		print('概念强弱排名:'+nob)
		
		print('控制股东:'+person1)
		print('实际控制人:'+person2)
		print('最终控制人:'+person3)
		'''

if __name__ == '__main__' :
	
	
	wb = xlwt.Workbook()
	#ws = wb.add_sheet('选股宝')
	wa = wb.add_sheet('同花顺')
	
	#xuan_gu_bao(ws)
	tong_hua_shun(wa)
	print("----------------------完成-----------------------------")
	wb.save('./股票.xls')	
