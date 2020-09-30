import requests,lxml.etree,bs4,re,time
for i in range(1,6):
	url=f"https://search.51job.com/list/071800,000000,0000,00,9,99,%25E4%25BC%259A%25E8%25AE%25A1,2,{i}.html?lang=c&postchannel=0000&workyear=99&cotype=99&degreefrom=99&jobterm=99&companysize=99&ord_field=0&dibiaoid=0&line=&welfare="
	headers={"User-Agent":"Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.106 Safari/537.36"}
	params={"lang":"c",
			"postchannel":"0000",
			"workyear":"99",
			"cotype":"99",
			"degreefrom":"99",
			"jobterm":"99",
			"companysize":"99",
			"ord_field":"0",
			"dibiaoid":"0",
			"line":"",
			"welfare":""	}
	spones=requests.get(url=url,headers=headers,params=params)
	spones.encoding="gbk"
	#print(spones.text)
	li=re.findall(r'"job_href":"(.*?)"',spones.text)
					
	
	for job_url in li:
		new_job=s=re.sub(r"\\","",job_url)
		spones=requests.get(url=new_job,headers=headers)
		spones.encoding="gbk"
		lis1=lxml.etree.HTML(spones.text)
		lis2=lis1.xpath("//p[@class='cname']/a[@class='catn']/text()")#招聘单位
		lis3=lis1.xpath("//div[@class='in']/div[@class='cn']/h1/text()")#招聘职位
		lis4=lis1.xpath("//div[@class='in']/div[@class='cn']/strong/text()")#薪资
		lis5=lis1.xpath("//div[@class='bmsg inbox']/p[@class='fp']/text()")#地址
		print('"职位："{},"招聘单位："{},"薪资"："{},"地址："{},"页面信息：{}'.format(lis3,lis2,lis4,lis5,new_job))
