'''
预备添加电话预约与巡检单的对比.同时电话预约数据可以迁移到巡检数据中。
'''
import tkinter as tk
import tkinter.filedialog#弹窗
from PIL import Image
from PIL import ImageTk
from tkinter import StringVar
import pymysql,tkinter.ttk#复选菜单模块,threading
import openpyxl,datetime,re
import win32api
import win32print



def mysql_jk_select(sql):#MySQL查询接口
	conn = pymysql.connect(host='localhost', port=3306, user='root', passwd='cq123456', db='业务', charset='utf8')
	cour=conn.cursor()
	effect_cour=cour.execute(sql)
	p=cour.fetchall()
	return p
def mysql_jk_insert(sql):#插入数据接口
	conn=pymysql.connect(user="root",host="localhost",passwd="cq123456",db="业务",port=3306,charset="utf8")
	cur=conn.cursor()
	cur.execute(sql)
	conn.commit()
	conn.close()
	return "ok"


def ibsert_print():#打印接口弹窗/导入已查询名单内
	def improt_data_min():
		s=lb.curselection()
		print(s)
		for one_database in s:
			t=num[one_database]
			file_id=t[0]
			file_name=t[1]
			improt_data_url=tk.filedialog.askopenfile(title=f"请选择文件:{file_name}")
			improt_data_url=str(improt_data_url)
			file_url=re.search(r"<_io.TextIOWrapper name='C:/Users/jade/Desktop/(.*?).jpg' mode='r' encoding='cp936'>",improt_data_url).group(1)
			file_date=datetime.date.today()
			file_date=str(file_date)
			sql=f"insert into world values({file_id},'{file_name}','{file_url}','{file_date}');"
			mysql_jk_insert(sql)
	
	def	selcet_moth_min():
		selecet_moth=selcet_moth_entry.get()
		sql1=f"SELECT*FROM `电话联系表` where  month(data)='{selecet_moth}' and Merchant_id NOT IN (SELECT Merchant_id FROM world)"
		global num
		num=mysql_jk_select(sql1)
		lb.delete(0,500)
		for i in num:
			
			lb.insert(500,i)
	def ibsert_print_qqq():#打印接口
		
		s=lb.curselection()
		print(s)

		for one_database in s:
			
			t=num[one_database]
			wb2=openpyxl.load_workbook(r"江苏农行工单竖版正式_修改版.xlsx")
			sheets = wb2.sheetnames
			wb=wb2[sheets[0]]                 
			wb["M4"]=f"商户编号:{t[0]}"
			wb["M5"]="商户名称:"+t[1]                  
			wb["M6"]=f"商户联系电话:{t[3]}"
			wb2.save(f"C:\\Users\\jade\\Desktop\\'{t[1]}_{t[0]}'巡检单.xlsx")
			wb2.close()
	#正文
	sql=tk.Toplevel()
	sql.title("选择打印模块")
	sql.geometry("800x700")

	
	insert_time_input.delete(0,"end")
	
	
	lable_rcoll=tk.Scrollbar(sql)
	lable_rcoll.place(x=750,height=800)
	lb=tk.Listbox(sql,yscrollcommand=lable_rcoll.set,selectmode='multiple')
	lb.place(x=20,y=30,height=650,width=730)
	
	lable_rcoll.config(command=lb.yview)
	ibsert_print_butt=tk.Button(sql,text="打印单据",command=ibsert_print_qqq).place(x=340,y=0)
	selcet_moth=tk.Button(sql,text="按月查询",command=selcet_moth_min).place(x=260,y=0)
	selcet_moth_lanbel=tk.Label(sql,text="月份:").place(x=50,y=5)
	selcet_moth_entry=tk.Entry(sql)
	selcet_moth_entry.place(x=90,y=5)
	improt_data=tk.Button(sql,text="巡检确认按钮",command=improt_data_min).place(x=400,y=0)

	
	sql.mainloop()

	
def ibsert_information():#电话预约录入按钮接口
	A=insert_id.get()
	B=insert_name.get()
	C=insert_pos_information.get()
	D=insert_phone.get()
	e=insert_time.get()
	f=insert_mac_min.get()
	if C=="pos退回":
		sql=f'INSERT into `电话联系表`values({int(A)},"{B}",1,{int(D)},"{e}","{f}");'
		my_sql=mysql_jk_insert(sql)
		
	if C=="pos丢失":
		sql=f'INSERT into `电话联系表`values({int(A)},"{B}",2,{int(D)},"{e}","{f}");'
		my_sql=mysql_jk_insert(sql)
		
	if C=="pos正常":
		sql=f'INSERT into `电话联系表`values({int(A)},"{B}",0,{int(D)},"{e}","{f}");'
		my_sql=mysql_jk_insert(sql)
	if C=="人不在本地,无法巡检":
		sql=f'INSERT into `电话联系表`values({int(A)},"{B}",3,{int(D)},"{e}","{f}");'
		my_sql=mysql_jk_insert(sql)		
	insert_id_input.delete(0,"end")
	insert_name_input.delete(0,"end")
	insert_is_pos_Information.current(0)
	insert_phone_input.delete(0,"end")
	insert_time_input.delete(0,"end")
	insert_mac.delete(0,"end")		
		
def reply():##查询（id)弹窗窗口模块
	
	s=e.get()
	print(type(s))

	
	conn = pymysql.connect(host='localhost', port=3306, user='root', 
            passwd='cq123456', db='业务', charset='utf8')
	cour=conn.cursor()
	effect_cour=cour.execute(f"SELECT * FROM world WHERE Merchant_id LIKE '%{s}%'")
	p=cour.fetchone()
	print(p)
	p0=p[0]
	p1=p[1]
	p2=p[2]
	cour.close()
	sql1=tk.Toplevel()
	sql1.title("查询软件")
	sql1.geometry("600x500")
	id_label_min=tk.Label(sql1,text=f"id:").place(x=0,y=10)
	id_man=StringVar()
	id_man.set(p0)
	id_input_min=tk.Entry(sql1,textvariable=id_man).place(x=20,y=10)
	name_label_min=tk.Label(sql1,text="商户名称").place(x=120,y=10)
	name_man=StringVar()
	name_man.set(p1)
	name_input_min=tk.Entry(sql1,textvariable=name_man).place(x=180,y=10)
	img=Image.open(r"E:\\img\\{}.jpg".format(p2))
	img=img.resize((500,400),Image.ANTIALIAS)
	img = ImageTk.PhotoImage(img)
	sss=tk.Label(sql1,image=img).place(x=0,y=60)
	sql1.mainloop()

sql=tk.Tk()
sql.title("查询软件")
sql.geometry("400x500")
selcet_name=tk.Label(sql,text="数据查询功能(查询商户名字不好用)").place(x=0,y=0)
id_label=tk.Label(sql,text="id:").place(x=20,y=30)
e=StringVar()
id_input=tk.Entry(sql,textvariable=e).place(x=60,y=30)
name_label=tk.Label(sql,text="商户名称:").place(x=0,y=60)
name_input=tk.Entry().place(x=60,y=60)
seclet_button=tk.Button(sql,text="查询",command=reply).place(x=300,y=30)
#电话预约录入/及查询打印功能：
insert_name_label_BT=tk.Label(sql,text="预约录入功能").place(x=0,y=150)#标题
insert_burron=tk.Button(sql,text="电话联系预约",width=13,command=ibsert_information).place(x=270,y=180)
insert_print=tk.Button(sql,text="打印接口",width=13,command=ibsert_print).place(x=270,y=230)
insert_id_label=tk.Label(sql,text="id:").place(x=20,y=180)
insert_id=StringVar()
insert_id_input=tk.Entry(sql,textvariable=insert_id)
insert_id_input.place(x=50,y=180)#修饰后可能会转义
insert_name_label=tk.Label(sql,text="商户名称:").place(x=0,y=210)

insert_name=StringVar()
insert_name_input=tk.Entry(sql,textvariable=insert_name)
insert_name_input.place(x=60,y=210)
insert_pos_information=StringVar()
insert_is_pos_Information_label=tk.ttk.Label(sql,text="特殊情况说明:").place(x=0,y=240)
insert_is_pos_Information=tk.ttk.Combobox(sql,values=["pos正常","pos退回","pos丢失",'人不在本地,无法巡检'],textvariable=insert_pos_information)
insert_is_pos_Information.place(x=80,y=240)
insert_is_pos_Information.current(0)
insert_phone_lable=tk.Label(sql,text="电话:")
insert_phone_lable.place(x=5,y=270)
insert_phone=StringVar()
insert_phone_input=tk.Entry(sql,textvariable=insert_phone)
insert_phone_input.place(x=45,y=270)
insert_time_lable=tk.Label(sql,text="时间:")
insert_time_lable.place(x=210,y=270)
insert_time=StringVar()
insert_time_input=tk.Entry(sql,textvariable=insert_time)
insert_time_input.place(x=240,y=270)
inseer_mac_lable=tk.Label(sql,text="地址：").place(x=5,y=310)
insert_mac_min=StringVar()
insert_mac=tk.Entry(sql,textvariable=insert_mac_min)
insert_mac.place(x=45,y=310)


sql.mainloop()
