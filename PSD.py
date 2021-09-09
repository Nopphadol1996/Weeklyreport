from tkinter import *
from tkinter import ttk,messagebox
import csv
import pandas as pd
import matplotlib.pyplot as plt
import sqlite3
from datetime import datetime

######### DB #####################
conn = sqlite3.connect('Weekly.db')
# สร้างตัวดำเนินการ (อยากได้อะไรใช้ตัวนี้ได้เลย)
c = conn.cursor()

c.execute("""CREATE TABLE IF NOT EXISTS weeklytable (
				ID INTEGER PRIMARY KEY AUTOINCREMENT,
				transactionid TEXT,
				Date_ TEXT,
				Station TEXT,
				Bound TEXT,
				Door TEXT,
				Time_ TEXT,
				Failre TEXT,
				Cause Text,
				Resolution Text,
				Work INTEGER,
				QTY INTEGER
			)""")

c.execute("""CREATE TABLE IF NOT EXISTS plotstation (
				ID INTEGER PRIMARY KEY AUTOINCREMENT,
				transectionid TEXT,
				station TEXT,
				qty INTEGER,
				week TEXT
			)""")


def insert_week_station(transactionid,station,qty,week): # เอาที่เราสร้างมาใส่
	ID = None
	with conn:
		c.execute("""INSERT INTO plotstation VALUES (?,?,?,?,?)""", # ? ต้องรวม ID = None
			(ID,transactionid,station,qty,week)) #ใส่ ID ไปด้วย
		conn.commit() # คือ การบันทึกข้อมูลลงในฐานข้อมูล ถ้าไม่รันตัวนี้จะไม่บันทึก
		#print('Insert Sucess...!')

def show_station_week():
	with conn:
		c.execute("SELECT *FROM plotstation")
		veryweek = c.fetchall() # คำสั่งให้ดึงข้อมูลมา
		#print(veryweek)
	return veryweek

def plot_station():


	df =pd.read_sql_query("SELECT * FROM plotstation",conn,)
	del df['ID']

	#print(df)
	try:
		
		a = df.pivot_table(index='station',columns='week',values='qty')
		a.plot(kind='bar',stacked=True,figsize=(12,6))

		plt.ylim(0,41)
		plt.grid(axis = 'y')
		plt.title('Station Report',color='green')
		plt.legend(bbox_to_anchor=(1.01, 1), loc=2, borderaxespad=0.,fontsize=7)
		plt.show(block=False) # ให้ผุ็ใช้เปิดได้หลายจอ

	except:
		messagebox.showerror('Error','ไม่มีข้อมูลที่จะแสดง')



def insert_work(transactionid,Date_,Station,Bound,Door,Time_,Failre,Cause,Resolution,Work,QTY): # เอาที่เราสร้างมาใส่
	ID = None
	with conn:
		c.execute("""INSERT INTO weeklytable VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""", # ? ต้องรวม ID = None
			(ID,transactionid,Date_,Station,Bound,Door,Time_,Failre,Cause,Resolution,Work,QTY)) #ใส่ ID ไปด้วย
		conn.commit() # คือ การบันทึกข้อมูลลงในฐานข้อมูล ถ้าไม่รันตัวนี้จะไม่บันทึก
		#print('Insert Sucess...!')

#insert_work('31/01/2021','CEN','EB','D07','10:20:31','AMC_S: Obstacle Detection','Door closed too slow','Reset DCU',600600123,1)

def show_expense():
	with conn:
		c.execute("SELECT *FROM weeklytable")
		expense = c.fetchall() # คำสั่งให้ดึงข้อมูลมา
		# print(expense)
	return expense

def update_expense(transactionid,Date_,Station,Bound,Door,Time_,Failre,Cause,Resolution,Work,QTY):
	with conn:
		########################## ต้องเหมิอนกับในdatabase ###############
		c.execute("""UPDATE weeklytable SET Date_=?, Station=?, Bound=?, Door=? ,Time_=?,Failre=?,Cause=?,Resolution=?,Work=?,QTY=? WHERE transactionid=?""",
			([Date_,Station,Bound,Door,Time_,Failre,Cause,Resolution,Work,QTY,transactionid]))### Were ID transactionidต้องมาอยู้หลัง
		conn.commit()
		#print('Data update')
#update_expense('202109081423752234','31/01/2021','E1','EB','D07','10:20:31','AMC_S: Obstacle Detection','Door closed too slow','Reset DCU',600600123,5)

def delete_fuilt(transactionid):
	with conn:
		c.execute("DELETE FROM weeklytable WHERE transactionid=?",([transactionid])) #ใส่เป็น list
	conn.commit()

	# print('------Data Deleted----')

def pivot_table_1():
	try:

		df =pd.read_sql_query("SELECT * FROM weeklytable",conn,)
		del df['Work']
		del df['ID']

		df.pivot_table(index=['Station','Bound','Door'], columns ='Cause',values='QTY',
	              margins=False, margins_name='Grand Total').plot(kind='bar',fontsize=15,stacked=True)


		plt.title('Weekly Report',color='green')
		plt.ylabel('Total Failure',color='green')

		plt.ylim(0,10)
		plt.grid(axis = 'y')
		plt.show(block=False)

	except Exception as e:
		messagebox.showerror('ERROR','ไม่มีข้อมูลในตาราง')
		#Py_Initialize()

######### DB #####################
'''
def Export_station():

	df =pd.read_sql_query("SELECT * FROM plotstation",conn,)
	del df['ID']

	try:

		filepath= 'C:/Nopphadol/Desktop/HHPSd/pythontoexe/My_chart_failure.xlsx'
		writer = pd.ExcelWriter(filepath)
		df.to_excel(writer,'Sheet1',index=False)
		writer.save()
		#print('------------')
		messagebox.showinfo('Sucess','บักทึกข้อมูลสำเร็จ')
	except:
		messagebox.showerror('Error','กรุณาปิด Excel ก่อนบันทึกข้อมูล')

def Export_failure():

	try:

		df =pd.read_sql_query("SELECT * FROM weeklytable",conn,)
		del df['Work']
		del df['ID']

		filepath = 'C:/Users/Nopphadol/Desktop/Project_beginer/export/my_failuretable.xlsx'
		writer = pd.ExcelWriter(filepath)
		df.to_excel(writer,'sheet1',index=False)
		writer.save()
		#print('------------')
		messagebox.showinfo('Sucess','บักทึกข้อมูลสำเร็จ')
	except:
		messagebox.showerror('Error','กรุณาปิด Excel ก่อนบันทึกข้อมูล')

'''
root = Tk()


w = 1300 # กว้าง
h = 670 # สูง

ws = root.winfo_screenwidth() #screen width เช็คความกว้างของหน้า
hs = root.winfo_screenheight() #screen height

x = (ws/2) - (w/2) # ws คือความกว้างของหน้าจอทั้งหมด /2 คือครึ่งหนึ่งคือ CENTER
y = (hs/2) - (h/2) - 45
root.geometry(f'{w}x{h}+{x:.0f}+{y:.0f}')

#root.resizable(width=False,height=False) #### ปิดขยายหน้าจอ
root.title('Weekly Report V.1.0')
root.iconbitmap(r'icon_title.ico')

def Exit():

	root.destroy()

def About():
	messagebox.showinfo('About','นี่คือโปรแกรม Weekly Report ของแผนก PSD\n	')



def Save():

	my_workorder  = E1_work.get()
	my_time = E2_time.get()
	my_days = dayschoosen.get()
	my_months = monthschoosen.get()
	my_years = yearschoosen.get()
	my_station = Stationchoosen.get()
	my_bound = Boundchoosen.get()
	my_door = Doorchoosen.get()
	my_failure = Failurechoosen.get()
	my_qty = qtychoosen.get()
	my_cause = Causeechoosen.get()
	my_Maintenance = Maintenancechoosen.get()

	textdate  = my_days+'/'+my_months+'/'+my_years
	stamp = datetime.now()
	dt = stamp.strftime('%Y-%m-%d %H:%M:%S')
	transactionid = stamp.strftime('%Y%m%d%H%M%f') # สร้าง transection ID
	#print(type(transactionid))

	#text = '{} {} {} {} {} {} {} {} {} {} '.format(transactionid,textdate,my_station,my_bound,my_door,my_time,my_failure,my_cause,my_Maintenance,my_workorder,my_qty)
	#print(text)

	
	dayschoosen.set('day')
	monthschoosen.set('months')
	yearschoosen.set('years')
	Stationchoosen.set('Station')
	Boundchoosen.set('Bound')
	Doorchoosen.set('Door')
	E2_time.set('')
	Failurechoosen.set('Failure log')
	Causeechoosen.set('Cause')
	Maintenancechoosen.set('Action')
	E1_work.set('')
	qtychoosen.set('QTY')
	try:


		insert_work(transactionid,textdate,my_station,my_bound,my_door,my_time,my_failure,my_cause,my_Maintenance,int(my_workorder),int(my_qty))
		'''
		########################## ถ้าต้องการsave csv ให้เปิดอันนี้ ##################
		with open('test.csv','a',encoding='utf-8',newline='') as f:
		
				# with คือ คำสั่งเปิดไฟล์แล้วปิดอัตโนมัติ
				# 'a' คือ การบันทึกไปเรื่อยๆ เพิ่มข้อมูลจากข้อมูลเก่า แต่ถ้า w  เคลียค่าเก่าแล้วบันทึกใหม่
				# newline='' คือการทำให้ข้อมูลไม่มีบรรทัดว่าง
				fw = csv.writer(f) # สร้างฟังก์ชั่นสำหรับเขียนข้อมูล
				data = [textdate,my_station,my_bound,my_door,my_time,my_failure,
										my_cause,my_Maintenance,my_workorder,my_qty] # เอา Transection ID มาใส่ ใน treeview
				fw.writerow(data)
		'''			
		update_table()
		messagebox.showinfo('Successfuly','บันทึกข้อมูลสำเร็จ')
		E1.focus()

	except Exception as e:
		#print('----',e)
		#print('โปรดตรวจสอบ:\n Work order ต้องเป็นตัวเลข หรือ\n รูปบแบบวันเวลาต้อง 00:00:00 หรือ\n เลือกจำนวน QTY')
		messagebox.showerror('ERROR','โปรดตรวจสอบ:\n Work order ต้องเป็นตัวเลข หรือ\n รูปบแบบวันเวลาต้อง 00:00:00 หรือ\n เลือกจำนวน QTY')
############### สร้าง TAB ###################
#root.bind('<Return>',Save) # ต้องเพิ่มใน def Save(event=None)

def update_table():

	resulttable.delete(*resulttable.get_children())
	data_db = show_expense()
	#insert_work(textdate,my_station,my_bound,my_door,my_time,my_failure,my_cause,my_Maintenance,int(my_workorder),int(my_qty))
	for d in data_db:
		alltransection[d[1]] = d[1:]
		resulttable.insert('','end',value=d[1:])

def update_table_T4():

	resulttableT4.delete(*resulttableT4.get_children())
	data_dbt4 = show_station_week()
	for d in data_dbt4:

		resulttableT4.insert('','end',value=d[1:])




def Save_station():

	N2 = Station_N2.get()	
	N3 = Station_N3.get()
	E1 = Station_E1.get()
	E4 = Station_E4.get()
	E5 = Station_E5.get()
	E6 = Station_E6.get()
	E9 = Station_E9.get()
	S2 = Station_s2.get()
	S3 = Station_s3.get()
	S5 = Station_s5.get()
	CEN = Station_CEN.get()
	try:

		QN2 = int(QTY_N2.get())
		QN3 = int(QTY_N3.get())
		QE1 = int(QTY_E1.get())
		QE4 = int(QTY_E4.get())
		QE5 = int(QTY_E5.get())
		QE6 = int(QTY_E6.get())
		QE9 = int(QTY_E9.get())
		QS2 = int(QTY_S2.get())
		QS3 = int(QTY_S3.get())
		QS5 = int(QTY_S5.get())
		QCEN = int(QTY_CEN.get())

		Week = Weekstation.get()

		stamp1 = datetime.now()
		dt = stamp1.strftime('%Y-%m-%d %H:%M:%S')
		transactionid = stamp1.strftime('%Y%m%d%H%M%f') # สร้าง transection ID

		#print(N2,N3,E1,E4,E5,E6,E9,S2,S3,S5,CEN,QTY,Week)

		insert_week_station(transactionid,N2,QN2,Week)
		insert_week_station(transactionid,N3,QN3,Week)
		insert_week_station(transactionid,E1,QE1,Week)
		insert_week_station(transactionid,E4,QE4,Week)
		insert_week_station(transactionid,E5,QE5,Week)
		insert_week_station(transactionid,E6,QE6,Week)
		insert_week_station(transactionid,E9,QE9,Week)
		insert_week_station(transactionid,S2,QS2,Week)
		insert_week_station(transactionid,S3,QS3,Week)
		insert_week_station(transactionid,S5,QS5,Week)
		insert_week_station(transactionid,CEN,QCEN,Week)
		
		QTY_N2.set('QTY_N2')
		QTY_N3.set('QTY_N3')
		QTY_E1.set('QTY_E1')
		QTY_E4.set('QTY_E4')
		QTY_E5.set('QTY_E5')
		QTY_E6.set('QTY_E6')
		QTY_E9.set('QTY_E9')
		QTY_S2.set('QTY_S2')
		QTY_S3.set('QTY_S3')
		QTY_S5.set('QTY_S5')
		QTY_CEN.set('QTY_CEN')
		Weekstation.set('Week')
		update_table_T4()
		messagebox.showinfo('Successfuly','บันทึกข้อมูลสำเร็จ')
	except Exception as e:
		#print(e)
		messagebox.showerror('Error','กรุณาเลือก QTY เป็นตัวเลขเท่านั้น')

Tab = ttk.Notebook(root)
T1 = Frame(Tab)
T2 = Frame(Tab)
T3 = Frame(Tab)
T4 = Frame(Tab)
Tab.pack(fill=BOTH,expand=1)

icon_t1 = PhotoImage(file='T1.png') # .subsample(2) ย่อขนาดลง2เท่าใช้ได้กับรูป png เท่านั้น
icon_t2 = PhotoImage(file='T2.png')
icon_t3 = PhotoImage(file='T3.png')
icon_t4 = PhotoImage(file='T4.png')
icon_b1 = PhotoImage(file='button_save.png')

Tab.add(T1,text=f'{"Writer":^{30}}',image=icon_t1,compound='top')
Tab.add(T2,text=f'{"Table Fault":^{30}}',image=icon_t2,compound='top')
Tab.add(T3,text=f'{"Station":^{30}}',image=icon_t3,compound='top')
Tab.add(T4,text=f'{"Table Station":^{30}}',image=icon_t4,compound='top')
'''
bg = PhotoImage(file='landscape.png')
my_label = Label(T1,image=bg)
my_label.place(x=0,y=0,relwidth=1,relheight=1)
'''

F1 = Frame(T1)
F2 = Frame(T2)
F3 = Frame(T3)
F4 = Frame(T4)
F1.pack()
#F1.place(x=220,y=50) # control ระยะ
F2.pack()
F3.pack()
F4.pack()
############### สร้าง TAB ###################
FONT1 = (None,18) # None เปลี่ยนเป็น 'Angsana New'

#############  Main Photo T1 #############

Main_icon = PhotoImage(file='MainiconT1.png')
Mainicon = Label(F1,image=Main_icon)
Mainicon.pack()

 ############## T1 ###############
L1 = ttk.Label(F1,text=f'{"Work order":^{15}}',font=FONT1,foreground='green')
L1.pack(ipadx=15)

E1_work = StringVar()
E1 = ttk.Entry(F1,textvariable=E1_work,font=FONT1)
E1.pack(ipadx=27)

L2 = ttk.Label(F1,text=f'{"Time":^{20}}',font=FONT1,foreground='green')
L2.pack(ipadx=15)

E2_time = StringVar()
E2 = ttk.Entry(F1,textvariable=E2_time,font=FONT1)
E2.pack(ipadx=27)

############## day ###############
days = StringVar()
dayschoosen = ttk.Combobox(F1, width = 50, 
                            textvariable = days,state='readonly')
  
dayschoosen['values'] = ('days', 
                          '01','02','03','04','05','06','07','08','09','10',
                          '11','12','13','14','15','16','17','18','19','20',
                          '21','22','23','24','25','26','27','28','29','30','31')
dayschoosen.pack(pady=7)
dayschoosen.current(0)

############## months ###############
months = StringVar()
monthschoosen = ttk.Combobox(F1, width = 50, 
                            textvariable = months,state='readonly')
monthschoosen['values'] = ('months','01','02','03','04','05','06','07','08','09','10','11','12')
monthschoosen.pack(pady=2)
monthschoosen.current(0)

############## years ###############
years = StringVar()
yearschoosen = ttk.Combobox(F1, width = 50, 
                            textvariable = years,state='readonly')
  
yearschoosen['values'] = ('years','2020','2021','2022','2023','2024', '2025', '2026', '2027', '2028', 
                          '2029', '2030')

yearschoosen.pack(pady=2)
yearschoosen.current(0)

############## Station ###############
Station = StringVar()
Stationchoosen = ttk.Combobox(F1, width = 50, 
                            textvariable = Station,state='readonly')
  
Stationchoosen['values'] = ('Station', 'E1','E4','E5','E6','E9','CEN','N2','N3','S2','S3','S5')
Stationchoosen.pack(pady=2)
Stationchoosen.current(0)

############## Bound ###############
Bound = StringVar()
Boundchoosen = ttk.Combobox(F1, width = 50,textvariable=Bound,state='readonly')
Boundchoosen['values'] = ('Bound', 
                          'EB',
                          'NB',
                          'SB',
                          'WB')
Boundchoosen.pack(pady=2)
Boundchoosen.current(0)

############## Door ###############
Doors = StringVar()
Doorchoosen = ttk.Combobox(F1, width = 50, 
                            textvariable = Doors,state='readonly')
Doorchoosen['values'] = ('Door', 
                          'D01','D02','D03','D04','D05','D06','D07',
                          'D08','D09','D10','D11','D12','D13','D14',
                          'D15','D16','D17','D18','D19','D20','D21','D22','D23','D24')
Doorchoosen.pack(pady=2)
Doorchoosen.current(0)
Failure = StringVar()
Failurechoosen = ttk.Combobox(F1, width = 50,textvariable=Failure,state='readonly')
Failurechoosen['values'] = ('Failure log',
					 'AMC_M: Obstacle Detection ,DMC:Obstacle Detection inconsistency between DMC_AMC M and AMC S',
                          'AMC_S: Obstacle Detection',
                          'AMC_S: Reset AMC_M:Reset',
                          'DMC:ASD close too slow')
Failurechoosen.pack(pady=2)
Failurechoosen.current(0)
############## Cause ###############
Cause = StringVar()
Causeechoosen = ttk.Combobox(F1, width = 50,textvariable=Cause,state='readonly')
Causeechoosen['values'] = ('Cause',
					 'Software error',
					 'The door(m) obstacle',
					 'The door(m)not open',
					 'The door(s)not open',
					 'The door(s)not closed',
                          'The door not open',
                          'The door closed too slow')
Causeechoosen.pack(pady=2)
Causeechoosen.current(0)

############## Maintenance ###############
Maintenance = StringVar()
Maintenancechoosen = ttk.Combobox(F1, width = 50,textvariable=Maintenance,state='readonly')
Maintenancechoosen['values'] = ('Action', 
                          'Reset DCU',
                          'Replace DCU')
Maintenancechoosen.pack(pady=2)
Maintenancechoosen.current(0)
############## QTY ###############
qty = StringVar()
qtychoosen = ttk.Combobox(F1, width = 50,textvariable=qty,state='readonly')
qtychoosen['values'] = ('QTY',
                          '0',
                          '1')
qtychoosen.pack(pady=2)
qtychoosen.current(0)

############## T2 ###############

LT2 = ttk.Label(F2,text=f'{"ตารางรางแสดงข้อมูล":>{5}}',font=FONT1,foreground='green')
LT2.pack(pady=20)

Main_icon2 = PhotoImage(file='MainiconT2.png')
Mainicon2 = Label(F2,image=Main_icon2)
Mainicon2.pack()

s = ttk.Style(F2)
s.theme_use("clam")

s.configure(".",font=('Angsana New',14))
s.configure("Treeview.Heading",foreground='red',font=('Helvetica',8,"bold"))


header = ['transactionid','Date','Station','Bound','Door','Time','Failure log','Cause','Resolution','Work order','QTY'] # สร้างHeader
headerwidth = [120,70,50,50,45,50,550,200,110,80,30]

resulttable = ttk.Treeview(F2,columns=header,show='headings',height=13) # สร้างTreeview height = 10 คือ จำนวนบรรทัดใน Treeview
resulttable.pack(pady=10)
 
for h in header:
	resulttable.heading(h,text=h) # นำ ข้อมูลใน list header ไปใส่ใน Treeview

for h,w in zip(header,headerwidth):
	resulttable.column(h,width=w) # กำหนดระยะ headerwidth เข้ากับ header โดยการ zip

#resulttable.insert('','end',value=['31/08/2021','CEN','EB','D10','10:30:21','AMC_S: Obstacle Detection','The door can not open','Reset DCU',
				#			'600100200','1'])  # ถ้าเป็น end อังคาร์จะขึ้นก่อนในตาราง

hsb = ttk.Scrollbar(F2,orient="horizontal")
hsb.configure(command=resulttable.xview)
resulttable.configure(xscrollcommand=hsb.set)
hsb.pack(fill=X,side=BOTTOM)

alltransection = {}
alltransectionstation={}

def UpdateSQL():
	data = list(alltransection.values())
	#print(data)
	#print('UPDATE SQL:',data[0]) # โชว์แค่ 1 record
	#print('dataupdata',data)
	for d in data:
		# transectionid,title,expense,quantity,total
		# d[0] = 202108300144088343,d[1]= จันทร์-2021-08-30 01:44:52,d[2]มะม่วง,d[3]=30,d[4]=2,d[5]60.0
		####### เราต้องการเปลี่ยนแค่ d0,2,3,4,5
		update_expense(d[0],d[1],d[2],d[3],d[4],d[5],d[6],d[7],d[8],d[9],d[10]) #ไปเรียก function update_expense มีจำนวน 6 ฟิว ใน database 


############################## F3 ################################
LT2 = ttk.Label(F3,text=f'{"Station":^{10}}',font=FONT1,foreground='red')
LT2.grid(row=0,column=0,pady=30)

LT3 = ttk.Label(F3,text=f'{"QTY":^{10}}',font=FONT1,foreground='red')
LT3.grid(row=0,column=1,pady=30)

LT4 = ttk.Label(F3,text=f'{"Week":^{10}}',font=FONT1,foreground='red')
LT4.grid(row=0,column=2,pady=30)

StationN2 = StringVar()
Station_N2 = ttk.Combobox(F3, width = 50, 
                            textvariable = StationN2,state='readonly')
  
Station_N2['values'] = ('N2')
Station_N2.grid(row=1,column=0,padx=10,pady=20)
Station_N2.current(0)

StationN3 = StringVar()
Station_N3 = ttk.Combobox(F3, width = 50, 
                            textvariable = StationN3,state='readonly')
  
Station_N3['values'] = ('N3')
Station_N3.grid(row=2,column=0,padx=1,pady=10)
Station_N3.current(0)

StationE1 = StringVar()
Station_E1 = ttk.Combobox(F3, width = 50, 
                            textvariable = StationE1,state='readonly')
  
Station_E1['values'] = ('E1')
Station_E1.grid(row=3,column=0,padx=1,pady=10)
Station_E1.current(0)

StationE4 = StringVar()
Station_E4 = ttk.Combobox(F3, width = 50, 
                            textvariable = StationE4,state='readonly')
  
Station_E4['values'] = ('E4')
Station_E4.grid(row=4,column=0,padx=1,pady=10)
Station_E4.current(0)

StationE5 = StringVar()
Station_E5 = ttk.Combobox(F3, width = 50, 
                            textvariable = StationE5,state='readonly')
  
Station_E5['values'] = ('E5')
Station_E5.grid(row=5,column=0,padx=1,pady=10)
Station_E5.current(0)

Station6 = StringVar()
Station_E6 = ttk.Combobox(F3, width = 50, 
                            textvariable = Station6,state='readonly')
  
Station_E6['values'] = ('E6')
Station_E6.grid(row=6,column=0,padx=1,pady=10)
Station_E6.current(0)

Station9 = StringVar()
Station_E9 = ttk.Combobox(F3, width = 50, 
                            textvariable = Station9,state='readonly')
  
Station_E9['values'] = ('E9')
Station_E9.grid(row=7,column=0,padx=1,pady=10)
Station_E9.current(0)

StationS2 = StringVar()
Station_s2 = ttk.Combobox(F3, width = 50, 
                            textvariable = StationS2,state='readonly')
  
Station_s2['values'] = ('S2')
Station_s2.grid(row=8,column=0,padx=1,pady=10)
Station_s2.current(0)

StationS3 = StringVar()
Station_s3 = ttk.Combobox(F3, width = 50, 
                            textvariable = StationS3,state='readonly')
  
Station_s3['values'] = ('S3')
Station_s3.grid(row=9,column=0,padx=1,pady=10)
Station_s3.current(0)

StationS5 = StringVar()
Station_s5 = ttk.Combobox(F3, width = 50, 
                            textvariable = StationS5,state='readonly')
  
Station_s5['values'] = ('S5')
Station_s5.grid(row=10,column=0,padx=1,pady=10)
Station_s5.current(0)

StationCEN = StringVar()
Station_CEN = ttk.Combobox(F3, width = 50, 
                            textvariable = StationCEN,state='readonly')
  
Station_CEN['values'] = ('CEN')
Station_CEN.grid(row=11,column=0,padx=1,pady=10)
Station_CEN.current(0)


QTYTN2 = StringVar()
QTY_N2 = ttk.Combobox(F3, width = 50, 
                            textvariable = QTYTN2,state='readonly')

QTY_N2['values'] = ('QTY_N2', 
                          '0','1','2','3','4','5')
QTY_N2.grid(row=1,column=1,padx=10,pady=20)
QTY_N2.current(0)


QTYN3 = StringVar()
QTY_N3 = ttk.Combobox(F3, width = 50, 
                            textvariable = QTYN3,state='readonly')

QTY_N3['values'] = ('QTY_N3', 
                          '0','1','2','3','4','5')
QTY_N3.grid(row=2,column=1,padx=1,pady=10)
QTY_N3.current(0)

QTYE1 = StringVar()
QTY_E1 = ttk.Combobox(F3, width = 50, 
                            textvariable = QTYE1,state='readonly')

QTY_E1['values'] = ('QTY_E1', 
                          '0','1','2','3','4','5')
QTY_E1.grid(row=3,column=1,padx=1,pady=10)
QTY_E1.current(0)

QTYE4 = StringVar()
QTY_E4 = ttk.Combobox(F3, width = 50, 
                            textvariable = QTYE4,state='readonly')

QTY_E4['values'] = ('QTY_E4', 
                          '0','1','2','3','4','5')
QTY_E4.grid(row=4,column=1,padx=1,pady=10)
QTY_E4.current(0)

QTYE5 = StringVar()
QTY_E5 = ttk.Combobox(F3, width = 50, 
                            textvariable = QTYE5,state='readonly')

QTY_E5['values'] = ('QTY_E5', 
                          '0','1','2','3','4','5')
QTY_E5.grid(row=5,column=1,padx=1,pady=10)
QTY_E5.current(0)

QTYE6 = StringVar()
QTY_E6 = ttk.Combobox(F3, width = 50, 
                            textvariable = QTYE6,state='readonly')

QTY_E6['values'] = ('QTY_E6', 
                          '0','1','2','3','4','5')
QTY_E6.grid(row=6,column=1,padx=1,pady=10)
QTY_E6.current(0)

QTYE9 = StringVar()
QTY_E9 = ttk.Combobox(F3, width = 50, 
                            textvariable = QTYE9,state='readonly')

QTY_E9['values'] = ('QTY_E9', 
                          '0','1','2','3','4','5')
QTY_E9.grid(row=7,column=1,padx=1,pady=10)
QTY_E9.current(0)

QTYS2 = StringVar()
QTY_S2 = ttk.Combobox(F3, width = 50, 
                            textvariable = QTYS2,state='readonly')

QTY_S2['values'] = ('QTY_S2', 
                          '0','1','2','3','4','5')
QTY_S2.grid(row=8,column=1,padx=1,pady=10)
QTY_S2.current(0)

QTYS3 = StringVar()
QTY_S3 = ttk.Combobox(F3, width = 50, 
                            textvariable = QTYS3,state='readonly')

QTY_S3['values'] = ('QTY_S3', 
                          '0','1','2','3','4','5')
QTY_S3.grid(row=9,column=1,padx=1,pady=10)
QTY_S3.current(0)

QTYS5 = StringVar()
QTY_S5 = ttk.Combobox(F3, width = 50, 
                            textvariable = QTYS5,state='readonly')

QTY_S5['values'] = ('QTY_S5', 
                          '0','1','2','3','4','5')
QTY_S5.grid(row=10,column=1,padx=1,pady=10)
QTY_S5.current(0)

QTYCEN = StringVar()
QTY_CEN = ttk.Combobox(F3, width = 50, 
                            textvariable = QTYCEN,state='readonly')

QTY_CEN['values'] = ('QTY_CEN', 
                          '0','1','2','3','4','5')
QTY_CEN.grid(row=11,column=1,padx=1,pady=10)
QTY_CEN.current(0)

Week_t3 = StringVar()
Weekstation = ttk.Combobox(F3, width = 50, 
                            textvariable = Week_t3,state='readonly')

Weekstation['values'] = ('Week', 
                          'Week01','Week02','Week03','Week04','Week05','Week06','Week07','Week08','Week09','Week10',
                          'Week11','Week12','Week13','Week14','Week15','Week16','Week17','Week18','Week19','Week20',
                          'Week21','Week22','Week23','Week24','Week25','Week26','Week27','Week28','Week29','Week30','Week31')
Weekstation.grid(row=1,column=2,padx=1,pady=10)
Weekstation.current(0)

########################## F4 #######################
LT4 = ttk.Label(F4,text=f'{"ตารางรางแสดงข้อมูล":>{5}}',font=FONT1,foreground='green')
LT4.pack(pady=20)

Main_icon4 = PhotoImage(file='MainiconT2.png')
Mainicon4 = Label(F4,image=Main_icon2)
Mainicon4.pack()

s = ttk.Style(F4)
s.theme_use("clam")

s.configure(".",font=('Angsana New',14))
s.configure("Treeview.Heading",foreground='red',font=('Helvetica',8,"bold"))


header4 = ['transectionid','Station','QTY','Week'] # สร้างHeader4
header4width = [150,150,150,150]

resulttableT4 = ttk.Treeview(F4,columns=header4,show='headings',height=13) # สร้างTreeview height = 10 คือ จำนวนบรรทัดใน Treeview
resulttableT4.pack(pady=10)
 
for h in header4:
	resulttableT4.heading(h,text=h) # นำ ข้อมูลใน list header4 ไปใส่ใน Treeview

for h,w in zip(header4,header4width):
	resulttableT4.column(h,width=w) # กำหนดระยะ header4width เข้ากับ header4 โดยการ zip

#resulttableT4.insert('','end',value=['31/08/2021','CEN','EB','D10','10:30:21','AMC_S: Obstacle Detection','The door can not open','Reset DCU',
				#			'600100200','1'])  # ถ้าเป็น end อังคาร์จะขึ้นก่อนในตาราง


def Delete(event=None):
	check = messagebox.askyesno('Confirm','คุณต้องการลบข้อมูลหรือไม่ ?')
	try:
		if check == True:

			select = resulttable.selection() # ไปเรียกฟังก์ชั่น พิเศษที่ คลิกใน Treeview
			# print(select)
			data = resulttable.item(select) # ดึง Item ที่เราเลือกมา จากตาราง (((ถ้าอยากได้มากว่า 1 รายการให้ Run for lop)))
			data = data['values'] # ไปดึง values ออกมา ((dic))
			#print(data)
			transectionid = data[0] # ให้ transectionid = รหัสรายการคือ data[0]
			#print(transectionid)
			delete_fuilt(transectionid) ### Delete in DB
			update_table() # Update data ใหม่่ทั้งหมดอัพโนมัติ
		else:
			pass
	except:

		messagebox.showerror('ERROR','กรุณาเลือกรายการที่จะลบ')
def Edit_record():

	try:


		POPUP = Toplevel()
		w = 500 # กว้าง
		h = 670 # สูง

		ws = POPUP.winfo_screenwidth() #screen width เช็คความกว้างของหน้า
		hs = POPUP.winfo_screenheight() #screen height


		x = (ws/2) - (w/2) # ws คือความกว้างของหน้าจอทั้งหมด /2 คือครึ่งหนึ่งคือ CENTER
		y = (hs/2) - (h/2) - 45

		POPUP.geometry(f'{w}x{h}+{x:.0f}+{y:.0f}')
		POPUP.title('Edit')

		 ############## EDIT ###############


		L1 = ttk.Label(POPUP,text=f'{"Work order":^{15}}',font=FONT1,foreground='green')
		L1.grid(row=0,column=0,padx=5,pady=14)

		E1_work = StringVar()
		E1 = ttk.Entry(POPUP,textvariable=E1_work,font=FONT1)
		E1.grid(row=0,column=1,padx=5,pady=14)

		L2 = ttk.Label(POPUP,text=f'{"Time":^{19}}',font=FONT1,foreground='green')
		L2.grid(row=1,column=0,padx=5,pady=14)

		E2_time = StringVar()
		E2 = ttk.Entry(POPUP,textvariable=E2_time,font=FONT1)
		E2.grid(row=1,column=1,padx=5,pady=14)
		#E2.pack(ipadx=27)

		L3 = ttk.Label(POPUP,text=f'{"Date":^{20}}',font=FONT1,foreground='green')
		L3.grid(row=2,column=0,padx=5,pady=14)

		E3_date = StringVar()
		E3 = ttk.Entry(POPUP,textvariable=E3_date,font=FONT1)
		E3.grid(row=2,column=1,padx=5,pady=14)

		L4 = ttk.Label(POPUP,text=f'{"Station":^{20}}',font=FONT1,foreground='green')
		L4.grid(row=3,column=0,padx=5,pady=14)

		E4_station = StringVar()
		E4 = ttk.Entry(POPUP,textvariable=E4_station,font=FONT1)
		E4.grid(row=3,column=1,padx=5,pady=14)

		L5 = ttk.Label(POPUP,text=f'{"Bound":^{20}}',font=FONT1,foreground='green')
		L5.grid(row=4,column=0,padx=5,pady=14)

		E5_bound = StringVar()
		E5 = ttk.Entry(POPUP,textvariable=E5_bound,font=FONT1)
		E5.grid(row=4,column=1,padx=5,pady=14)

		L6 = ttk.Label(POPUP,text=f'{"Door":^{20}}',font=FONT1,foreground='green')
		L6.grid(row=5,column=0,padx=5,pady=14)

		E6_door = StringVar()
		E6 = ttk.Entry(POPUP,textvariable=E6_door,font=FONT1)
		E6.grid(row=5,column=1,padx=5,pady=14)
		#E6.pack(ipadx=27)

		L7 = ttk.Label(POPUP,text=f'{"Failure log":^{20}}',font=FONT1,foreground='green')
		#L7.pack(ipadx=10)
		L7.grid(row=6,column=0,padx=5,pady=14)
		E7_failure = StringVar()
		E7 = ttk.Entry(POPUP,textvariable=E7_failure,font=FONT1)
		#E7.pack(ipadx=27)
		E7.grid(row=6,column=1,padx=5,pady=14)

		L8 = ttk.Label(POPUP,text=f'{"Cause":^{20}}',font=FONT1,foreground='green')
		#L8.pack(ipadx=10)
		L8.grid(row=7,column=0,padx=5,pady=14)

		E8_cause = StringVar()
		E8 = ttk.Entry(POPUP,textvariable=E8_cause,font=FONT1)
		#E8.pack(ipadx=27)
		E8.grid(row=7,column=1,padx=5,pady=14)
		L9 = ttk.Label(POPUP,text=f'{"Resolution":^{20}}',font=FONT1,foreground='green')
		#L9.pack(ipadx=10)
		L9.grid(row=8,column=0,padx=5,pady=14)

		E9_resolution = StringVar()
		E9 = ttk.Entry(POPUP,textvariable=E9_resolution,font=FONT1)
		#E9.pack(ipadx=27)
		E9.grid(row=8,column=1,padx=5,pady=14)

		L10 = ttk.Label(POPUP,text=f'{"QTY":^{20}}',font=FONT1,foreground='green')
		#L10.pack(ipadx=10)
		L10.grid(row=9,column=0,padx=5,pady=14)

		E10_qty = StringVar()
		E10 = ttk.Entry(POPUP,textvariable=E10_qty,font=FONT1)
		#E10.pack(ipadx=27)
		E10.grid(row=9,column=1,padx=5,pady=14)

		def Edit():

			olddata = alltransection[str(transectionid)]
			#print(olddata)

			
			my_workorder  = E1_work.get()
			my_time = E2_time.get()
			my_days = E3_date.get()
			my_station = E4_station.get()
			my_bound = E5_bound.get()
			my_door = E6_door.get()
			my_failure =  E7_failure.get()
			my_cause = E8_cause.get()
			my_resolution = E9_resolution.get()
			my_qty = E10_qty.get()
		
			newdata = (olddata[0],my_days,my_station,my_bound,my_door,my_time,my_failure,my_cause,my_resolution,int(my_workorder),int(my_qty)) # ตำแหน่งที่ 0,1 เราไม่ต้องแก้ไข
			#print(newdata)
			alltransection[str(transectionid)] = newdata
			#print(alltransection)
			UpdateSQL()
			update_table()
			POPUP.destroy() ########### สั่งปิด POPUP ###################
			
		B2 = ttk.Button(POPUP,text=f'{"Save":>{10}}',image=icon_b1,compound='left',command=Edit) #### ให้ไปเรียก function Edit
		B2.grid(row=10,column=1,pady=10)
			
		#global transectionid

		select = resulttable.selection() # ไปเรียกฟังก์ชั่น พิเศษที่ คลิกใน Treeview
		#print(select)
		data = resulttable.item(select) # ดึง Item ที่เราเลือกมา จากตาราง (((ถ้าอยากได้มากว่า 1 รายการให้ Run for lop)))
		data = data['values'] # ไปดึง values ออกมา ((dic))
		transectionid = data[0]
		#print(transectionid)
		############## ดึงข้อมูลเก่ามาใส่ใน ช่องกรอกที่เราจะแก้ไข ######################
		E1_work.set(data[9])
		E2_time.set(data[5])
		E3_date.set(data[1])
		E4_station.set(data[2])
		E5_bound.set(data[3])
		E6_door.set(data[4])
		E7_failure.set(data[6])
		E8_cause.set(data[7])
		E9_resolution.set(data[8])
		E10_qty.set(data[10])

		POPUP.mainloop()
	except:

		POPUP.destroy()
		messagebox.showerror('Error','กรุณาเลือกข้อมูลที่จะแก้ไข')
rightclick = Menu(root,tearoff=0)
rightclick.add_command(label='Edit',command=Edit_record)
rightclick.add_command(label='Delete',command=Delete) # ไปเรียก function Delete
resulttable.bind('<Delete>',Delete) # กดปุ่ม Delete เพื่อลบข้อมูล	

menuber = Menu(root)
root.config(menu=menuber)

# File menu
filemenu = Menu(menuber,tearoff=0) # tearoff=0 ปิดฟังก์ชั่นย่อย
menuber.add_cascade(label='File',menu=filemenu) # add label file menuber
filemenu.add_command(label='Submit Work order',command=Save)
filemenu.add_command(label='Submit Station',command=Save_station)
filemenu.add_command(label='Exit',command=Exit)

Plot = Menu(menuber,tearoff=0)
menuber.add_cascade(label=f'{"Plot Grahp":^{5}}',menu=Plot) # add label file menuber
Plot.add_command(label=f'{"Plot Grahp Failure":^{5}}',command=pivot_table_1) # เทื่อกดปุ่มให้ไปเรียกฟังก์ชั่น pivot_table_1
Plot.add_command(label=f'{"Plot Grahp Station":^{5}}',command=plot_station)

export_excel = Menu(menuber,tearoff=0)
menuber.add_cascade(label=f'{"Export":^{5}}',menu=export_excel) # add label file menuexcel
#export_excel.add_command(label=f'{"Export_failure":^{5}}',command=Export_failure)
#export_excel.add_command(label=f'{"Export_Station":^{5}}',command=Export_station)


helpemenu = Menu(menuber,tearoff=0)
menuber.add_cascade(label=f'{"Help":^{5}}',menu=helpemenu) # add label file menuber
helpemenu.add_command(label=f'{"About":^{5}}',command=About) # เทื่อกดปุ่มให้ไปเรียกฟังก์ชั่น About



def menupopup(event=None): # ใส่ Event ด้วยจ๊ะ

	if left_click == True: ######### เดี๋ยวมาทำทีหลัง ทำเอง คลิก ซ้ายเลือกก่อนที่จะแสดง POP UP

		# print(event.x_root,event.y_root) # บอกตำแหน่งของแนวแกน x y 
		rightclick.post(event.x_root,event.y_root) # บอกตำแหน่งของแนวแกน x y  ที่คลิกใน resulttable
resulttable.bind('<Button-3>',menupopup) # มีการคลิกขวาที่ตาราง resulttable ให้แสดงข้อมูลในfunction menupopup , Button-3 คือคลิก ขวา
##################### Right Click Menu ###########################

left_click = False

def leftclick(event=None): 
	global left_click
	left_click = True   ######### เดี๋ยวมาทำทีหลัง ทำเอง คลิก ซ้ายเลือกก่อนที่จะแสดง POP UP
	#print(left_click1)

resulttable.bind('<Button-1>',leftclick)

update_table()
update_table_T4()
root.mainloop()