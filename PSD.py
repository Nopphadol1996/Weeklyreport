from tkinter import *
from tkinter import ttk,messagebox
import csv
import pandas as pd
import matplotlib.pyplot as plt

# basicsqlite3.py

import sqlite3

######### DB #####################
conn = sqlite3.connect('Weekly.db')
# สร้างตัวดำเนินการ (อยากได้อะไรใช้ตัวนี้ได้เลย)
c = conn.cursor()

c.execute("""CREATE TABLE IF NOT EXISTS weeklytable (
				ID INTEGER PRIMARY KEY AUTOINCREMENT,
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

def insert_work(Date_,Station,Bound,Door,Time_,Failre,Cause,Resolution,Work,QTY): # เอาที่เราสร้างมาใส่
	ID = None
	with conn:
		c.execute("""INSERT INTO weeklytable VALUES (?,?,?,?,?,?,?,?,?,?,?)""", # ? ต้องรวม ID = None
			(ID,Date_,Station,Bound,Door,Time_,Failre,Cause,Resolution,Work,QTY)) #ใส่ ID ไปด้วย
		conn.commit() # คือ การบันทึกข้อมูลลงในฐานข้อมูล ถ้าไม่รันตัวนี้จะไม่บันทึก
		print('Insert Sucess...!')

#insert_work('31/01/2021','CEN','EB','D07','10:20:31','AMC_S: Obstacle Detection','Door closed too slow','Reset DCU',600600123,1)

def show_expense():
	with conn:
		c.execute("SELECT *FROM weeklytable")
		expense = c.fetchall() # คำสั่งให้ดึงข้อมูลมา
		# print(expense)
	return expense


def pivot_table_1():
	df =pd.read_sql_query("SELECT * FROM weeklytable",conn,)
	del df['Work']
	del df['ID']
	print(df)


	'''
	df.pivot_table(index=['Station','Bound','Door'], columns ='Cause',values='QTY',fill_value=0,
					).plot(kind='bar',fontsize=15)'''

	df.pivot_table(index=['Station','Bound','Door'], columns ='Cause',values='QTY', aggfunc='count',fill_value=0,
              margins=False, margins_name='Grand Total').plot(kind='bar')


	plt.title('Weekly Report',color='green')
	plt.ylabel('Total Failure',color='green')

	plt.ylim(0,5)

	plt.show()
	print(df)

	'''
	filepath = 'C:/Users/Nopphadol/Desktop/Project_beginer/Myfile.xlsx'
	writer = pd.ExcelWriter(filepath)
	df.to_excel(writer,'Mysheet2',index=False)
	writer.save()
	print('------------')
	'''


######### DB #####################

root = Tk()

root.title('โปรแกรม Weekly Report')


w = 795 # กว้าง
h = 670 # สูง

ws = root.winfo_screenwidth() #screen width เช็คความกว้างของหน้า
hs = root.winfo_screenheight() #screen height

x = (ws/2) - (w/2) # ws คือความกว้างของหน้าจอทั้งหมด /2 คือครึ่งหนึ่งคือ CENTER
y = (hs/2) - (h/2) - 45
root.geometry(f'{w}x{h}+{x:.0f}+{y:.0f}')

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

#	text = '{} {} {} {} {} {} {} {} {} {} '.format(textdate,my_station,my_bound,my_door,my_time,my_failure,my_cause,my_Maintenance,my_workorder,my_qty)
#	print(text)

	
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

		insert_work(textdate,my_station,my_bound,my_door,my_time,my_failure,my_cause,my_Maintenance,int(my_workorder),int(my_qty))
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
		E1.focus()

	except:
		print('โปรดตรวจสอบ:\n Work order ต้องเป็นตัวเลข หรือ\n รูปบแบบวันเวลาต้อง 00:00:00 หรือ\n เลือกจำนวน QTY')
		messagebox.showerror('ERROR','โปรดตรวจสอบ:\n Work order ต้องเป็นตัวเลข หรือ\n รูปบแบบวันเวลาต้อง 00:00:00 หรือ\n เลือกจำนวน QTY')
############### สร้าง TAB ###################
def update_table():

	resulttable.delete(*resulttable.get_children())
	data_db = show_expense()
	#insert_work(textdate,my_station,my_bound,my_door,my_time,my_failure,my_cause,my_Maintenance,int(my_workorder),int(my_qty))
	for d in data_db:

		resulttable.insert('','end',value=d[1:])

Tab = ttk.Notebook(root)
T1 = Frame(Tab)
T2 = Frame(Tab)
Tab.pack(fill=BOTH,expand=1)

icon_t1 = PhotoImage(file='T1_expens.png') # .subsample(2) ย่อขนาดลง2เท่าใช้ได้กับรูป png เท่านั้น
icon_b1 = PhotoImage(file='button_save.png')


Tab.add(T1,text=f'{"Writer":^{30}}',image=icon_t1,compound='top')
Tab.add(T2,text=f'{"Reader":^{30}}',image=icon_t1,compound='top')
'''
bg = PhotoImage(file='landscape.png')
my_label = Label(T1,image=bg)
my_label.place(x=0,y=0,relwidth=1,relheight=1)
'''
F1 = Frame(T1)
F2 = Frame(T2)
#F1.pack()
F1.place(x=190,y=50) # control ระยะ
F2.pack()
############### สร้าง TAB ###################

FONT1 = (None,18) # None เปลี่ยนเป็น 'Angsana New'

#############  Main Photo T1 #############

Main_icon = PhotoImage(file='MainiconT1.png')
Mainicon = Label(F1,image=Main_icon)
Mainicon.pack()


#############  Main Photo T1 #############
 ############## T1 ###############
L1 = ttk.Label(F1,text='Work order',font=FONT1,foreground='green')
L1.pack(ipadx=15)

E1_work = StringVar()
E1 = ttk.Entry(F1,textvariable=E1_work,font=FONT1)
E1.pack(ipadx=27)

L2 = ttk.Label(F1,text='Time',font=FONT1,foreground='green')
L2.pack(ipadx=15)

E2_time = StringVar()
E2 = ttk.Entry(F1,textvariable=E2_time,font=FONT1)
E2.pack(ipadx=27)

############## day ###############
days = StringVar()
dayschoosen = ttk.Combobox(F1, width = 50, 
                            textvariable = days,state='readonly')
  
# Adding combobox drop down list
dayschoosen['values'] = ('days', 
                          '01','02','03','04','05','06','07','08','09','10',
                          '11','12','13','14','15','16','17','18','19','20',
                          '21','22','23','24','25','26','27','28','29','30','31')
dayschoosen.pack(pady=7)
 # Shows february as a default value
dayschoosen.current(0)
############## day ###############

############## months ###############
months = StringVar()
monthschoosen = ttk.Combobox(F1, width = 50, 
                            textvariable = months,state='readonly')
  
# Adding combobox drop down list
monthschoosen['values'] = ('months','01','02','03','04','05','06','07','08','09','10','11','12')
monthschoosen.pack(pady=2)
 # Shows february as a default value
monthschoosen.current(0)
############## months ###############

############## years ###############
years = StringVar()
yearschoosen = ttk.Combobox(F1, width = 50, 
                            textvariable = years,state='readonly')
  
# Adding combobox drop down list
yearschoosen['values'] = ('years','2020','2021','2022','2023','2024', '2025', '2026', '2027', '2028', 
                          '2029', '2030')

yearschoosen.pack(pady=2)
 # Shows february as a default value
yearschoosen.current(0)
############## years ###############

############## Station ###############
Station = StringVar()
Stationchoosen = ttk.Combobox(F1, width = 50, 
                            textvariable = Station,state='readonly')
  
# Adding combobox drop down list
Stationchoosen['values'] = ('Station', 'E1','E4','E5','E6','E9','CEN','N2','N3','S2','S3','S5')
Stationchoosen.pack(pady=2)
 # Shows february as a default value
Stationchoosen.current(0)
############## Station ###############

############## Bound ###############
Bound = StringVar()
Boundchoosen = ttk.Combobox(F1, width = 50,textvariable=Bound,state='readonly')
  
# Adding combobox drop down list
Boundchoosen['values'] = ('Bound', 
                          'EB',
                          'NB',
                          'SB',
                          'WB')
Boundchoosen.pack(pady=2)
 # Shows february as a default value
Boundchoosen.current(0)
############## Bound ###############

############## Door ###############
Doors = StringVar()
Doorchoosen = ttk.Combobox(F1, width = 50, 
                            textvariable = Doors,state='readonly')
  
# Adding combobox drop down list
Doorchoosen['values'] = ('Door', 
                          'D01','D02','D03','D04','D05','D06','D07',
                          'D08','D09','D10','D11','D12','D13','D14',
                          'D15','D16','D17','D18','D19','D20','D21','D22','D23','D24')
Doorchoosen.pack(pady=2)
 # Shows february as a default value
Doorchoosen.current(0)
############## Door ###############

############## Failure ###############
Failure = StringVar()
Failurechoosen = ttk.Combobox(F1, width = 50,textvariable=Failure,state='readonly')
  
# Adding combobox drop down list
Failurechoosen['values'] = ('Failure log', 
                          'AMC_S: Obstacle Detection',
                          'DMC:ASD close too slow')
Failurechoosen.pack(pady=2)
 # Shows february as a default value
Failurechoosen.current(0)
############## Failure ##############

############## Cause ###############
Cause = StringVar()
Causeechoosen = ttk.Combobox(F1, width = 50,textvariable=Cause,state='readonly')
  
# Adding combobox drop down list
Causeechoosen['values'] = ('Cause', 
                          'Door not open',
                          'Door closed too slow')
Causeechoosen.pack(pady=2)
 # Shows february as a default value
Causeechoosen.current(0)
############## Cause ###############

############## Maintenance ###############
Maintenance = StringVar()
Maintenancechoosen = ttk.Combobox(F1, width = 50,textvariable=Maintenance,state='readonly')
  
# Adding combobox drop down list
Maintenancechoosen['values'] = ('Action', 
                          'Reset DCU',
                          'Replace DCU')
Maintenancechoosen.pack(pady=2)
 # Shows february as a default value
Maintenancechoosen.current(0)
############## Maintenance ###############

############## QTY ###############
qty = StringVar()
qtychoosen = ttk.Combobox(F1, width = 50,textvariable=qty,state='readonly')
  
# Adding combobox drop down list
qtychoosen['values'] = ('QTY',
                          '0',
                          '1')
qtychoosen.pack(pady=2)
 # Shows february as a default value
qtychoosen.current(0)

############## QTY ###############
B1 = ttk.Button(F1,text=f'{"Save":>{10}}',image=icon_b1,compound='left',command=Save) #### ให้ไปเรียก function Edit
#B1.place(x=310,y=580)
B1.pack(pady=5)

############## T2 ###############

LT2 = ttk.Label(F2,text='ตารางรางแสดงข้อมูล',font=FONT1,foreground='green')
LT2.pack(pady=20)

header = ['Date','Station','Bound','Door','Time','Failure log','Cause','Resolution','Work order','QTY'] # สร้างHeader
headerwidth = [67,60,60,60,60,160,130,80,80,30]

resulttable = ttk.Treeview(F2,columns=header,show='headings',height=20) # สร้างTreeview height = 10 คือ จำนวนบรรทัดใน Treeview
resulttable.pack(pady=20)
 
for h in header:
	resulttable.heading(h,text=h) # นำ ข้อมูลใน list header ไปใส่ใน Treeview

for h,w in zip(header,headerwidth):
	resulttable.column(h,width=w) # กำหนดระยะ headerwidth เข้ากับ header โดยการ zip

#resulttable.insert('','end',value=['31/08/2021','CEN','EB','D10','10:30:21','AMC_S: Obstacle Detection','The door can not open','Reset DCU',
				#			'600100200','1'])  # ถ้าเป็น end อังคาร์จะขึ้นก่อนในตาราง

B2 = ttk.Button(F2,text=f'{"Plot":>{10}}',image=icon_b1,compound='left',command=pivot_table_1) #### ให้ไปเรียก function Edit
B2.pack(pady=5)


update_table()
root.mainloop()