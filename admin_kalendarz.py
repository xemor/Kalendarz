from pickle import load as pickle_load
from os.path import join, exists, abspath
from googleapiclient.discovery import build
from tkinter import *
from tkinter import ttk, messagebox
from tkcalendar import Calendar, DateEntry
from datetime import datetime, timezone, timedelta
from threading import Timer, Thread
from sys import exit
import os
from webbrowser import open as web_open
from babel.numbers import *
from infi.systray import SysTrayIcon#tray
import configparser
import ctypes
SCOPES = ['https://www.googleapis.com/auth/calendar']
event = {
  'summary': '',
  'description': '',
  'start': {
    'dateTime': '2015-05-28T09:00:00+01:00',
  },
  'end': {
    'dateTime': '2015-05-28T17:00:00+01:00',
  }
}

def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

def add_event(event):
	global label_newest_event
	global button_delete_newest_event
	global newest_event_id
	try:
		utc_dt = datetime.now(timezone.utc)
		time_zone = utc_dt.astimezone().isoformat()[-6:]
		person_name = raw_person_name.get()
		if person_name != "": person_name =  "Osoba prowadząca: " + person_name + '\n'
		event_name = raw_entry_name.get()
		event_eq = raw_entry_eq.get()
		if event_eq != "": event_eq = "Sprzęt: " + event_eq + '\n'
		event_type = combobox_type.get()
		event_agenda = combobox_agenda.get()
		if (event_name == "" or person_name =="" or event_agenda ==""):
			raise Exception
		event['summary'] = 'x' + event_agenda + ' - ' + event_name
		event['description'] = person_name + event_eq + event_type
		start_dateTime = datetime.strptime(entry_start.get()+combobox_start_time.get(),'%d.%m.%Y%H:%M')
		event['start']['dateTime'] = start_dateTime.strftime('%Y-%m-%dT%H:%M:00'+time_zone)

		end_dateTime = datetime.strptime(entry_end.get()+combobox_end_time.get(),'%d.%m.%Y%H:%M')
		event['end']['dateTime'] = end_dateTime.strftime('%Y-%m-%dT%H:%M:00'+time_zone)

		event = service.events().insert(calendarId='primary', body=event).execute()
		entry_name.delete(0,END)
		entry_eq.delete(0,END)
		entry_person_name.delete(0,END)
		combobox_agenda.delete(0,END)
		combobox_type.delete(0,END)

	
		label_newest_event.destroy()
		button_delete_newest_event.destroy()

		label_newest_event = Label(latest_events_frame, wraplength=550, text=entry_start.get()+' - '+combobox_start_time.get()+' - '+combobox_end_time.get()+'\n'+ event_agenda + ' - ' + event_name+'\n' + person_name + event_eq + event_type, padx=10,pady=10,background='dodger blue',borderwidth=2,relief='solid')
		label_newest_event.grid(column=0,row=0,columnspan=2,sticky='ew')

		newest_event_id = event.get('id')
		button_delete_newest_event = Button(tab1,text='Usuń wydarzenie',command=lambda: delete_event(newest_event_id ,button_delete_newest_event,label_newest_event))
		button_delete_newest_event.grid(column=1,row=9,pady=10)
		show_ten_events()

	except Exception as e:
		messagebox.showerror("","Wprowadź poprawne dane!")

def delete_event(event_id, bde=None, ae = None):
	global newest_event_id
	service.events().delete(calendarId='primary', eventId=event_id).execute()
	if(ae!=None):
		ae.destroy()
	if(bde!=None):
		bde.destroy()
	elif event_id == newest_event_id:
		label_newest_event.destroy()
		button_delete_newest_event.destroy()
		newest_event_id = ''
	show_ten_events()

def time_selected(e):
	combobox_end_time.set((datetime.strptime(combobox_start_time.get(), '%H:%M')+timedelta(hours=1)).strftime('%H:%M'))

def day_selected(e):
	entry_end.set_date(entry_start.get_date())

def show_ten_events(e=None):
	global event_index
	global button_delete_event
	event_index=0
	all_events_text=''

	for label in labels:
		label.destroy()
	del labels[:]
	button_delete_event.destroy()
	
	now = datetime.utcnow().isoformat() + 'Z'

	if combobox_agenda_filtr.get() == '[Wszystkie]':
		events_result = service.events().list(calendarId='primary', timeMin=now,
											maxResults=combobox_amount.get(), singleEvents=True,
											orderBy='startTime').execute()
	else:
		searched = 'x'+combobox_agenda_filtr.get()
		events_result = service.events().list(calendarId='primary', timeMin=now,
											  maxResults=combobox_amount.get(), singleEvents=True,
											  orderBy='startTime', q=searched).execute()


	events = events_result.get('items', [])
	if not events:
		scrolly.grid_forget()
		labels.append(Label(events_frame,text="Brak nadchodzących wydarzeń.",padx=126,pady=10,background='dodger blue',borderwidth=2,relief='solid'))
		labels[0].grid(column=0,row=0,sticky='we',pady=230)
	else:
		for event in events:

			start = event['start'].get('dateTime')
			end = event['end'].get('dateTime')
			if ":" == start[-3]:
				start = start[:-3]+start[-2:]
			if ":" == end[-3]:
				end = end[:-3]+end[-2:]
			start = datetime.strptime(start, '%Y-%m-%dT%H:%M:%S%z')#Tu pobierac timezone zamiast na sztywno +1
			end = datetime.strptime(end, '%Y-%m-%dT%H:%M:%S%z')
			text = str(start.date()) +" - " + str(end.date()) + "\n" + str(start.time())[:-3] + " - " + str(end.time())[:-3]+ "\n"+event['summary'][1:]+ "\n" +event['description']
			#Nie rozróżnia jeszcze dni
			delta = (datetime.strptime(start.strftime('%Y-%m-%d %H:%M'),'%Y-%m-%d %H:%M')-datetime.strptime(datetime.now().strftime('%Y-%m-%d %H:%M'),'%Y-%m-%d %H:%M'))
			
			if delta.days ==0 and delta.seconds/60 == 30 and var_half_hour.get() == 1: 
				Thread(target=lambda:ctypes.windll.user32.MessageBoxW(0, event['summary'][1:]+ "\n" +event['description'], 'Nadchodzące wydarzenie:', 4096)).start()

			labels.append(Radiobutton(canvas,width=56, cursor='hand2',variable=v,value=event_index, indicatoron = 0,text=text,background='deep sky blue',selectcolor='dodger blue',activebackground='deep sky blue',wraplength=380))
			
			canvas.create_window(1, event_index*150, anchor='center', window=labels[event_index], height=147)
			event_index = event_index + 1
			
			if datetime.now().day == start.day:
				all_events_text = all_events_text+text+ '\n----------------------------------\n'

		canvas.configure(scrollregion=canvas.bbox('all'), yscrollcommand=scrolly.set)

		canvas.grid(column=0,row=0)
		scrolly.grid(column=1,row=0,sticky='ns')
	
		button_delete_event = Button(tab2,text='Usuń zaznaczone',command=lambda: delete_event(events[v.get()].get('id')))
		button_delete_event.grid(column=1,row=8)

		#Staff
		button_add_staff = Button(labelframe_staff,text='Dodaj operatora',command=lambda: add_staff(events[v.get()].get('id')))
		button_add_staff.grid(column=0,row=1,sticky='s', padx=13, pady=7)

		button_delete_staff = Button(labelframe_staff,text='Usuń operatora',command=lambda: delete_staff(events[v.get()].get('id')))
		button_delete_staff.grid(column=0,row=2,sticky='n',pady=7,ipadx=2)

		after_start_text.set(all_events_text)
		
def on_closing_root(event=None):
	if tray_quited == 1:
		root.destroy()
		sys.exit(0)
	else:
		root.withdraw()
	

def on_quit_tray(systray):
	global tray_quited
	root.deiconify()
	tray_quited = 1

def display_root(event=None):
	root.deiconify()

def refresh_events():
	global refresh_sec
	refresh_sec = 60
	show_ten_events()
	Timer(60.0, refresh_events).start()

def refresh_info():
	global refresh_sec
	refresh_sec = refresh_sec - 1
	refresh_var.set("Odświeżanie za: %ds"%refresh_sec)
	Timer(1.0,refresh_info).start()

def half_hour_cmd(event=None):
	Config.set('config','half_hour',str(var_half_hour.get()))
	with open("config.ini", 'w') as configfile:
		Config.write(configfile)

def after_start_cmd(event=None):
	Config.set('config','after_start',str(var_after_start.get()))
	with open("config.ini", 'w') as configfile:
		Config.write(configfile)

def add_staff(event_id):
	event = service.events().get(calendarId='primary', eventId=event_id).execute()
	event['description'] = event['description'] + '\n' + 'Operator: ' + combobox_staff.get()
	event['description'] = event['description'].replace('\n\n', '\n')
	updated_event = service.events().update(calendarId='primary', eventId=event['id'], body=event).execute()
	show_ten_events()

def delete_staff(event_id):
	event = service.events().get(calendarId='primary', eventId=event_id).execute()
	staff_index = event['description'].find("Operator: ")
	if staff_index == -1:
		messagebox.showerror("","Brak operatora!")
	else:
		event['description'] = event['description'][0:staff_index-1]
		updated_event = service.events().update(calendarId='primary', eventId=event['id'], body=event).execute()
		show_ten_events()


#-----------------------------------------------MAIN-----------------------------------------

root = Tk()
root.title('Kalendarz')
root.geometry("695x522")
root.protocol("WM_DELETE_WINDOW", on_closing_root)#tray
style = ttk.Style(root)
style.configure("lefttab.TNotebook", tabposition='wn')
root.withdraw()
menu_options = (("Pokaż", None, display_root),)

token_path = resource_path("token.pickle")
icon_path = resource_path("kalendarze.ico")
root.iconbitmap(icon_path) #Small App Icon

systray = SysTrayIcon(icon_path, "Kalendarz", menu_options,on_quit=on_quit_tray)
systray.start()

tab_control = ttk.Notebook(root,style='lefttab.TNotebook')
#button_calendar = Button(tab_control,text='Kalendarz',command=lambda: webbrowser.open("https://calendar.google.com/calendar/embed?src=xxxxxxxxxxx&ctz=Europe%2FWarsaw"))
tab1 = ttk.Frame(tab_control)
tab2 = ttk.Frame(tab_control)
button_calendar = ttk.Button(tab_control,text='Cały kalendarz',command=lambda: web_open("https://calendar.google.com/calendar/embed?src=xxxxxxxxxx&ctz=Europe%2FWarsaw"))

tab_control.add(tab2,text=f'{"Najbliższe wydarzenia":^20s}')
tab_control.add(tab1,text=f'{"Dodaj wydarzenie":^23s}')
tab_control.add(button_calendar)
tab_control.pack(expand=1,fill="both")
#button_calendar.grid(sticky='s')
button_calendar.pack(side='left',anchor='s',ipadx=20,ipady=20)
all_row = 0
all_padx = 10
#Add Event
#Agenda
raw_entry_agenda = StringVar()
agenda_choices = ['F1','F2','F3','F4','F5','F6','F7']
	
label_agenda = Label(tab1,text="Miejsce wydarzenia:",padx=all_padx,pady=10)
label_agenda.grid(column=0,row=all_row)
combobox_agenda = ttk.Combobox(tab1,width=10)
combobox_agenda['values'] = agenda_choices
combobox_agenda.grid(column=1,row=all_row,sticky='w')
all_row = all_row + 1

#Person name
person_name = Label(tab1,text="Osoba prowadząca:",padx=all_padx,pady=10)
person_name.grid(column=0,row=all_row)
raw_person_name = StringVar()
entry_person_name = Entry(tab1,textvariable=raw_person_name,width=30)
entry_person_name.grid(column=1,row=all_row,sticky='w')
all_row = all_row + 1

#Event name
label_name = Label(tab1,text="Nazwa wydarzenia:",padx=all_padx,pady=10)
label_name.grid(column=0,row=all_row)
raw_entry_name = StringVar()
entry_name = Entry(tab1,textvariable=raw_entry_name,width=30)
entry_name.grid(column=1,row=all_row,sticky='w')
all_row = all_row + 1

#Equipment
label_eq = Label(tab1,text="Potrzebny sprzęt:",padx=all_padx,pady=10)
label_eq.grid(column=0,row=all_row)
raw_entry_eq = StringVar()
entry_eq = Entry(tab1,textvariable=raw_entry_eq,width=30)
entry_eq.grid(column=1,row=all_row,sticky='w')
all_row = all_row + 1

#Event Type
raw_entry_type = StringVar()
type_choices = ['Nagranie','Spotkanie Zoom','Spotkanie na żywo']
	
label_type = Label(tab1,text="Rodzaj:",padx=10,pady=10)
label_type.grid(column=0,row=all_row)
combobox_type = ttk.Combobox(tab1,width=20)
combobox_type['values'] = type_choices
combobox_type.grid(column=1,row=all_row,sticky='w')
all_row = all_row + 1

#Event Start Date
label_start_date = Label(tab1,text="Data rozpoczęcia:",padx=10,pady=10)
label_start_date.grid(column=0,row=all_row)
raw_entry_start = StringVar()
entry_start = DateEntry(tab1,width=10,textvariable=raw_entry_start, locale='pl_PL', calendar_cursor='hand2',background='dodger blue',foreground='white',borderwidth=2)
entry_start.grid(column=1,row=all_row,padx=0,pady=10,sticky='w')
all_row = all_row + 1

#Event Start Time
raw_entry_time = StringVar()
time_choices = ['07:00','07:30'
				,'08:00','08:30','09:00','09:30','10:00','10:30','11:00','11:30'
				,'12:00','12:30','13:00','13:30','14:00','14:30','15:00','15:30'
				,'16:00','16:30','17:00','17:30','18:00','18:30','19:00','19:30'
				,'20:00','20:30','21:00','21:30','22:00','22:30','23:00','23:30']
raw_entry_time.set(datetime.strftime(datetime.today(),'%H:%M'))
	
label_start_time = Label(tab1,text="Godzina rozpoczęcia:",padx=10,pady=10)
label_start_time.grid(column=0,row=all_row)
combobox_start_time = ttk.Combobox(tab1,width=10)
combobox_start_time['values'] = time_choices
combobox_start_time.grid(column=1,row=all_row,sticky='w')
combobox_start_time.set(datetime.strftime(datetime.today(),'%H:%M'))
all_row = all_row + 1

#Event End Date
label_end_date = Label(tab1,text="Data zakończenia:",padx=10,pady=10)
label_end_date.grid(column=0,row=all_row)
raw_entry_end = StringVar()
entry_end = DateEntry(tab1,width=10,textvariable=raw_entry_end, locale='pl_PL', calendar_cursor='hand2',background='dodger blue',foreground='white',borderwidth=2)
entry_end.grid(column=1,row=all_row,padx=0,pady=10,sticky='w')

entry_start.bind("<<DateEntrySelected>>",day_selected)
all_row = all_row + 1

#Event End Time
label_end_time = Label(tab1,text="Godzina zakończenia:",padx=10,pady=10)
label_end_time.grid(column=0,row=all_row)
combobox_end_time = ttk.Combobox(tab1,width=10)
combobox_end_time['values'] = time_choices
combobox_end_time.grid(column=1,row=all_row,sticky='w')
combobox_end_time.set((datetime.strptime(combobox_start_time.get(), '%H:%M')+timedelta(hours=1)).strftime('%H:%M'))

combobox_start_time.bind("<<ComboboxSelected>>", time_selected)
all_row = all_row + 1

button_add_event = Button(tab1,text='Dodaj wydarzenie',command=lambda: add_event(event))
button_add_event.grid(column=0,row=all_row,pady=10)

events_frame = LabelFrame(tab2,text='Wydarzenia:')
events_frame.grid(column=0,row=0,pady=1,rowspan=10,sticky='ns')

canvas = Canvas(events_frame, width=400, height=495)
scrolly = Scrollbar(events_frame, orient='vertical', command=canvas.yview)

latest_events_frame = LabelFrame(tab1,text='Ostatnio dodane wydarzenie:')
latest_events_frame.grid(column=0,row=10,columnspan=2,sticky='ew')

label_newest_event = Label(latest_events_frame, wraplength=380, text='Nie dodano wydarzenia' ,padx=215,pady=40,background='dodger blue',borderwidth=2,relief='solid')
label_newest_event.grid(column=0,row=0,sticky='ew')

labelframe_filtr = LabelFrame(tab2,text="Filtry:")
labelframe_filtr.grid(column=1,row=0,pady=1,sticky='n')

#Amount of events
label_amount = Label(labelframe_filtr,text='Liczba wydarzeń:')
label_amount.grid(column=0,row=0,padx=13,sticky='n')
combobox_amount = ttk.Combobox(labelframe_filtr,width=12,state='readonly')
combobox_amount['values'] = [1,2,5,10,15,20]
combobox_amount.grid(column=0,row=0,pady=25)
combobox_amount.set(5)
combobox_amount.bind("<<ComboboxSelected>>",show_ten_events)

#Agenda filtr
agenda_filtr_choices = ['F1','F2','F3','F4','F5','F6','F7']
	
label_agenda_filtr = Label(labelframe_filtr,text="Miejsce wydarzenia:")
label_agenda_filtr.grid(column=0,row=1,sticky='n')
combobox_agenda_filtr = ttk.Combobox(labelframe_filtr,width=12)
combobox_agenda_filtr['values'] = agenda_filtr_choices
combobox_agenda_filtr.grid(column=0,row=1,pady=25)
combobox_agenda_filtr.set('[Wszystkie]')
combobox_agenda_filtr.bind("<<ComboboxSelected>>",show_ten_events)

#Notifications
labelframe_notification = LabelFrame(tab2,text="Powiadomienia:")
labelframe_notification.grid(column=1,row=1,sticky='n')

var_half_hour = BooleanVar()
checkbutton_half_hour = Checkbutton(labelframe_notification, text = 'Pół godziny przed', variable=var_half_hour,command=half_hour_cmd)
checkbutton_half_hour.grid(column=0,row=0)

var_after_start = BooleanVar()
checkbutton_after_start = Checkbutton(labelframe_notification, text = 'Po uruchomieniu\n(plan dnia)', variable=var_after_start,command=after_start_cmd)
checkbutton_after_start.grid(column=0,row=1)

#Staff
labelframe_staff = LabelFrame(tab2,text='Operator:')
labelframe_staff.grid(column=1,row=2,sticky='n', padx=7)
combobox_staff = ttk.Combobox(labelframe_staff,width=12,state='readonly')
combobox_staff['values'] = ['Mateusz','Inny']
combobox_staff.grid(column=0,row=0,sticky='n',pady=5, padx=13)
combobox_staff.set('Mateusz')


#Config file
Config = configparser.ConfigParser()
if not exists("config.ini"):
	f = open("config.ini", "a")
	f.write("[config]\nhalf_hour = True\nafter_start = True")
	f.close()
	
Config.read("config.ini")
var_half_hour.set(Config.get('config', 'half_hour'))
var_after_start.set(Config.get('config', 'after_start'))

#Refresh info
refresh_var= StringVar()
refresh_sec = 60
refresh_var.set("Odświeżanie za: %ds"%refresh_sec)

label_refresh_info = Label(tab2,textvariable=refresh_var)
label_refresh_info.grid(column=1,row=9,sticky='s')



if exists(token_path):
	with open(token_path, 'rb') as token:
		creds = pickle_load(token)
		
		
try:
	service = build('calendar', 'v3', credentials=creds)
except Exception as e:
	
	messagebox.showerror("","Nie udało się połączyć z kalendarzem! Sprawdź połączenie z internetem lub skontaktuj się z administratorem!")
	sys.exit(0)

v = IntVar()
v.set(0)

after_start_text = StringVar()

event_index = 0
labels = []
label_newest_event = Label()
button_delete_newest_event = Button()
newest_event_id = StringVar()
button_delete_event = Button()
refresh_events()
refresh_info()
tray_quited = 0

if var_after_start.get() == 1 and after_start_text.get() != '':
	Thread(target=lambda:ctypes.windll.user32.MessageBoxW(0, after_start_text.get(), 'Dzisiejsze wydarzenia:', 4096)).start()

root.mainloop()
