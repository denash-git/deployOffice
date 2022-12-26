
def wmic(command):
# request command wmic via cmd
    result = str(subprocess.check_output(command, shell=True))
    for i in ["\\r", "\\n", "b", "'"]:
        result = result.replace(i, "")
    result = result.split('=')
    # return list[parameter, value]
    return result

def architecture(name):
#bit request  by name pc
    value = 0
    # /user:setup_app
    command = 'wmic /node:' + name + ' OS get OSArchitecture /VALUE'
    result = wmic(command)
    if '64' in result[1]: value = 64
    if '32' in result[1]: value = 32
    if value == 0: print('error parameter: modul archtechure')
    # return bit value 32 or 64
    return value

def path_work():
# working path current user
    user = os.path.expanduser('~')
    path = fr'{user}\AppData\Local\Temp'
    return path

def xml(arch):
# makefile configure deployment
    path = path_work()
    deploypath = cfg.get('default','deployment path')
    file = open(fr"{path}\task.xml", 'w', encoding='utf-8')
    # required parameter
    file.write('<Configuration ID="2977850e-b900-44b6-ae3e-41a2d5d6becf">\n')
    if arch == 64:
        file.write(fr'  <Add OfficeClientEdition="64" Channel="PerpetualVL2021" SourcePath="{deploypath}\2021\64" AllowCdnFallback="FALSE">\n')
    else:
        file.write(fr'  <Add OfficeClientEdition="32" Channel="PerpetualVL2021" SourcePath="{deploypath}\2021\32" AllowCdnFallback="FALSE">\n')
    # word, excel, powerpoint
    if office_value.get() == 1:
        file.write('    <Product ID="ProPlus2021Volume" PIDKEY="FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH">\n')
        file.write('      <Language ID="ru-ru" />\n      <ExcludeApp ID="Access" />\n      <ExcludeApp ID="OneDrive" />\n      <ExcludeApp ID="Groove" />\n      <ExcludeApp ID="OneNote" />\n      <ExcludeApp ID="Outlook" />\n      <ExcludeApp ID="Lync" />\n      <ExcludeApp ID="Teams" />\n')
        # if YES publisher
        if publisher_value.get() == 0:
            file.write('      <ExcludeApp ID="Publisher" />\n')
        file.write('    </Product>\n    <Product ID="ProofingTools">\n      <Language ID="en-us" />\n      <Language ID="ru-ru" />\n    </Product>\n')
    # viso
    if visio_value.get() == 1:
        file.write('	<Product ID="VisioPro2021Volume" PIDKEY="KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4">\n')
        file.write('      <Language ID="ru-ru" />\n')
        file.write('    </Product>\n')
    # project
    if project_value.get() == 1:
        file.write('    <Product ID="ProjectPro2021Volume" PIDKEY="FTNWT-C6WBT-8HMGF-K9PRX-QV9H8">\n')
        file.write('      <Language ID="ru-ru" />\n')
        file.write('    </Product>\n')
    # required parameter
    file.write('  </Add>\n')
    file.write('  <Property Name="PinIconsToTaskbar" Value="FALSE" />\n')
    file.write('  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />\n')
    file.write('  <Property Name="AUTOACTIVATE" Value="1" />\n')
    file.write('  <Updates Enabled="FALSE" />\n')
    file.write('  <RemoveMSI />\n')
    file.write('  <Display Level="Full" AcceptEULA="TRUE" />\n')
    file.write('</Configuration>')
    file.close()
    return

def deploy(name):
# the deployment process
    login = cfg.get('default','local admin')
    deploypath = cfg.get('default','deployment path')
    userpath = path_work()
    pull = [fr'copy /y "{deploypath}\setup.exe" "\\{name}\admin$\Temp\setup.exe"',
            fr'copy /y "{userpath}\task.xml" "\\{name}\admin$\Temp\task.xml"',
            fr'schtasks /create /s {name} /ru {login} /rp /st 23:00 /sc ONCE /tn "deploy" /rl highest /f /tr "C:\Windows\Temp\setup.exe /configure C:\Windows\Temp\task.xml"',
           f'schtasks /run /s {name} /i /tn "deploy"',
            f'schtasks /delete /f /s {name} /tn "deploy"'
            ]
    for command in pull:
        if "/rp" in command: print(f'Введите пароль {login}:')
        result = subprocess.run(command, shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, encoding='cp866')
        if result.returncode == 0:
            print('Good')
        else:
            print(result.stdout)
    return

# selection checkbox
def click():
    #publisher is active if the office selected
    if office_value.get() == 1:
        publisher.config(state='normal')
    else:
        publisher.config(state='disabled')
        publisher_value.set(0)
    value = 0 #input field and button is active, if checkbox selected
    for enum in [office_value, visio_value, project_value]:
        value = value + enum.get()
    if value > 0:
        btn.config(state='normal')
        name.config(state='normal')
    else:
        btn.config(state='disabled')
        name.config(state='disabled')
    return

def ping(host):
    res = 0
    value = subprocess.run(['ping', '-n', '1', host], stdout=subprocess.PIPE, encoding='cp866')
    for i in ['Заданный узел недоступен', 'Превышен интервал ожидания для запроса']:
        if i in value.stdout:
            res = 1
    return res # 0-up 1-down

def name_pc():
    return name.get()

def start():
# click buttom root window
    name = name_pc()                   # get name pc
    if ping(name) == 0:
        arch = architecture(name)          # get architecture OS
        xml(arch)                          # make XML file
        deploy(name)                       # deployment Office
    else:
        showerror(title="Ошибка", message="Host не доступен")
    return

from tkinter import *
from tkinter.messagebox import showerror
import subprocess
import configparser
import os

# reading default settings from config.ini
cfg = configparser.ConfigParser()
filename = os.path.basename(__file__)
filepath = os.path.abspath(__file__).replace(filename, '')
with open(fr'{filepath}\config.ini') as config:
    cfg.read_file(config)

# root windows
win = Tk()
win.title("deployment MS Office 2021")
win.geometry("450x350")
win.resizable(False, False)

frame1 = Frame(win)
frame2 = Frame(win)
frame3 = Frame(win)

frame1.pack()
frame2.pack(pady=25)
frame3.pack(pady=25)

lbl = Label(frame1, text="Выбирите вариант установки :", font=("Arial", 16, "bold"), height=2)
lbl.pack()

office_value = IntVar()
office = Checkbutton(frame1, text="Word, Excel, PowerPoint", font=("Arial", 14), width=20, anchor='w', onvalue=1,
                         offvalue=0, variable=office_value, command=click)
office.pack()

publisher_value = IntVar()
publisher = Checkbutton(frame1, text="+ Publisher", font=("Arial", 14), width=20, anchor='w', onvalue=1, offvalue=0,
                            variable=publisher_value, state='disabled')
publisher.pack()

visio_value = IntVar()
visio = Checkbutton(frame1, text="+ Visio", font=("Arial", 14), width=20, anchor='w', onvalue=1, offvalue=0,
                        variable=visio_value, command=click)
visio.pack()

project_value = IntVar()
project = Checkbutton(frame1, text="+ Project", font=("Arial", 14), width=20, anchor='w', onvalue=1, offvalue=0,
                          variable=project_value,  command=click)
project.pack()

lbname = Label(frame2, text='Имя ПК: ', font=("Arial", 14))
name = Entry(frame2, text="", font=("Arial", 14), width=10, state='disabled')
lbname.pack(side='left')
name.pack(side='left')

btn = Button(frame3, text="Выполнить", font=("Arial", 14), width=20,  state='disabled', command=start)
btn.pack()

win.mainloop()


