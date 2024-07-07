# pylint: disable = line-too-long, unused-argument, subprocess-run-check
# pylint: disable = missing-module-docstring, import-error, no-name-in-module, ungrouped-imports
import sys
import os
from time import sleep
from subprocess import run, PIPE
from psutil import win_service_get
from win32com.client import Dispatch
from PIL import Image, ImageDraw
from win32com.shell.shell import ShellExecuteEx
from pystray import Icon, Menu, MenuItem

# To run:
# python -m pip install -r .\src\requirements.txt
# pyhton src/phc.py

# To compile:
# python -m pip install pyinstaller
# # old # python -m PyInstaller --onefile src/phc.py --noconsole --name PickHackChew --version-file src/phc.rs -i src/phc.ico
# python -m PyInstaller src/phc.spec

# Statup lnk file
LNK_PATH = os.path.join(os.getenv('APPDATA'),r"Microsoft\Windows\Start Menu\Programs\Startup\PickHackChew.lnk")
# SDDL (Service Descriptor Definition Language) that give all rights to AU (Authenticated Users)
MAGIC_SDDL = '(A;;CCLCSWRPWPDTLOCRRC;;;AU)'

def get_exec_path():
    """Return path to currently running executable, None if run as python script."""
    if getattr(sys, 'frozen', False):
        return sys.executable
    return None

def is_run_at_startup_enabled():
    """Check if run at startup is enabled by checking lnk file presence."""
    return os.path.isfile(LNK_PATH)

def run_at_startup(enable):
    """Enable or disable run at session startup."""
    if is_run_at_startup_enabled():
        os.remove(LNK_PATH)
    if enable:
        exec_path = get_exec_path()
        if exec_path:
            shell = Dispatch('WScript.Shell')
            shortcut = shell.CreateShortCut(LNK_PATH)
            shortcut.Targetpath = exec_path
            shortcut.WorkingDirectory = os.path.dirname(exec_path)
            shortcut.save()
    reflesh_status()

def get_service(name):
    """Get the specified service as psutil.service object."""
    service = None
    try:
        service = win_service_get(name)
    except Exception as ex:
        # raise psutil.NoSuchProcess if no service with such name exists
        print(str(ex))
    return service

def is_service_running(name):
    """Check if the specified service is running"""
    service = get_service(name)
    return (service and service.status() == 'running')

def has_AU_rights_for(name):
    """Check if right has been granted to AU (AuthenticatedUsers) to start and stop the specified service."""
    process = run(["sc", 'sdshow', name], stdout=PIPE, text=True)
    if MAGIC_SDDL in process.stdout.rstrip():
        return True
    return False

def service_set_status(name, up):
    """Start or stop specified service."""
    if not has_AU_rights_for(name):
        set_rights(True)
    if up and get_service_starttype(name) == "Disabled":
        set_service_starttype(name,'demand')
    run(["net", ("stop","start")[up], name], stdout=PIPE, text=True)

def reflesh_status():
    """Refresh all status in case something had been changed by somthing else than this software."""
    global states
    global my_systray_icon
    states = {'stAgentSvc': is_service_running('stAgentSvc'), 'pangps': is_service_running('pangps'), 'webclient': is_service_running('webclient'), 'startup': is_run_at_startup_enabled()}
    if my_systray_icon:
        my_systray_icon.icon=create_image(64, 64, 'red', 'blue',[states['stAgentSvc'],states['pangps'],not states['webclient']])

def on_service_clicked(icon, item, service):
    """Toggle service status"""
    service_set_status(service,not item.checked)
    icon.notify(f'service {service} is now {('stopped','started')[states[service]]}')
    reflesh_status()

def create_image(width, height, color1, color2, list_of_band):
    image = Image.new('RGB', (width, height), color1)
    dc = ImageDraw.Draw(image)
    for i in range(len(list_of_band)):
        if list_of_band[i]:
            dc.rectangle(
                (0, height*i//len(list_of_band), width, height*(i+1)//len(list_of_band)),
                fill=color2)
    return image

def set_rights(give):
    AU_rights = ("",MAGIC_SDDL)[give]
    webclient_starttype = ("disabled","demand")[give]
    ShellExecuteEx(lpVerb='runas', lpFile='cmd.exe', lpParameters='/c '+'sc.exe sdset stAgentSvc "D:(A;;CCLCSWRPLORCWD;;;BA)'+AU_rights+
                         '(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCLCSWLOCRRC;;;IU)(A;;CCLCSWLOCRRC;;;SU)S:(AU;FA;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;WD)" && sc.exe sdset pangps "D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)'+AU_rights+
                         '(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;IU)(A;;CCLCSWLOCRRC;;;SU)S:(AU;FA;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;WD)" && sc.exe sdset webclient "D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)'+AU_rights+
                         '(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;IU)(A;;CCLCSWLOCRRC;;;SU)" && sc.exe config webclient start= '+webclient_starttype)
    sleep(1)

def get_service_starttype(service):
    process = run(["powershell", f"Get-Service -Name {service} | select -ExpandProperty starttype"], stdout=PIPE, text=True)
    return process.stdout.strip()

def set_service_starttype(service, starttype):
    ShellExecuteEx(lpVerb='runas', lpFile='cmd.exe', lpParameters=f'/c sc.exe config {service} start= '+starttype)

def be_free(icon,item):
    service_set_status('pangps',False)
    service_set_status('stAgentSvc',False)
    service_set_status('webclient',True)
    icon.notify('You\'re free.')
    reflesh_status()

def be_corporate(icon,item):
    service_set_status('webclient',False)
    service_set_status('stAgentSvc',True)
    service_set_status('pangps',True)
    icon.notify('Network complient.')
    reflesh_status()

def victor_the_cleaner(icon,item):
    service_set_status('webclient',False)
    service_set_status('stAgentSvc',True)
    service_set_status('pangps',True)
    set_rights(False)
    run_at_startup(False)
    icon.notify('Ass clean.')
    terminate(icon,item)

def abort_shutdown():
    run(['shutdown','-a'], stdout = PIPE)

def terminate(icon, item):
    icon.stop()
    sys.exit(0)

####

my_systray_icon=None
reflesh_status()
my_menu = Menu(
        MenuItem(
            'Be free !!',
            be_free),
        MenuItem(
            'Be corporate !!',
            be_corporate),
        Menu.SEPARATOR,
        MenuItem(
            'Netscope',
            lambda icon, item: on_service_clicked(icon, item, 'stAgentSvc'),
            checked=lambda item: states['stAgentSvc']),
        MenuItem(
            'Global Protect',
            lambda icon, item: on_service_clicked(icon, item, 'pangps'),
            checked=lambda item: states['pangps']),
        MenuItem(
            'Webclient',
            lambda icon, item: on_service_clicked(icon, item, 'webclient'),
            checked=lambda item: states['webclient']),
        Menu.SEPARATOR,
        MenuItem(
            'Run at startup',
            lambda icon, item: run_at_startup(not item.checked),
            checked=lambda item: states['startup']),
        Menu.SEPARATOR,
        MenuItem(
            'Abort planned Shutdown',
            lambda icon, item: abort_shutdown()),
        Menu.SEPARATOR,
        MenuItem(
            'Erase the evidences && Exit',
            victor_the_cleaner),
        Menu.SEPARATOR,
        MenuItem(
            'Exit',
            terminate))
my_systray_icon = Icon(
    'test name',
    icon=None,
    menu=my_menu)
reflesh_status()
my_systray_icon.run()
