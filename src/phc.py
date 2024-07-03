import pystray, subprocess, sys, os
import psutil, time, winshell
from win32com.client import Dispatch
from PIL import Image, ImageDraw
import win32com.shell.shell as shell

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
        service = psutil.win_service_get(name)
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
    process = subprocess.run(["sc", 'sdshow', name], stdout=subprocess.PIPE, text=True)
    if MAGIC_SDDL in process.stdout.rstrip():
        return True
    return False

def service_set_status(name, up):
    """Start or stop specified service."""
    if not has_AU_rights_for(name):
        setRights(True)
    process = subprocess.run(["net", ("stop","start")[up], name], stdout=subprocess.PIPE, text=True)

def reflesh_status():
    """Refresh all status in case something had been changed by somthing else than this software."""
    global states
    global mysystrayicon
    states = {'stAgentSvc': is_service_running('stAgentSvc'), 'pangps': is_service_running('pangps'), 'webclient': is_service_running('webclient'), 'startup': is_run_at_startup_enabled()}
    if mysystrayicon:
        mysystrayicon.icon=create_image(64, 64, 'red', 'blue',[states['stAgentSvc'],states['pangps'],not states['webclient']])

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

def setRights(give):
    AURights = ("",MAGIC_SDDL)[give]
    webclient_start_behavior = ("disabled","demand")[give]
    shell.ShellExecuteEx(lpVerb='runas', lpFile='cmd.exe', lpParameters='/c '+'sc.exe sdset stAgentSvc "D:(A;;CCLCSWRPLORCWD;;;BA)'+AURights+
                         '(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCLCSWLOCRRC;;;IU)(A;;CCLCSWLOCRRC;;;SU)S:(AU;FA;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;WD)" && sc.exe sdset pangps "D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)'+AURights+
                         '(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;IU)(A;;CCLCSWLOCRRC;;;SU)S:(AU;FA;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;WD)" && sc.exe sdset webclient "D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)'+AURights+
                         '(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;IU)(A;;CCLCSWLOCRRC;;;SU)" && sc.exe config webclient start= '+webclient_start_behavior)
    time.sleep(1)

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
    setRights(False)
    run_at_startup(False)
    icon.notify('Ass clean.')
    terminate(icon,item)

def abort_shutdown():
    process = subprocess.run(['shutdown','-a'], stdout = subprocess.PIPE)

def terminate(icon, item):
    icon.stop()
    sys.exit(0)

####

mysystrayicon=None
reflesh_status()
mymenu = pystray.Menu(
        pystray.MenuItem(
            'Be free !!',
            be_free),
        pystray.MenuItem(
            'Be corporate !!',
            be_corporate),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem(
            'Netscope',
            lambda icon, item: on_service_clicked(icon, item, 'stAgentSvc'),
            checked=lambda item: states['stAgentSvc']),
        pystray.MenuItem(
            'Global Protect',
            lambda icon, item: on_service_clicked(icon, item, 'pangps'),
            checked=lambda item: states['pangps']),
        pystray.MenuItem(
            'Webclient',
            lambda icon, item: on_service_clicked(icon, item, 'webclient'),
            checked=lambda item: states['webclient']),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem(
            'Run at startup',
            lambda icon, item: run_at_startup(not item.checked),
            checked=lambda item: states['startup']),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem(
            'Abort planned Shutdown',
            lambda icon, item: abort_shutdown()),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem(
            'Erase the evidences \\& Exit',
            victor_the_cleaner),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem(
            'Exit',
            terminate))
mysystrayicon = pystray.Icon(
    'test name',
    icon=None,
    menu=mymenu)
reflesh_status()
mysystrayicon.run()
