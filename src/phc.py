# pylint: disable = line-too-long, unused-argument, subprocess-run-check
# pylint: disable = missing-module-docstring, import-error, no-name-in-module, ungrouped-imports
import sys
import os
from time import sleep
from subprocess import run, PIPE, STARTUPINFO, STARTF_USESHOWWINDOW
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
# startupinfo for subprocess.run()
STARTUP_INFO = STARTUPINFO()
STARTUP_INFO.dwFlags = STARTF_USESHOWWINDOW
STARTUP_INFO.wShowWindow = 0
# proposed DNS
DNS_LIST={"CloudFlare":("1.1.1.1","1.0.0.1"),
"Google": ("8.8.8.8", "8.8.4.4"),
"Control D": ("76.76.2.0", "76.76.10.0"),
"Quad9": ("9.9.9.9", "149.112.112.112"),
"OpenDNS Home": ("208.67.222.222", "208.67.220.220"),
"AdGuard DNS": ("94.140.14.14", "94.140.15.15"),
"CleanBrowsing": ("185.228.168.9", "185.228.169.9"),
"Alternate DNS": ("76.76.19.19", "76.223.122.150")}

# ICMP keys
ICMP=["AllowOutboundDestinationUnreachable",
"AllowOutboundSourceQuench",
"AllowRedirect",
"AllowInboundEchoRequest",
"AllowInboundRouterRequest",
"AllowOutboundTimeExceeded",
"AllowOutboundParameterProblem",
"AllowInboundTimestampRequest",
"AllowInboundMaskRequest",
"AllowOutboundPacketTooBig"]

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
    process = run(["sc", 'sdshow', name], stdout=PIPE, text=True, startupinfo=STARTUP_INFO)
    if MAGIC_SDDL in process.stdout.rstrip():
        return True
    return False

def service_set_status(name, up):
    """Start or stop specified service."""
    if not has_AU_rights_for(name):
        set_rights(True)
    if up and get_service_starttype(name) == "Disabled":
        set_service_starttype(name,'demand')
    run(["net", ("stop","start")[up], name], startupinfo=STARTUP_INFO)

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
    reflesh_status()
    icon.notify(f'service {service} is now {('stopped','started')[states[service]]}')

def create_image(width, height, color1, color2, list_of_band):
    image = Image.new('RGB', (width, height), color1)
    dc = ImageDraw.Draw(image)
    for i in range(len(list_of_band)):
        if list_of_band[i]:
            dc.rectangle(
                (0, height*i//len(list_of_band), width, height*(i+1)//len(list_of_band)),
                fill=color2)
    return image

def set_rights(enable):
    AU_rights = ("",MAGIC_SDDL)[enable]
    webclient_starttype = ("disabled","demand")[enable]
    _ = ShellExecuteEx(lpVerb='runas', lpFile='cmd.exe', lpParameters='/c '+'sc.exe sdset stAgentSvc "D:(A;;CCLCSWRPLORCWD;;;BA)'+AU_rights+
                         '(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCLCSWLOCRRC;;;IU)(A;;CCLCSWLOCRRC;;;SU)S:(AU;FA;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;WD)" && sc.exe sdset pangps "D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)'+AU_rights+
                         '(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;IU)(A;;CCLCSWLOCRRC;;;SU)S:(AU;FA;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;WD)" && sc.exe sdset webclient "D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)'+AU_rights+
                         '(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;IU)(A;;CCLCSWLOCRRC;;;SU)" && sc.exe config webclient start= '+webclient_starttype)
    sleep(1)

def get_service_starttype(service):
    process = run(["powershell", "-WindowStyle", "hidden", f"Get-Service -Name {service} | select -ExpandProperty starttype"], stdout=PIPE, text=True, startupinfo=STARTUP_INFO)
    return process.stdout.strip()

def set_service_starttype(service, starttype):
    _ = ShellExecuteEx(lpVerb='runas', lpFile='cmd.exe', lpParameters=f'/c sc.exe config {service} start= '+starttype)

def get_firewall_status():
    process = run(["powershell", "-WindowStyle", "hidden", "Get-ItemProperty -Path 'HKLM:\\SOFTWARE\\Policies\\Microsoft\\WindowsFirewall\\DomainProfile' -name 'EnableFirewall'  | select -ExpandProperty EnableFirewall"], stdout=PIPE, text=True, startupinfo=STARTUP_INFO)
    return process.stdout.strip() == "1"

def set_firewall_status(enable):
    flag = ('0','1')[enable]
    cmd = "/c powershell \""+\
    f"Set-ItemProperty -Path 'HKLM:\\SOFTWARE\\Policies\\Microsoft\\WindowsFirewall\\DomainProfile' -name 'EnableFirewall' -Value {flag}; "+\
    f"Set-ItemProperty -Path 'HKLM:\\SOFTWARE\\Policies\\Microsoft\\WindowsFirewall\\PublicProfile' -name 'EnableFirewall' -Value {flag}; "+\
    f"Set-ItemProperty -Path 'HKLM:\\SOFTWARE\\Policies\\Microsoft\\WindowsFirewall\\StandardProfile' -name 'EnableFirewall' -Value {flag}; "+\
    f"Set-ItemProperty -Path 'HKLM:\\SYSTEM\\CurrentControlSet\\services\\SharedAccess\\Parameters\\FirewallPolicy\\DomainProfile' -name 'EnableFirewall' -Value {flag}; "+\
    f"Set-ItemProperty -Path 'HKLM:\\SYSTEM\\CurrentControlSet\\services\\SharedAccess\\Parameters\\FirewallPolicy\\PublicProfile' -name 'EnableFirewall' -Value {flag}; "+\
    f"Set-ItemProperty -Path 'HKLM:\\SYSTEM\\CurrentControlSet\\services\\SharedAccess\\Parameters\\FirewallPolicy\\StandardProfile' -name 'EnableFirewall' -Value {flag}\""
    _ = ShellExecuteEx(lpVerb='runas', lpFile='cmd.exe', lpParameters=cmd)

def get_ping_allowed():
    process = run(["powershell", "-WindowStyle", "hidden", "Get-ItemProperty -Path 'HKLM:\\SOFTWARE\\Policies\\Microsoft\\WindowsFirewall\\StandardProfile\\IcmpSettings' -name AllowInboundEchoRequest  | select -ExpandProperty AllowInboundEchoRequest"], stdout=PIPE, text=True, startupinfo=STARTUP_INFO)
    return process.stdout.strip() == "1"

# def set_ping_allowed(enable):
#     flag = ('0','1')[enable]
#     cmd = f"/c powershell \"Set-ItemProperty -Path 'HKLM:\\SOFTWARE\\Policies\\Microsoft\\WindowsFirewall\\StandardProfile\\IcmpSettings' -name 'AllowInboundEchoRequest' -Value{flag}\""
#     _ = ShellExecuteEx(lpVerb='runas', lpFile='cmd.exe', lpParameters=cmd)

def set_ping_allowed(enable):
    flag = ('0','1')[enable]
    cmd = "/c " + " && ".join(f"reg add HKEY_LOCAL_MACHINE\\SOFTWARE\\Policies\\Microsoft\\WindowsFirewall\\StandardProfile\\IcmpSettings /v {entryname} /t REG_DWORD /d {flag} /f" for entryname in ICMP)
    _ = ShellExecuteEx(lpVerb='runas', lpFile='cmd.exe', lpParameters=cmd)

def get_hibarnate():
    process = run(["powershell", "-WindowStyle", "hidden", "Get-ItemProperty -Path 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Power' -name HibernateEnabled  | select -ExpandProperty HibernateEnabled"], stdout=PIPE, text=True, startupinfo=STARTUP_INFO)
    return process.stdout.strip() == "1"

def set_hibarnate(enable):
    flag = ('0','1')[enable]
    cmd = f"/c reg add HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Control\\Power /v HibernateEnabled /t REG_DWORD /d {flag} /f && reg add HKEY_LOCAL_MACHINE\\Software\\Policies\\Microsoft\\Windows\\Explorer /v ShowHibernateOption /t REG_DWORD /d {flag} /f"
    _ = ShellExecuteEx(lpVerb='runas', lpFile='cmd.exe', lpParameters=cmd)

def be_free(icon,item):
    service_set_status('pangps',False)
    service_set_status('stAgentSvc',False)
    service_set_status('webclient',True)
    reflesh_status()
    icon.notify('You\'re free.')

def be_corporate(icon,item):
    service_set_status('webclient',False)
    service_set_status('stAgentSvc',True)
    service_set_status('pangps',True)
    flush_dns(get_first_interface_conencted())
    set_firewall_status(True)
    reflesh_status()
    icon.notify('Network complient.')

def victor_the_cleaner(icon,item):
    service_set_status('webclient',False)
    service_set_status('stAgentSvc',True)
    service_set_status('pangps',True)
    flush_dns(get_first_interface_conencted())
    set_firewall_status(True)
    set_rights(False)
    run_at_startup(False)
    icon.notify('Ass clean.')
    terminate(icon,item)

def abort_shutdown():
    run(['shutdown','-a'], startupinfo=STARTUP_INFO)

def terminate(icon, item):
    icon.stop()
    sys.exit(0)

def get_first_interface_conencted():
    process = run(["netsh", "interface", "ip", "show", "interfaces"], stdout=PIPE, text=True, startupinfo=STARTUP_INFO)
    return next(filter(lambda y: "connected" in y and not "Loopback" in y, process.stdout.rstrip().splitlines()), "connected").split("connected")[1].strip()

def is_dns(interface, dns_name):
    dns_ip = DNS_LIST.get(dns_name)[0]
    next_line_is_the_one = False
    process = run(["netsh", "interface", "ip", "show", "dnsservers"], stdout=PIPE, text=True, startupinfo=STARTUP_INFO)
    for line in process.stdout.rstrip().splitlines():
        if interface in line:
            next_line_is_the_one = True
            continue
        if next_line_is_the_one:
            if dns_ip in line:
                return True
            else:
                return False
    return False

def set_dns(icon, dns_name):
    dns_ips = DNS_LIST.get(dns_name)
    interface = get_first_interface_conencted()
    if is_dns(interface,dns_name):
        flush_dns(interface)
        icon.notify('DNS will now be provided by the Router.')
    else:
        cmd = 'netsh interface ip set dns name="'+interface+'" static '+dns_ips[0]+' & netsh interface ip add dns name="'+interface+'" '+dns_ips[1]+' index=2'
        _ = ShellExecuteEx(lpVerb='runas', lpFile='cmd.exe', lpParameters='/c '+cmd)

def flush_dns(interface):
    _ = ShellExecuteEx(lpVerb='runas', lpFile='cmd.exe', lpParameters='/c '+'netsh interface ip set dns name="'+interface+'" dhcp')

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
            'Firewall',
            lambda icon, item: set_firewall_status(not item.checked),
            checked=lambda item: get_firewall_status()),
        Menu.SEPARATOR,
        MenuItem(
            'Run at startup',
            lambda icon, item: run_at_startup(not item.checked),
            checked=lambda item: states['startup']),
        Menu.SEPARATOR,
        MenuItem(
            'Dns',
            Menu(
                *(MenuItem(
                    dns_name,
                    lambda icon, item: set_dns(icon, item.text),
                    checked=lambda item: is_dns(get_first_interface_conencted(),item.text),
                    radio=True)
                for dns_name in DNS_LIST.keys())
                )
            ),
        Menu.SEPARATOR,
        MenuItem(
            'Miscellaneous',
            Menu(
                MenuItem(
                    'Hibarnate Feature',
                    lambda icon, item: set_hibarnate(not item.checked) and icon.notify('Restart Needed.'),
                    checked=lambda item: get_hibarnate()),
                MenuItem(
                    'ICMP Features',
                    lambda icon, item: set_ping_allowed(not item.checked),
                    checked=lambda item: get_ping_allowed()))
            ),
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
