#$language="python"
#$interface="1.0"

'''
This script can accomplish the following
1. SECURE CRT, AUTO OPEN TAB, AUTO TELNET
You need to first edit the "SessionList.txt", add the ip address which you want to telnet to.

2. Each tab have a log file(Use the logging save option of the SecureCRT)

3. Send commands to the Tabs.
    "show clock",
    "show version",
    "show system",
    "show cpuinfo",
    "show meminfo",
    "show ip route",
    "show envm",
    "show log",
    "show ha log",
    "show ip ospf neighbor",
    "show ip bgp summary",
    "show isis neighbor",
    "show cable modem summary total",
    "show tech",
    "show version"
'''

import re
import os
import subprocess

SCRIPT_TAB=crt.GetScriptTab()
SCRIPT_TAB.Screen.Synchronous=True
SCRIPT_TAB.Screen.IgnoreEscape = True

LOG_DIRECTORY = os.path.join(
	os.getcwd(), 'Log')   ##The log folder

COMMANDS = [
    "show clock",
    "show version",
    "show system",
    "show cpuinfo",
    "show meminfo",
    "show ip route",
    "show envm",
    "show log",
    "show ha log",
    "show ip ospf neighbor",
    "show ip bgp summary",
    "show isis neighbor",
    "show cable modem summary total",
    "show tech",
    "show version"
	]


def main():

    if not os.path.exists(LOG_DIRECTORY):
        os.mkdir(LOG_DIRECTORY)

    if not os.path.isdir(LOG_DIRECTORY):
        crt.Dialog.MessageBox(
            "Log output directory %r is not a directory" % LOG_DIRECTORY)
        return

    ##You can delete the bellowing def and disable the definition of the below if you already have tabs connected.
    AutoConnectTab(os.getcwd()+"\SessionList.txt")


    ###Confirm if it is safe to input commands
    while True:
        if not SCRIPT_TAB.Screen.WaitForCursor(1):
            break

    #####Get the tab, and send commands to the tabs
    skippedTabs=""
    Count=crt.GetTabCount()

    for i in range(1, Count+1):
        tab = crt.GetTab(i)
        tab.Activate()
        ##Get the hostname
        strCmd = "show run | in ^hostname\r"
        tab.Screen.Send(strCmd)
        tab.Screen.WaitForStrings([strCmd + "\r", strCmd + "\n"])
        host = tab.Screen.ReadString(["\r", "\n"])
        if "" in host:
            Host = host.split()[1]
            hostname=Host[1:-1]
        else:
            Host = "unknown-hostname"
            hostname=Host

        ##Identify the filename and enable the logging
        tab.Session.LogFileName = os.path.join(LOG_DIRECTORY, hostname + '.txt')

        if tab.Session.Connected==True:
            tab.Session.Log(True)   ##Save the session log
            tab.Screen.Send("page-off\r")

        ###Send the commands to the tab
            for command in COMMANDS:
                try:
                    tab.Screen.Send(command+"\r")
                    tab.Screen.WaitForString(hostname+"#")
                except:
                    return
        else:  ##Check if there is any skippedTabs
            if skippedTabs=="":
                skippedTabs=str(i)
            else:
                skippedTabs=skippedTabs+","+str(i)

	##Detect if it is time to disconnect the tab session. If the system detect the "Show version" in the end. It will disconnect the tab
        response=tab.Screen.WaitForStrings("show version")
        if response:
            tab.Session.Log(False)
            tab.Session.Disconnect()

    if skippedTabs!="":
        skippedTabs="\n\nThe following tabs did not receive the command:"+skippedTabs
        crt.Dialog.MessageBox("We have skippedTabs:" + str(skippedTabs))

    if crt.Dialog.MessageBox( "All  inspection CLI commands are successfully executed  for all the %s sessions."%(str(Count)), "Please confirm:", 48|0):
        crt.Quit()

    LaunchViewer(LOG_DIRECTORY)


def AutoConnectTab(file):  ###Auto connect sessions from a txt
    ####This is to Open the "SessionList.txt" and get a list of the ip address
    '''If you want to ssh to the ip address, Use the following. Before that you need to assign the pwd,user,host,port, e.g:
            port=xx
            user=xx
            cmd = "/SSH2 /PASSWORD %s %s@%s /P %s" %(password,user,host,port)   ##equals to "ssh hhr@50.206.125.254 -p 99 -PASSWORD xxx
            crt.Session.Connect(cmd)'''
	
	##
    if not os.path.exists(file):
        return
    sessionFile = open(file, "r")
    sessionArray = []
    for line in sessionFile:
        session = line.strip()
        if session:
            sessionArray.append(session)
    sessionFile.close()

    # Receive variable: user, password
    objNewTab=crt.Session.ConnectInTab("/TELNET %s 23" % sessionArray[0])
    user = crt.Dialog.Prompt("Enter user name:", "Login", "", False)
    password = crt.Dialog.Prompt("Enter password:", "Login", "", True)
    login(user,password)
	###If the password is not correct. Pop up a window for relogin
    while not objNewTab.Screen.WaitForString(">", 3):
        user=crt.Dialog.Prompt("Relogin! Enter user name:", "Login", "", False)
        password=crt.Dialog.Prompt("Relogin! Enter password:", "Login", "", True)
        login(user,password)
	##If the password is correct, enter the enable mode.
    enablemode()
	##For the remaining windows, use the login information that you enter in the first tab.
    for session in sessionArray[1:]:
        try:
            crt.Session.ConnectInTab("/TELNET %s 23" % session)
            login(user,password)
            enablemode()
        except IOError:
            pass
        if not tab.Session.Connected:
            return

def login(user,password):
    objNew=crt.GetActiveTab()
    objNew.Screen.Synchronous = True
    objNew.Screen.IgnoreEscape = True
    objNew.Screen.WaitForString("login:")
    objNew.Screen.Send(user + "\r")
    objNew.Screen.WaitForString("Password:")
    objNew.Screen.Send(password + "\r")

def enablemode():
    objNew=crt.GetActiveTab()
    objNew.Screen.Send("enable\r")
    objNew.Screen.WaitForString("Password:")
    objNew.Screen.Send("123\r")
    objNew.Screen.WaitForString("#")

##For call out the folder
def LaunchViewer(filename):
    try:
        os.startfile(filename)
    except AttributeError:
        subprocess.call(['open', filename])



main()
