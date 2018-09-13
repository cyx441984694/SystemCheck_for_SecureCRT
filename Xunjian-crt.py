#$language="python"
#$interface="1.0"

'''
##Please note that the password of the enable mode is by default.
##Please change it if the
We set up two files in the same path.
You can add or remove any paragraph in the list. You can also add "#" in it. We will ignore the comment in the file.
1. CommandList.txt
2. SessionList.txt
The py will help to accomplish the following:
1. For each host, we have a separate log file named as the hostname under the "Log" folder.
2. Open a new tab and auto telnet the ip addresses which are in the "SessionList.txt".
3. Send commands to each tabs.

This script merges some code from the example logged on the vanskype forum.
If needed, it can also achieve a functionality to log each command to one single log and complement the error information.
'''
import time
import re
import os
import subprocess

##Need to change the following if you want to name the txt by yourself.
g_strCommandsFile="CommandList.txt"
g_strHostsFile="SessionList.txt"
g_strComment = "#"


SCRIPT_TAB= crt.GetScriptTab()
SCRIPT_TAB.Screen.Synchronous=True
SCRIPT_TAB.Screen.IgnoreEscape = True

##The log folder
LOG_DIRECTORY = os.path.join(
	os.getcwd(), 'Log')


def main():

    if not os.path.exists(LOG_DIRECTORY):
        os.mkdir(LOG_DIRECTORY)

    if not os.path.isdir(LOG_DIRECTORY):
        crt.Dialog.MessageBox(
            "Log output directory %r is not a directory" % LOG_DIRECTORY)
        return

    ##Read in the hosts file and populate an array with non-comment lines
    vReturnValsHostsFile = ReadDataFromFile(g_strHostsFile, g_strComment)
    if not vReturnValsHostsFile[0]:
        DisplayMessage("No hosts were found in file: \r\n    " + g_strHostsFile)
        return
    g_vHosts = vReturnValsHostsFile[1]
    g_nHostCount = vReturnValsHostsFile[2]
    crt.Dialog.MessageBox("Read in {0} hosts".format(g_nHostCount))

    # If the g_strCommandsFile path exists, load those commands into a
    # global array and populate global variables for use when a host needs
    # to send commands from this file.
    # Call the ReadDataFromFile() function to populate the array of commands
    # that will be sent for this host.
    if os.path.isfile(g_strCommandsFile):
        vReturnValsCmdsFile = ReadDataFromFile(g_strCommandsFile, g_strComment)
        if not vReturnValsCmdsFile[0]:
            strError = (
                "Error attempting to read host-specific file for host '")
            DisplayMessage(strError)
        else:
            g_vCommands = vReturnValsCmdsFile[1]
            g_vCommandCount = vReturnValsCmdsFile[2]
    else:
        strError=("This is not a file!")
        DisplayMessage(strError)

    #Check if vReturnValsCmdsFile is True/False. Should be always True.
    if not vReturnValsCmdsFile[0]:
        strError = (
            "Error attempting to read host-specific file for host '")
        DisplayMessage(strError)

    ##This is the setting the auto option when you need to ssh into the systems.
    #AuthPrompt()

    ##You can delete the bellowing def and disable the definition of the below if you already have tabs connected.
    AutoConnectTab(g_strHostsFile,g_vHosts)

    ###Confirm if it is safe to input commands
    while True:
        if not SCRIPT_TAB.Screen.WaitForCursor(1):
            break

    #####Get the tab, and send commands to the tabs
    skippedTabs = ""
    Count=crt.GetTabCount()


    for i in range(1, Count+1):
        tab = crt.GetTab(i)
        tab.Activate()

        tab.Screen.Synchronous = True
        tab.Screen.IgnoreEscape = True

        ##Acquire the hostname
        strCmd = "show run | in ^hostname\r"
        tab.Screen.Send(strCmd)
        tab.Screen.WaitForStrings([strCmd + "\r", strCmd + "\n"])
        host = tab.Screen.ReadString(["\r", "\n"])
        if "" in host:
            Host = host.split()[1]
            hostname=Host[1:-1]
        else:
            Host = "UnknownHost"
            hostname=Host

        ##Identify the filename and enable the logging
        ##The log will be named as "hostname".txt
        tab.Session.LogFileName = os.path.join(LOG_DIRECTORY, hostname + '.txt')

        if tab.Session.Connected==True:
            ##Save the session log
            tab.Session.Log(True)
            tab.Screen.Send("page-off\r")

            # Send each command one-by-one to the remote system:
            for strCommand in g_vCommands:
                if strCommand =="":
                    break
                # Send the command text to the remote
                # crt.Dialog.MessageBox("About to send cmd: {0}".format(str(strCommand)))
                try:
                    tab.Screen.Send("{0}\r".format(strCommand))
                    tab.Screen.WaitForString(hostname+"#")
                except:
                    return

        else:  ##Check if there is any skippedTabs
            if skippedTabs=="":
                skippedTabs=str(i)
            else:
                skippedTabs=skippedTabs+","+str(i)

        #Disconnect the crt when detecting the "show version" in the end
        tab.Screen.Send("show version\r")
        response=tab.Screen.WaitForStrings("show version")
        if response:
            tab.Session.Log(False)
            tab.Session.Disconnect()


    if skippedTabs!="":
        skippedTabs="\n\nThe following tabs did not receive the command:"+skippedTabs
        crt.Dialog.MessageBox("We have skippedTabs:" + str(skippedTabs))

    ##Pop up a messageBox in the end, when users click the "Confirm" button, Exit the crt
    if crt.Dialog.MessageBox("All inspection CLI commands are successfully executed for all the %s sessions with %s commands."%(str(Count),g_vCommandCount), "Please confirm:", 48|0):
        crt.Quit()

    ##Pop up the Log window
    LaunchViewer(LOG_DIRECTORY)


def AutoConnectTab(file,g_vHosts):  ###Auto connect sessions from a txt
    ####This is to Open the "SessionList.txt" and get a list of the ip address
    '''If you want to ssh to the ip address, Use the following. Before that you need to assign the pwd,user,host,port, e.g:
            port=xx
            user=xx
            cmd = "/SSH2 /PASSWORD %s %s@%s /P %s" %(password,user,host,port)   ##equals to "ssh hhr@50.206.125.254 -p 99 -PASSWORD xxx
            crt.Session.Connect(cmd)'''
    if not os.path.exists(file):
        return
    if not os.path.isfile(file):
        return

    # Receive variable
    objNewTab=crt.Session.ConnectInTab("/TELNET %s 23" % g_vHosts[0])
    user = crt.Dialog.Prompt("Enter user name:", "Login", "", False)
    password = crt.Dialog.Prompt("Enter password:", "Login", "", True)
    login(user,password)
    ##If the login is incorrect. Pop up thte window again
    while not objNewTab.Screen.WaitForString(">", 3):
        user=crt.Dialog.Prompt("Relogin! Enter user name:", "Login", "", False)
        password=crt.Dialog.Prompt("Relogin! Enter password:", "Login", "", True)
        login(user,password)
    #If the login is correct, enter the enablemode
    enablemode()

    ##Pass the login information to the remaining tabs.
    for session in g_vHosts[1:]:
        try:
            crt.Session.ConnectInTab("/TELNET %s 23" % session)
            login(user,password)
            enablemode()
        except IOError:
            pass
        if not SCRIPT_TAB.Session.Connected:
            return

def ReadDataFromFile(strFile,strComment):
    #        strFile: IN  parameter specifying full path to data file.
    #     strComment: IN  parameter specifying string that preceded
    #                    by 0 or more space characters will indicate
    #                    that the line should be ignored.
    # Return value:
    #   Returns an array where the elements of the array have the following
    #   meanings.
    #     vParams[0]: OUT parameter indicating success/failure of this
    #                 function
    #     vParams[1]: OUT parameter (destructive) containing array
    #                 of lines read in from file.
    #     vParams[2]: OUT parameter integer indicating number of lines read
    #                 in from file.
    #     vParams[3]: OUT parameter indicating number of comment/blank lines
    #                 found in the file
    global g_strError
    vLines=[]
    nLineCount=0
    nCommentLines=0
    vOutParams=[False,vLines,nLineCount,nCommentLines]

    #Check to see if the file exists...if not, bail early.
    if not os.path.exists(strFile):
        DisplayMessage("File not found:{}".format(strFile))
        g_strError="File not found:{}".format(strFile)
        return vOutParams

    ##Used to detect comment lines
    ##(.*?#.*?)
    p=re.compile("(^[ \\t]*(?:{0})+.*$)|(^[ \\t]+$)|(^$)".format(strComment))
    try:
        objFile=open(strFile,'r')
    except Exception as objInst:
        g_strError="Unable to open '{0}' for reading: {1}".format(strFile,str(objInst))
        return vOutParams

    ##Read to see if the encoding is right or not
    try:
        b=str(objFile.read(1))
    except Exception as objInst:
        g_strError = "Unable to open '{0}' for reading: {1}".format(strFile, str(objInst))
        return vOutParams

    ##Check if it is empty
    if len(b) < 1:
        g_strError="File is empty:{0}".format(strFile)
        return vOutParams

    if ord(b) == 239:
        objFile.close()
        strMsg="UTF-8 format is not supported. File must be saved in ANSI format:\r\n{0}".format(strFile)
        DisplayMessage(strMsg)
        g_strError=strMsg
        return vOutParams
    elif ord(b)== 255 or ord(b) ==254:
        objFile.close()
        strMsg="Unicode format is not supported. File must be saved in ANSI format:\r\n{0}".format(strFile)
        DisplayMessage(strMsg)
        g_strError=strMsg
        return vOutParams
    else:
        #Close and re-open so that we don't lose the first byte
        objFile.close()
        objFile=open(strFile,'r')

    for strLine in objFile:
        strLine=strLine.strip(' \r\n')
        #Look for comment line
        if p.match(strLine):
            nCommentLines+=1
        else:
            vLines.append(strLine)
            nLineCount+=1

    if nLineCount<1:
        vOutParams=[False,vLines,nLineCount,nCommentLines]
        g_strError="No valid lines foud in file:{0}".format(strFile)
        return vOutParams

    vOutParams=[True,vLines,nLineCount,nCommentLines]
    return vOutParams


def DisplayMessage(strText):
    crt.Dialog.MessageBox(strText)

def AuthPrompt():
# Before attempting any connections, ensure that the "Auth Prompts In
# Window" option in the Default session is already enabled.  If not, prompt
# the user to have the script enable it automatically before continuing.
# Before continuing on with the script, ensure that the session option
# for handling authentication within the terminal window is enabled
    objConfig=crt.OpenSessionConfiguration("Default")
    bAuthInTerminal=objConfig.GetOption("Auth Prompt In Window")
    if not bAuthInTerminal:
        strMessage = ("" +
            "The 'Default' session (used for all ad hoc " +
            "connections) does not have the 'Display logon prompts in " +
            "terminal window' option enabled, which is required for this " +
            "script to operate successfully.\r\n\r\n")
        if not PromptYesNo(
            strMessage+
            "Would you like to have this script automatically enable this " +
            "option in the 'Default' session so that next time you run " +
            "this script, the option will already be enabled?"):
            return
        #Use answered prompt with Yes, so the option is saved
        objConfig.SetOption("Auth Prompts In Window", True)
        objConfig.Save()

def PromptYesNo(strText):
    vbYesNo = 4
    return crt.Dialog.MessageBox(strText, "SecureCRT", vbYesNo)

def login(user,password):
    objNew = crt.GetActiveTab()
    objNew.Screen.Synchronous = True
    objNew.Screen.IgnoreEscape = True
    objNew.Screen.WaitForString("login:")
    objNew.Screen.Send(user + "\r")
    objNew.Screen.WaitForString("Password:")
    objNew.Screen.Send(password + "\r")

def enablemode():
    objNew = crt.GetActiveTab()
    objNew.Screen.Send("enable\r")
    objNew.Screen.WaitForString("Password:")
    objNew.Screen.Send("casa\r")
    objNew.Screen.WaitForString("#")

def LaunchViewer(filename):
    try:
        os.startfile(filename)
    except AttributeError:
        subprocess.call(['open', filename])


main()