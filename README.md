# SystemCheck_for_SecureCRT

It contains two parts. <br>
1. Script running in the SecureCRT generates log files for each host. The logs are under Log folder<br>
2. Filter some key words from the log file.<br>
3. Merge the csv to one xls and add some formats.<br>

Xunjian-crt.py
--------------

'This nees to run in the SecureCRT.'<br>

How to use:<br>
1. Edit a file "SessionList.txt" under the same folder.<br>
e.g:<br>
SessionList.txt<br>
192.168.11.161<br>
192.168.11.162<br>
2. The script will help to achieve the following tasks.<br>
  a. Auto connect to the ip addresses listed in the text<br>
	b. Each host has a separate log file<br>
	c. Send commands to each host.<br>

Xunjian.py
-----

The py is under the Xunjian folder. <br> 
It will first create a bat and run the bat to open a SecureCRT for running the "Xunjian-crt.py" automaticall.<br> 
Then it will generate several csv.<br> 
It will generate a final xlsx under the folder.<br> 

Xunjian.rar
----------------

'It contains the Xunjian.exe and the SecureCRT.exe.'<br>

How to use:<br>
1. Download the rar<The path could not contain Chinese and blank space<br>
2. Run the Xunjian.exe. It is a shortcut and equal to the Xunjian.py mentioned above.<br>
3. FinalReport could be seen in the same file.<br>
