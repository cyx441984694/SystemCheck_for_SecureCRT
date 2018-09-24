# SystemCheck_for_SecureCRT

It contains three parts. <br>
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

HOW TO USE THIS PROJECT
----------------

'Download and unzip "SecureCRTPortable" from the Internet'<br>

How to use:<br>
1. Download <The path could not contain Chinese and blank space<br>
2. Download the proj. ect and locate it under the same location. (SessionList.txt is on the same path as the SecureCRTPortable.exe<br>
3. Run the "Xunjian..py" in Xunjian folder or run the shorcut Xunjian.py at the root directly.
4. FinalReport could be seen in the same file.<br>

![Example](https://github.com/cyx441984694/SystemCheck_for_SecureCRT/blob/master/system-check.png)
