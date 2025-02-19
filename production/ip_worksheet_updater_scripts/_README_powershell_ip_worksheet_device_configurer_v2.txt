1. Make sure you have installed the crestron.exe powershell edk (if you have run any of my previous scripts you should have already done this. 
	1.1 You must install the Import-Excel library
	1.2 Open powershell as administrator, then run command Install-Module Import-Excel -Clobber

2. Run Script

	2.1 The script will discover any crestron devices that are on the network. if there are no devices available it will tell you to fix this, then exit. 

	2.2 If there are devices on the network at all, it will ask you to select an IP worksheet from your computer. (the formatting of the IP worksheet IS IMPORTANT. it must follow the standard of the USNH IP worksheet.)

	2.3 The script will now run through all the options, so do as it says. Its pretty self-explanatory
4. Complete
	4.2 thats about it. just wait for the script to progress through all the devices, then it will print the text "Press Enter to Exit" which means the script has attempted to do all the things you asked of it, and is all done. 