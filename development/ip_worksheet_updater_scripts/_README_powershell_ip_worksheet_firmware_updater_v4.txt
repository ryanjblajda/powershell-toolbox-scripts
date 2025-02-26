1. Make sure you have installed the crestron.exe powershell edk (if you have run any of my previous scripts you should have already done this. 
	1.1 You must install the Import-Excel library
	1.2 Open powershell as administrator, then run command Install-Module Import-Excel -Clobber

2. Run Script

	2.1 The script will discover any crestron devices that are on the network. if there are no devices available it will tell you to fix this, then exit. 

	2.2 If there are devices on the network at all, it will ask you to select an IP worksheet from your computer. (the formatting of the IP worksheet IS IMPORTANT. it must follow the standard of the USNH IP worksheet.)

	2.3 The script will ask you which sheet within the workbook you want to import, this should be pretty self explanatory, type in the worksheet that has the devices you want to target in it. (this is CASE specific)

	2.4 The script will now ask you the Building & Rooms you wish to target, and then have you confirm your entry afterwards. This will allow you to target multiple buildings and rooms at the same time. To target all rooms in a building you SHOULD be able to target all, by typing in "all" instead of a room number. [i have to go back and test but this should work fine] 

	2.5 The script will then ask you what model of device you want to target with your firmware update, and then confirm you entered it correctly and allow you to re-do it if necessary.
	
	BE SPECIFIC. more specific is better than less specific. in theory the devices shouldnt load firmware they are incompatible with, but lets not brick anything. 

	for example, you want to update only DMPS3-4K-xxx-c devices on your network, make sure you enter DMPS3-4K (not case senstive) and not just DMPS3. if you just typed in DMPS3, you could end up sending firmware to DMPS3-300-C's or other 3 series DMPS units that arent 4K. this applies to other types of devices, i.e. XX60 vs XX50 or XX52 touchpanels. hopefully you get where im going with this. 

	this script is powerful, but with great power comes great responsibility, we dont want to accidentally brick 100 processors at once by sending the wrong firmware, as mentioned previously being more specific is better than being vague. 

	2.6 the script will now ask you what the update command is for the devices you are sending firmware to. by this point you have already downloaded the firmware file, so if your file is a .zip, then you are MOST LIKELY sending firmware to a DM/DMPS so the update command is PUSHUPDATE FULL (not case sensitive. this FULL flag ensures we update all the endpoints, scalers, transmitters, etc).

	if the file you are going to send is a .puf file, then the command is simply PUF (once again not case sensitive)

	2.7 the script will now show you every devices it plans to update. MAKE SURE ITS RIGHT. im not perfect, and the IP worksheet might not be perfect, so read through the list and make sure everything seems right. you still have the ability to kill the script at any point while its running, and there is still one step before the updates begin. but its good to double check at this point. 

	2.8 lastly, the script will ask you to select the firmware file. after you select it and hit okay, the updates will begin automatically. the script will post a message saying what device it is updating at what IP and the filename of the firmware its sending, so if anything jumps out and seems wrong, just close the script window. 

3. Wait Patiently
	
	3.1 the FTP process takes a long ass time, and there is no way to measure progress (at least not that i have implemented as of yet)
		3.1.1 you can use filezilla to check the FIRMWARE directory and refresh every few seconds to see the firmware file increase in size, to verify things are working. 

	3.2 if the above seems to be happening, do some other stuff, cause this will take a bit. once the firmware is uploaded the script will report the response of the unit until it reboots. 
		3.2.1 if its a dm device youll see a bunch of info just like in the web browser with some progress 1%2%4% etc. 
		3.2.2 if its a PUF device most likely youll just see "Inflating PUF" and then it will close the connection
4. Complete
	4.2 thats about it. just wait for the script to progress through all the devices, then it will print the text "Press Enter to Exit" which means the script has attempted to do all the things you asked of it, and is all done. 