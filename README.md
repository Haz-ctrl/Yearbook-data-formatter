# Yearbook data formatter

This is documentation on how to use the Yearbook Data Formatter. This application is compatible with Windows Operating Systems ONLY. 
Follow the steps carefully to ensure that no problems occur during execution of the app. 

Other than this .txt file, there is an application file (YearbookData.exe) and a src folder. You don't need to worry about the src folder, as that just contains the files and scripts necessary to execute the application. 


Before running the application file, ensure you check these requirements:

1) You have the spreadsheet file you need to convert. Make sure this is in .xlsx format. To download a copy of a Google Sheet as an xlsx file, open up the sheet in Google Sheets, then click the 'File' tab. Inside the 'File' tab, click on 'Download' -> 'Microsoft Excel (.xlsx)'. A download of the spreadsheet should appear in your 'Downloads' folder on your machine. Be wary that the app will only let you pick files with an extension of  '.xlsx', so make sure your spreadsheet is like this.

2) Once you have a copy of the spreadsheet on the machine as an xlsx file, open it up in Microsoft Excel just to make sure there're no formatting errors. Normally, minor things such as fonts will change, so don't be too worried about this.

3) You also want to check that your data includes the main things that'll go into each student's yearbook profile. Examples of these include: Name, Nickname, Year joined RGC, Memorable moments, Best quotes, Friends, Embarrassing moments, General paragraph etc. The program still functions even if there isn't much data, but you don't want lacking profiles. 


Assuming you meet all these, then you can run the app. Click on the application file in the folder (YearbookData.exe), and wait a few seconds for the app to load. After the application loads, you should see a screen with 3 buttons. 

1) Hit the 'Open File' button first. This will allow you to select the spreadsheet file you've downloaded. The window will only display xlsx files. Once you've selected your file, either double-click on it or hit 'Open' in the bottom right of the window. You should now see a file path (like 'something:/something/something/something...') appearing in the white box above the buttons. If you want to switch to a different file, hit the 'Open File' button again and select the file you want. 

2) After you can see the file path in the window, hit the 'Convert Data' button. This will convert the file from a '.xlsx' file to a '.docx' file. Once the conversion has finished, a window will pop up asking you for a file name and where you want to save your converted data. Once you've selected a valid destination and typed in the name of your file, hit 'Save'.

3) You will have found that the '.docx' file appears in the destination and with the name you specified. You can now view it online in Google Docs, or in Microsoft Word. If you make any changes to the spreadsheet data, you'll need to run the application again to generate a new docx file. This should now be able to work fine with Adobe InDesign.


I hope this guide is helpful in your use of the Yearbook Data Formatter application, and this software is covered under an MIT license, so you're allowed to modify and reproduce the software in whatever way you see fit. If you want to take a look at the code, open up the '.py' script in the 'src' folder and start editing if you can find a way to optimise things. 
