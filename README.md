	1.	Installing Python: On Linux Distributions, MAC OS X, and Unix machines; Python is by default installed. However, on Windows machines, it needs to be installed separately. Python installers for different OS can be downloaded from this link: https://www.python.org/downloads/

	2.	First go to the directory where you’ve installed Python. For eg. I have the latest Python version 3.6, and its location is <C:\python36> Use the <pip> tool to install required Packages -
  Go to command Prompt and type following in <C:\python36>
  >>pip install selenium
  >>pip install xlutils
  >>pip install xlrd

	3.	Check Chrome Browser Version and Download appropriate ChromeDriver from https://chromedriver.chromium.org/downloads

	4.	Save Pyhton Script named ‘Airtable_Automation.py’ anywhere on your computer
	5.	Open ‘Airtable_Automation.py’ as text and change location of following
  Work_Order_Path = ("Desktop/Airtable_workorder.xlsx") [create a new excel sheet where you will paste work order numbers in first column of sheet]
 Save_Path = ('Desktop/AirTable_0123.xls') this is the “output excel sheet”. Just Provide path where you want output file to be stored
 Webdriver_Path = ('Downloads/chromedriver') from Step 5 give path of downloaded chromedriver

	6.	Copy work order numbers from Airtable and paste them in first column of a new excel sheet that you created e.g Airtable_workorder.xlsx

	7.	In command Prompt go to the directory where you have stored Airtable_Automation.py and type
  
  >> python Airtable_Automation.py


