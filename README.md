# clean_db

INFOS

	- Langage : Python
	- Version : 3.7.3
	- Modules : PySide2, gspread, autoPEP8, oauth2client

BEFORE STARTING

	- Here is my first script posted on github. 
	- This project was built in my very first month of learning code. 
	- Will come better version of this script on june. 	
	- Feel free to comment.

WHAT THE SCRIPT DOES

	- Directly from the script interface, you can clean your googlesheet’s databases. 
	- The script completly ignore Camel Case thanks to RegExp.
 
SETUP
  1. CREATE YOUR ENV

    Follow the steps 1,2 and 3 from this link :
    https://www.liquidweb.com/kb/how-to-setup-a-python-virtual-environment-on-ubuntu-18-04/

  2. INSTALL PIP MODULE FROM REQUIREMENT.TXT

	pip install -r requirements.txt

  3. CONNECT WITH GOOGLE API

    Follow  « Google Drive API and Service Accounts » indications from this link :
    https://www.twilio.com/blog/2017/02/an-easy-way-to-read-and-write-to-a-google-spreadsheet-in-python.html

    - Don’t forget to put your own client_secret.json (see link) in the main folder,
    - Don’t forget to share your googlesheet (see link).

  4. HOW TO USE THE SCRIPT

    STEP 1
        - Enter the name of the google file,
        - Select your sheet,
        - Select your column,
    STEP 2
        - Create a list,
        - Enter keyWord,
        - Clean !

PROBLEMS IDENTIFIED

	- The script is slow (due to unoptimized construction),
	- API quota exceeded (too many request due to unoptimized construction),
	- Will come better version of this script on june.
