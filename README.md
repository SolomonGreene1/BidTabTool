# BidTabTool
JWI code for bid tabulations


# JWI Bid Tabulation Tool

	The Bid Tabulation Tool is a graphical user interface (GUI) application designed to assist with you with getting up to date information on bid tabulations. It provides a user-friendly interface for entering and managing bid information, generating tabulation reports, and more. This tool only tabulates Tollway pay items. For a different tool that tabulates tollway as well as IDOT pay items follow this link: https://flexureflow.com/biditem/search/idot

## MOST IMPORTANT NOTES, READ ME!

	DO NOT EDIT THE Source Data DO NOT EDIT folder. It contains all of the files used to present data and is automatically updated by the program with the check for new bid tabs button. 
	
	While downloading new data, DO NOT CLOSE THE PROGRAM EVEN IF IT SAYS IT IS NOT RESPONDING, downloading data may take some time. Please be patient.
	
	UPON YOUR FIRST SEARCH, all the bid tab data will need to be downloaded from oneDrive. This may take a while. DO NOT CLOSE THE PROGRAM. Once all files are downloaded onto your computer the program should run smoothly.

## Usage

	Open the Bid Tabulations folder in the JWI "Dept 20 Projects - Documents" OneDrive and then open the "Bid Tab Tool.exe" file. It should open in a few seconds.

	When the file opens you will be presented with a few options for how you would like to utilize the tool. You can enter pay-item numbers in three ways:
		
		1. Single item search: Type in a pay item number (such as jt280530) to the search bar and press search. Entering pay items is not case sensitive. The screen will begin to display relevant information from all instances of this pay item in Tollway contracts dating back to 2010.
		2. Bulk processing search: press the Choose File button to select an excel file with a list of pay item names. Your list must be formatted with item codes all in column A. The tool will run all of the items and automatically save them as a .csv and a .txt file.
		3. Search by name: Click on the Search for pay item by name checkbox and type in a pay item name or a word or phrase from a pay item name. Anything more than three letters will work. Pressing search will then display a window with all matching pay item codes and names where you can copy paste your desired item into the search bar.
		
	Hitting the search button or pressing the enter key will work to initiate a search. 
	
	Upon searching or pressing most buttons, green text will appear next to the search bar. This is the update label and will tell you what the program is currently doing. On rare occasions the program may run slowly. Please be patient and contact JSchroeder@jwincorporated.com or SGreene@jwincorporated.com if you believe there is an error. 
		
## Saving data and Location

	All data will be automatically saved as a .csv and .txt file upon search completion. csv files are formatted for data manipulation without headers, labels, and empty lines and txt files are formatted as you see the data presented on screen. 

	All files are located in BTT Outputs folder under a subfolder with today's date YYYY-MM-DD_Bid_Tabulations. All files are saved named by their pay-item code followed by their most abbreviated pay-item name. Note: this name may not match every instance of a pay-item. If a certain file has already been created it will not be created again that day.

## Check for new bid tabs button

	The check for new bid tabs button allows you to make sure that the data you are gathering is the most up to date data from the tollway's website. When pressed this button automatically compares the downloaded files in the source data folder to all available files on the tollway website: https://www.illinoistollwaybidding.com/jobs/678/specs/bid-tabulations. 

	If the downloaded files are up to date, a message will appear informing you that no new bid tabs are available. If there are files missing, a message will appear with the 4 digit code of the bid tab(s) and the option to download the file(s). If you press "Ok" the file(s) will be downloaded and saved in the source data folder under their corresponding year subfolder. Downloading files may take some time, DO NOT CLOSE THE PROGRAM EVEN IF IT SAYS IT IS NOT RESPONDING. Progress should be printed to the screen for downloads of more than one file. Note: Once one person downloads a new file, the file will be available for everyone. The program will know that this file has been downloaded for every user.

## Cancel Search

	The cancel search button is a useful feature allowing you to immediately stop a single or bulk item search. Upon pressing it, any saved data will be deleted unless it has already been searched and the search terminated. This button will save the search bar text of a single item search for you to edit if you may have mistyped it.

## Auto Scroll

	Auto scroll is automatically enabled upon opening the tool. To turn it off you may uncheck the Auto Scroll box or simply left click in the text box during a search. 

## Clear and Close buttons

	The Clear button clears the current screen and the search bar. The Close button closes the program. The program will not close when in the middle of a search but will close immediately after the current search is completed. All data will still be saved.

## Additional Features

	The program will auto scale to your monitor by enlarging with the Maximize button (second from top right). 
	
	Inside the search bar and text box you can use copy paste by selecting and right clicking or through the Command + C and Command + V keybindings. 
	
	The program will initially search for the pay-items name to display for you. Some pay-items have data but may not have a name and other items may have a name but no data. If the program finds no data, no csv or txt will be created.
	
	During searching or other operations, some buttons may not be functional. This is by design and should help guide you around the tool.
