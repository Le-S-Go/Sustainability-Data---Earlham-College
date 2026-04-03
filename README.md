# Sustainability-Data---Earlham-College
This repository contains Python scripts to automate the processing of Earlham College's natural gas, electricity, and water bills. All three scripts work in much the same way: they scrape each utility bill to extract the meter number, usage, and cost for each meter, then update those values in a spreadsheet. 
## Using the scripts
1. Create a folder for the month you are updating the information about this utility
2. Go to the appropriate website (Centerpoint for natural gas, Richmond Power and Light for electricity, Indiana American Water for water) and download all of the utility bills for the appropriate month into the folder
3. Go to the command line of an environment with Python3 installed and call the script with the following parameters

-m : the month you are updating

-f : the file path to the folder containing all bills for this utility for this month

-e : the file path to a blank copy of the master spreadsheet

For example, if you wanted to update the electricity data for January and stored the bills in a folder titled "ElectricityBills_January", you would call
```
$ python3 electricityScan.py -m 'Jan.' -f 'ElectricityBills_January' -e 'Blank Sheet.xlsx'
```
This will create a new, updated spreadsheet that you can then copy the relevant columns of information from into the master spreadsheet. 
