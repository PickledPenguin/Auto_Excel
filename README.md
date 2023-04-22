# Description

This script extracts data from the Excel files in the "Excel" directory and outputs the gathered data for each Excel file according to a configuration set by the user via simulating mouse movement and keypresses via the python package pynput. 


# Fun Background Info

This is a script based on another script I created for the Rochester Institute of Technology's Software Engineering department to help automate the process of inputting the student employees' hours from Excel spreadsheets into an external website. Previously, this had to be done manually, cell-by-cell, line-by-line, which was tedious and unnecessary. So I decided to build this script to fix that problem and allow the SEO employees to free up a significant portion of their time for less drudging, more pressing tasks.


# Architecture Details

This script does not determine where to input the Excel data from the external website itself, but rather relies on a manual setup process in which the user directs the program to the points of interest (waypoints) on the monitor. There are two main reasons for this approach: First, I was unable to access the external website while developing this script as it was restricted and only accessible by members of the SE department. Second, in the case that the SE department decides to use a different external website or the original external website is changed or updated, this script would still function, and it would only need to be reconfigured rather than redesigned. 
This flexibility comes with several drawbacks: The initial setup process is more complex than I would like and the script is dependent on the screen being completely static and consistent every single time the script is run. I believe these drawbacks are worth the flexibility, at least for such a small-scale project.


# Required Installations

1.) python3

2.) pip or pip3

3.) all project requirements in requirements.txt

Once you have pip or pip3, type the following bash command in the project directory to install all the project requirements:

"pip install -r requirements.txt" (for pip)

or 

"pip3 install -r requirements.txt" (for pip3)


# Configuring the Script

To configure the script, type the following bash command in the project directory:

"python3 ./setup.py"

You will be able to configure the script's wait time between actions, datetime format string, execution type, and waypoints.

*Setting wait time*: 

The wait time is the time (in seconds) between any meaningful action the script simulates. This gives the website time to catch up and helps make sure that no inputs or data is lost. If your connection is a little unstable, I would recommend increasing this wait time. The default wait time is 0.5 seconds.

*Setting datetime format string*: 

The datetime format string is a string that uses certain format codes as standard directives for specifying the format in which you want to represent datetime type data. A comprehensive list of all the format codes can be found at this link: https://strftime.org Using the format codes, you can "insert" parts of the datetime data into your desired format. For example the format string: "%H:%M %p" will format the data like this: "**Hour**:**Minute** **AM or PM** and the format string: "Student completed work on %m/%d/%Y, which was a %A" will format the data like this: "Student completed work on **month**/**day**/**year**, which was a **Day of the week**"

*Setting execution type*: 

This program supports 2 execution types: file-based execution and row-based execution. Here is how each works: 

File-based execution: The script will run the configured waypoints for each Excel file in the Excel directory.
This option is meant to be used for cases where your Excel files have only one subset of data contained within them. For example, if you were entering employee hours contained in dozens of Excel files, each containing a spreadsheet with the hours that an employee worked for the week, you would use file-based execution to execute the waypoints for each Excel file.

Row-based execution: The script will run the configured waypoints for each row in each sheet in each Excel file in the Excel directory
This option is meant to be used for cases where your Excel files have multiple subsets of data contained within them, organized by rows. For example, if you were entering employee hours contained in a single Excel file with each employee's hours listed on a single row, you would use row-based execution to execute the waypoints for each row in the Excel file. 
Row-based execution executes for every sheet in each Excel file. For example, if you have an Excel file containing employee hours organized by row, with different departments listed on different sheets, row-based execution will execute for each row in each sheet in the Excel file. 
Row-based execution also executes for every Excel file in the Excel directory. For example, if you have an Excel file containing employee hours organized by row, with different departments listed on different sheets, and an Excel file for every company location, row-based execution will execute for each row in each sheet, in each file.

*Setting header configuration*:

A lot of Excel files will have the first row as a header. This is something that you don't want the program to read or execute, especially for row-based execution. This setting allows you to set the program to either ignore or keep the first row of every sheet in every Excel file.

For row-based execution, if you have a header, you should set the header configuration to "ignore" so you don't input the header data in the first row. However, if you have no header and your files just contain data, you should set the header configuration to "keep" so you don't lose the data in the first row.

NOTE: If the header configuration is set to "ignore", it will not read in the first row at all, shifting all rows up by one. This means that any data in row 2 will now be stored in row 1, row 3 will be stored in row 2 etc.
Because of this, if you are using file-based execution, I would recommend setting the header configuration to "keep" so that the rows and columns the program reads in match exactly with your Excel file to avoid confusion when you manually input the sheet, row, and column of the section that your data is contained in.

*Setting up "waypoints"*: 

Waypoints are points of interest that you indicate on the screen for the program to either click, input data, or perform some other action. Currently, there are 5 different types of waypoints: Click, Double click, Tab, Enter, Paste, Insert, and Wait.
Each type of waypoint has a character (a key on a keyboard) associated with it. To create a waypoint, move your mouse to the point on the screen where you want the waypoint to be executed, then hit the key on your keyboard that matches the appropriate waypoint action. Some waypoints require additional data which you can input in the python window. Below is a list of all the current waypoint actions and the characters / keys associated with them:

'c' = Click (Click the left mouse button at the point on your screen where your mouse is hovering)

'd' = Double click (Double-click the left mouse button at the point on your screen where your mouse is hovering)

't' = Tab (Hit tab)

'e' = Enter (Hit enter)

'p' = Paste (Type out a specified text)
After hitting this key, return to the python window to input the desired text. After this is complete, the script will resume listening for other waypoints.

'i' = Insert Data (Insert / type data from the current Excel file)
After hitting this key, return to the python window to input the Excel sheet that the data you want to be inserted is in, as well as the section in that sheet that the data is located at. If you are using file-based execution, enter the column and row that the data in contained in. If you are using row-based execution, enter the column that the data is contained in each row. If the data is a datetime type data, it will be formatted according to the format string you configured, or as a default string if no format string is configured. After this is complete, the script will resume listening for other waypoints.

'w'= Wait (Wait for a specified number of seconds). 
After hitting this key, return to the python window to input the desired number of seconds. After this is complete, the script will resume listening for other waypoints.

Once you are finished, hit the 'esc' key to tell the script to stop listening for waypoints and create the config.json file.


# Running the Script

Once you are done configuring the script, you are ready to run it!
First, ensure that all the desired Excel files are stored in the "Excel" directory / folder.
Then, type the following bash command in the project directory:

"python3 ./auto_excel.py"

A countdown will begin, which gives you time to complete the final step: Moving your mouse cursor to the screen / window / website where you want to execute the Excel data.
That's it! If everything is set up correctly, the script will now automatically extract data from the given Excel files and follow the waypoints you set up to input the data into the website!


# Common Gotchas

There are a few common gotchas that may pop up while using this script:

*Non-Static Screen*:

Since this script interfaces directly with your computer screen, any slight change in the screen could throw off the script completely. Make sure that your screen stays completely static while the program is running!

*Non-Looping configuration*:

This script executes the waypoints you configured for every Excel file in the "Excel" directory, so it is critical that your configuration brings the script back to the "starting place" so that it can rinse and repeat with the data from the next Excel file.

*Wrong Excel sheet / columns / rows*:

The way that python reads Excel files may eliminate "empty whitespace" around the Excel data. So, if your first row in your Excel file is completely blank, this script will not read it and instead treat the second row as the first row, assuming the second row has data. To make sure you have the right column and row, execute a quick test run for a sample Excel file, look at what is printed out under the Excel data section, find the column and row of the desired data on the array, and use that as your column and row.
Sometimes errors may arise if the script cannot access the Excel sheet you specified. Make sure the case matches, and you count the columns and rows correctly!

*Wierd Excel data*:

Excel files can contain some unusual data types that this script is not built to handle. This script converts whatever is in the specified cell into a string before typing it out, but anything outside of text or numbers may be an issue.
