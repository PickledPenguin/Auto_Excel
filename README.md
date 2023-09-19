# Description

This script extracts data from the Excel files in the "Excel" directory and outputs the gathered data for each Excel file according to a configuration set by the user via simulating mouse movement and keypresses.

I have written this README file to be as informative as possible for readers who are not as experienced in handling scripts such as this one and have included links to helpful reasources


# Project Background Info

This is a script I created for the Rochester Institute of Technology's Software Engineering department to help automate inputting the student employees' hours from Excel spreadsheets into an external website. Previously, this had to be done manually, cell-by-cell, line-by-line, which was tedious and unnecessary. So I decided to build this script to fix that problem and allow the SEO employees to free up a significant portion of their time for less drudging, more pressing tasks.


# Architecture Details

This script does not determine where to input the Excel data from the external website itself but rather relies on a manual setup process in which the user directs the program to the points of interest (waypoints) on the screen. There are three main reasons I went with this approach: First, I was unable to access the external website while developing this script as it was restricted and only accessible by members of the SE department. Second, in the case that the SE department decides to use a different external website or the original external website is changed or updated, this script would still function, and it would only need to be reconfigured rather than redesigned. Lastly, a waypoint system can be more easily understood by someone newer to running scripts and "coding" in general, as opposed to an xpath approach or another more Website-orented or HMTL-based configuration.
This flexibility does come with a few drawbacks: The initial setup process can take a while and the script is dependent on the screen being completely static and consistent. I think these drawbacks are worth the flexibility and simplicity, at least for this solo small-scale project.


# Required Installations

1.) python3 [link to install python3](https://www.python.org/downloads/)

2.) pip3 [link to install pip3](https://pip.pypa.io/en/stable/installation/)

3.) All modules listed in requirements.txt [installing requirements](https://note.nkmk.me/en/python-pip-install-requirements/)


# Configuring the Script

To configure the script, first open a [cmd window](https://www.wikihow.com/Open-the-Command-Prompt-in-Windows) if you have a Windows computer or open a [terminal window](https://support.apple.com/guide/terminal/open-or-quit-terminal-apd5265185d-f365-44cb-8b09-71a064a42125/mac#:~:text=Terminal%20for%20me-,Open%20Terminal,%2C%20then%20double%2Dclick%20Terminal.)) if you have an Apple or Linux computer. Then [navigate to the directory](https://www.wikihow.com/Change-Directories-in-Command-Prompt) you saved this project to and type the following command inside the project directory:

"python3 ./setup.py"

You will now be able to configure the script's wait time between actions, datetime format string, and waypoints in the command line, by entering the appropriate number shown next to the appropriate option.

*Setting wait time*: 

The wait time is the time (in seconds) between any meaningful action the script simulates. This gives the website time to catch up and helps make sure that no inputs or data is lost. If your connection is a little unstable, I would recommend increasing this wait time. The default wait time is 0.5 seconds.

*Setting datetime format string*: 

The datetime format string is a string that uses certain format codes as standard directives for specifying the format in which you want to represent datetime type data. A comprehensive list of all the format codes can be found [here] (https://strftime.org) Using the format codes, you can "insert" parts of the datetime data into your desired format. For example, the format string: "%H:%M %p" will format the data like this: "**Hour**:**Minute** **AM or PM**" and the format string: "Student completed work on %m/%d/%Y, which was a %A" will format the data like this: "Student completed work on **month**/**day**/**year**, which was a **Day of the week**"

*Setting up "waypoints"*: 

Waypoints are points of interest that you indicate on the screen for the program to either click, input data, or perform some other action. There are 5 different types of waypoints: Click, Double click, Tab, Enter, Paste, Insert, and Wait.
Each type of waypoint has a key on a keyboard associated with it. To create a waypoint, move your mouse to the point on the screen where you want the waypoint to be executed, then hit the key on your keyboard that matches the appropriate waypoint action. Some waypoints require additional data which you can input in the python window. Below is a list of all the current waypoint actions and the characters / keys associated with them:

'c' = Click (Click the left mouse button at the point on your screen where your mouse is hovering)

'd' = Double click (Double-click the left mouse button at the point on your screen where your mouse is hovering)

't' = Tab (Hit tab)

'e' = Enter (Hit enter)

'p' = Paste (Type out a specified text)
After hitting this key, you can return to the command window to input the desired text. After this is completed, the script will resume listening for other waypoints.

'i' = Insert Data (Insert / type data in a specified sheet, column, and row of the current Excel file)
After hitting this key, return to the command window to input the Excel sheet in which the data you want to be inserted, as well as the column and row of the cell in that sheet where the data is located. If the data is a datetime type data, it will be formatted according to the format string you configured, or as a default string if no format string is configured. After this is complete, the script will resume listening for other waypoints.

'w'= Wait (Wait for a specified number of seconds). 
After hitting this key, return to the command window to input the desired number of seconds. After this is complete, the script will resume listening for other waypoints.

Once you are finished, hit the 'esc' key to tell the script to stop listening for waypoints and create/edit the config.json file.


# Running the Script

Once you are done configuring the script, you are ready to run it!
First, ensure that all the desired Excel files are stored in the "Excel" directory / folder.
Then, type the following command within the project directory:

"python3 ./auto_excel.py"

A countdown will begin, which gives you time to complete the final imporant step: Moving your mouse cursor to the screen / window / website where you want to execute the Excel data and clicking on that window to ensure the first click or input will register.
That's it! If everything is set up correctly, the script will now automatically extract data from the given Excel files and follow the waypoints you set up to input the data into the website!


# Common Gotchas

There are a few common gotchas that may pop up while using this script:

*Non-Static Screen*:

Since this script interfaces directly with your computer screen, any slight change in the screen could throw off the script completely. Make sure that your screen stays completely static while the program is running!

*Non-Looping configuration*:

This script executes the waypoints you configured for every Excel file in the "Excel" directory, so it is critical that your configuration brings the script back to the "starting place" so that it can rinse and repeat with the data from the next Excel file.

*Wrong Excel sheet / columns / rows*:

The way that python reads Excel files may eliminate "empty whitespace" around the Excel data. So, if the first row in your Excel file is completely blank, this script will not read it and instead treat the second row as the first row, assuming the second row has data. To make sure you have the right column and row, execute a quick test run for a sample Excel file, look at what is printed to the command line out under the Excel data section, find the column and row of the desired data on the array, and use that as your column and row.
Sometimes errors may arise if the script cannot access the Excel sheet you specified. Make sure the case matches, and you count the columns and rows correctly!

*Wierd Excel data*:

Excel files can contain some unusual data types that this script is not built to handle. This script converts whatever is in the specified cell into a string before typing it out, but anything outside of text or numbers may be an issue.
