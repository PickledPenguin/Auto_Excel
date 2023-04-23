import sys

from pynput.mouse import Controller
from pynput import keyboard
from pynput.keyboard import Key, Listener
import os
import json

config_counter = 0
mouse = Controller()
wait_input = False
paste_input = False
data_insert = False
configuration = {}


def optional_set_wait_time():
    """ Set the wait time between actions """

    wait_configuration = {"WAIT_DELAY_IN_SECONDS": 0.5}

    while True:
        wait_time = input("Enter your your preferred wait time (between 0.0 and 100.0) in seconds between actions. ("
                          "Press enter for default value of 0.5 seconds): ")
        try:
            if wait_time == '':
                print("Wait time set to default value of 0.5 seconds")
                wait_time = 0.5
                break
            elif not (0.0 <= float(wait_time) <= 100.0):
                print(f"Wait time must be between 0.0 and 100.0 seconds (you entered: {float(wait_time)} seconds)")
                continue
            else:
                print(f"Wait time set to {float(wait_time)} seconds")
                break
        except (ValueError, NameError):
            print(f"Wait time must be a number between 0.0 and 100.0 (you entered: {float(wait_time)})")
            continue

    wait_configuration["WAIT_DELAY_IN_SECONDS"] = float(wait_time)

    return wait_configuration


def optional_set_datetime_format():
    """ Set the format for datetime type data """

    format_configuration = {"DATETIME_FORMAT_STR": "default"}

    format_string = input("Enter the format string for datetime type data (press enter for default string "
                          "conversion): ")
    # if enter was pressed
    if format_string == '':
        print("Format set to default string conversion")
        format_string = 'default'
    # if there is a format string
    else:
        print(f"Format set to {format_string}")

    format_configuration["DATETIME_FORMAT_STR"] = format_string

    return format_configuration


def optional_set_execution_type():
    """ Set the configuration execution type """

    execution_configuration = {"EXECUTION_TYPE": "file"}

    print("This program supports 2 execution types: file-based execution and row-based execution.\n"
          "file-based execution: Run the waypoint configuration for each excel file in the Excel directory\n"
          "row-based execution: Run the waypoint configuration for each row in each excel file in the Excel directory\n")

    execution_type_input = input("Enter your preferred execution type. Enter \"f\" for file-based, \"r\" for "
                                 "row-based, or enter for default (file-based): ")

    if execution_type_input == 'f':
        print("Execution type set to file-based execution")
        execution_type = 'file-based'
    elif execution_type_input == 'r':
        print("Execution type set to row-based execution")
        execution_type = 'row-based'
    else:
        print("Execution type set to default file-based execution")
        execution_type = 'file-based'

    execution_configuration["EXECUTION_TYPE"] = execution_type

    return execution_configuration


def optional_set_ignore_header():
    header_configuration = {"IGNORE_HEADER": False}

    ignore_header = input("Enter your desired header setting. Enter \"i\" to ignore headers, \"k\" to keep the "
                          "headers, or enter for default (keep headers): ")

    if ignore_header == "i":
        print("\"i\" pressed, ignoring first row of every Excel file")
        header_configuration["IGNORE_HEADER"] = True
    elif ignore_header == "k":
        print("\"k\" pressed, keeping first row of every Excel file")
        header_configuration["IGNORE_HEADER"] = False
    else:
        print("\"enter\" pressed, using default of keeping first row of every Excel file")
        header_configuration["IGNORE_HEADER"] = False

    return header_configuration


def on_release(key):
    """ Listen for waypoints, record mouse position at those waypoints, and return the information """

    global configuration, config_counter, mouse, wait_input, paste_input, data_insert

    if key == keyboard.Key.esc:
        # Stop listener
        print("Esc pressed, stopped listening for waypoints")
        print("------------------------------------------------------")
        return False

    elif hasattr(key, 'char'):

        # click waypoint
        if key.char == 'c':
            print("\"c\" pressed - \"click\" waypoint at point [%f, %f]" % (mouse.position[0], mouse.position[1]))
            configuration["WAYPOINTS"][config_counter] = {"type": "click", "pos": mouse.position}

        # double click waypoint
        elif key.char == 'd':
            print("\"d\" pressed - \"double click\" waypoint at point [%f, %f]" % (
                mouse.position[0], mouse.position[1]))
            configuration["WAYPOINTS"][config_counter] = {"type": "double-click", "pos": mouse.position}

        # press key waypoint
        elif key.char == 'p':
            print("\"p\" pressed - \"paste\" waypoint at point")
            configuration["WAYPOINTS"][config_counter] = {"type": "paste"}
            paste_input = True
            print("------------------------------------------------------")
            print("Paused listening for waypoints, please input what you want pasted below")
            print("------------------------------------------------------")
            return False

        # tab waypoint
        elif key.char == 't':
            print("\"t\" pressed - \"tab\" waypoint")
            configuration["WAYPOINTS"][config_counter] = {"type": "tab"}

        # enter waypoint
        elif key.char == 'e':
            print("\"e\" pressed - \"enter\" waypoint")
            configuration["WAYPOINTS"][config_counter] = {"type": "enter"}

        # insert data waypoint
        elif key.char == 'i':
            print("\"i\" pressed - \"insert data\" waypoint")
            configuration["WAYPOINTS"][config_counter] = {"type": "insert-data"}
            data_insert = True
            print("------------------------------------------------------")
            print("Paused listening for waypoints, please input desired column and row the data is contained in below")
            print("------------------------------------------------------")
            return False

        # wait waypoint
        elif key.char == 'w':
            print("\"w\" pressed - \"wait\" waypoint.")
            configuration["WAYPOINTS"][config_counter] = {"type": "wait", "seconds": 1}
            wait_input = True
            # stop listening
            print("------------------------------------------------------")
            print("Paused listening for waypoints, please input desired wait time below")
            print("------------------------------------------------------")
            return False

        else:
            # decrement to offset increment
            config_counter -= 1
        # increment config counter
        config_counter += 1


def resume_listening():
    with keyboard.Listener(on_release=on_release) as listener:
        listener.join()
    input_data()


def input_data():
    """listen for waypoints and gather user input for waypoints until the user hits esc to stop listening and
    finalize the config.json file"""

    global configuration, config_counter, wait_input, paste_input, data_insert

    # if listening ended to collect input for a wait waypoint
    if wait_input:
        while True:
            # get input from the user
            wait_time = input("How many seconds would you like to wait for?: ")
            try:
                # if the wait time is unreasonably large
                if int(wait_time) > sys.maxsize:
                    print(
                        f"ERROR: you have set the \"wait\" waypoint for more than {sys.maxsize} seconds, which "
                        f"is FAR too long!")
                    continue
                # if the wait time is negative
                elif int(wait_time) < 0:
                    print(f"\"wait\" waypoint cannot be negative (you entered {int(wait_time)} seconds)")
                    continue
                # if the wait time is a valid number
                else:
                    print(f"\"wait\" waypoint set for {int(wait_time)} seconds")
                    configuration["WAYPOINTS"][config_counter].setdefault("seconds", int(wait_time))
                    config_counter += 1
                    # set flag variable to False
                    wait_input = False

                    # add a warning if the wait time is larger than an hour
                    if int(wait_time) >= 3600:
                        print(f"WARNING: you set the \"wait\" waypoint for {int(wait_time)} seconds, which is "
                              f"longer than an hour!")

                    print("------------------------------------------------------")
                    print("Resumed listening for waypoints")
                    print("------------------------------------------------------")
                    break
            # if the wait time is not a number
            except ValueError:
                print(f"\"wait\" waypoint must be set as a number (you entered {str(wait_time)})")
                continue

        resume_listening()

    # if listening was ended to collect input for a paste waypoint
    if paste_input:
        while True:
            # get input from the user
            paste_str = input("What would you like to paste?: ")
            print(f"\"paste\" waypoint set to \"{paste_str}\"")
            configuration["WAYPOINTS"][config_counter].setdefault("paste", paste_str)
            config_counter += 1
            # set flag variable to False
            paste_input = False
            print("------------------------------------------------------")
            print("Resumed listening for waypoints")
            print("------------------------------------------------------")
            break

        resume_listening()

    # if listening was ended to collect input for an insert data waypoint
    if data_insert:
        while True:
            if configuration["EXECUTION_TYPE"] == "row-based":
                # get input from the user
                col = input(f"Enter the column the data is contained in (1st column = 1, "
                            f"2nd column = 2, etc): ")
                try:
                    if int(col) <= 0:
                        print(
                            f"\"input\" waypoint cannot have a negative column (you entered column: {int(col)})")
                        continue
                    # if the col and row are valid numbers
                    else:
                        print(f"\"input\" waypoint set for data")
                        configuration["WAYPOINTS"][config_counter].setdefault("excel_col", int(col))
                        config_counter += 1
                        # set flag variable to False
                        data_insert = False
                        print("------------------------------------------------------")
                        print("Resumed listening for waypoints")
                        print("------------------------------------------------------")
                        break
                # if the col is not a number
                except ValueError:
                    print(
                        f"\"input\" waypoint must be a number (you entered column: {str(col)})")
                    continue

            else:
                # get input from the user
                sheet = input("Enter the name of the sheet that the data is contained in (Case matters!): ")
                col = input(f"Enter the column in the sheet \"{sheet}\" the data is contained in (1st column = 1, "
                            f"2nd column = 2, etc): ")
                row = input(f"Enter the row in the sheet \"{sheet}\" the data is contained in (1st row = 1, 2nd row = "
                            f"2, etc): ")
            try:
                # if the col or row is zero or negative
                if (int(col) <= 0) or (int(row) <= 0 and configuration["EXECUTION_TYPE"] != "row-based"):
                    print(
                        f"\"input\" waypoint cannot have a negative column or row (you entered column: {int(col)}, row: {int(row)})")
                    continue
                # if the col and row are valid numbers
                else:
                    print(f"\"input\" waypoint set for data")
                    configuration["WAYPOINTS"][config_counter].setdefault("sheet", sheet)
                    configuration["WAYPOINTS"][config_counter].setdefault("excel_row", int(row))
                    configuration["WAYPOINTS"][config_counter].setdefault("excel_col", int(col))
                    config_counter += 1
                    # set flag variable to False
                    data_insert = False
                    print("------------------------------------------------------")
                    print("Resumed listening for waypoints")
                    print("------------------------------------------------------")
                    break
            # if the col or row is not a number
            except ValueError:
                print(f"\"input\" waypoint must be set as a number (you entered column: {str(col)}, row: {str(row)})")
                continue

        resume_listening()


def default_config_values_if_necessary():
    global configuration

    configuration.setdefault("WAIT_DELAY_IN_SECONDS", 0.5)
    configuration.setdefault("DATETIME_FORMAT_STR", "default")
    configuration.setdefault("EXECUTION_TYPE", "file-based")
    configuration.setdefault("IGNORE_HEADER", False)
    configuration.setdefault("WAYPOINTS", {})

    return configuration


def config():
    """ Run configuration of wait time and waypoints and store collected data in the config.json file """

    global config_counter, configuration

    while True:

        print("\n1.) Set wait time between actions \n2.) Set format string for datetime type data \n3.) Set execution "
              "type (row-based or file-based)\n4.) Set header configuration (ignore header or keep header) \n5.) "
              "Set/overwrite waypoints\n6.) Save and Exit\n")

        response = input("Select an option: ")
        # make it look neat
        print()

        if response == '1':
            configuration["WAIT_DELAY_IN_SECONDS"] = optional_set_wait_time()["WAIT_DELAY_IN_SECONDS"]

        elif response == '2':
            configuration["DATETIME_FORMAT_STR"] = optional_set_datetime_format()["DATETIME_FORMAT_STR"]

        elif response == '3':
            configuration["EXECUTION_TYPE"] = optional_set_execution_type()["EXECUTION_TYPE"]

        elif response == '4':
            configuration["IGNORE_HEADER"] = optional_set_ignore_header()["IGNORE_HEADER"]

        # reconfigure waypoints
        elif response == '5':

            # overwrite waypoints
            configuration["WAYPOINTS"] = {}

            print("------------------------------------------------------")

            print("Now listening for waypoints. Move your mouse to a point of interest on your screen and hit one of "
                  "the following keys to create a waypoint:\n")
            print("\'c\' = Click (Click the left mouse button once at that point)\n")
            print("\'d\' = Double Click (Double-click the left mouse button at that point)\n")
            print("\'t\' = Tab (hit the Tab key)\n")
            print("\'e\' = Enter (hit the Enter key)\n")
            print("\'p\' = Paste (Paste a specified text). After hitting this key, return to the python window to "
                  "input the desired text. When this is complete, the script will resume listening for other "
                  "waypoints.\n")
            print("\'i\' = Insert Data (Insert / type out data in a specified section of the current Excel "
                  "file). After hitting this key, return to the python window to input the desired section."
                  "(for file-based execution, input the desired sheet, column and row)"
                  "(for row-based execution, input the desired column) "
                  "When this is complete, the script will resume listening for other waypoints.\n")
            print("\'w\' = wait (Wait for a specified number of seconds). After hitting this "
                  "key, return to the python window to input the desired wait time. When this is complete, the script "
                  "will resume listening for other waypoints.\n")
            print("Once you are finished, hit esc to end listening\n")

            # Collect events until released
            with keyboard.Listener(on_release=on_release) as listener:
                listener.join()
            input_data()

            config_counter = 0

        else:
            print("Saved and Exited")
            break

    return configuration


def reset_config():
    """ Reset the configuration """

    global configuration

    filename = "config.json"

    # load the old configuration from config.json so the user can modify it
    with open(filename, 'r') as f:
        configuration = json.load(f)

    # set the default values if necessary
    default_config_values_if_necessary()

    print(configuration)

    # RUN SETUP SCRIPT
    config()

    # remove the previous config.json file and make a new one
    try:
        os.remove(filename)
        print("Replacing config.json file:")
    except FileNotFoundError:
        print("No config.json file to remove, continuing")

    with open(filename, 'a+') as f:
        # create new config.json file
        json.dump(configuration, f, indent=4)

    return configuration


if __name__ == "__main__":
    # Reset the configuration
    config_file = reset_config()
    print("\n------------------------------------------------------")
    print("Current config.json file:")
    print(config_file)
    print("------------------------------------------------------")
