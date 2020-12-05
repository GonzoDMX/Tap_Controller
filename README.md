# Tap_Controller
Control Application for LED Tap Shoes wireless system

This application connects to an custom-made transmitter (based on Adafruit Feather M0) over USB. It allows direct control of the transmitter to control LED Tap Shoes. The application is also capable of creating a cue list and exporting cuelists to .xml. The application can also configure the DMX address of the transmitters for when they are being used in a perfomance.

I included the Inno Scripts as a reference for anyone building .exe executables for windows.
The Inno Script targets a python build created with CX_Freeze for anyone interested in deploying Python Apps on Windows. The setup.py file is where CX_Freeze finds the configuration and dependencies of the App.

The application is built in python using wxWidgets and xlsxwriter packages.

The rest of the project this app was developed for was under contract and I do not have permission to release the rest of the code. The Tap Controller App was developed on the side to learn wxWidgets for python and to experiement with a few other things as well as creating a debugging tool while developing the LED Tap Shoes.

More details on the LED Tap Shoes project can be found here: https://www.andrewoshei.com/projects
