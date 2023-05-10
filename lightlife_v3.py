import clr
from termcolor import colored, cprint
# Import python sys module
import sys
import math
# Import os module
import os
import time
# Import System.IO for saving and opening files
from System.IO import *
# Import C compatible List and String
from System import String
from System.Collections.Generic import List

# Add needed dll references
sys.path.append(os.environ['LIGHTFIELD_ROOT'])
sys.path.append(os.environ['LIGHTFIELD_ROOT']+"\\AddInViews")
clr.AddReference('PrincetonInstruments.LightFieldViewV5')
clr.AddReference('PrincetonInstruments.LightField.AutomationV5')
clr.AddReference('PrincetonInstruments.LightFieldAddInSupportServices')

# PI imports
from PrincetonInstruments.LightField.Automation import Automation
from PrincetonInstruments.LightField.AddIns import CameraSettings
from PrincetonInstruments.LightField.AddIns import DeviceType
from PrincetonInstruments.LightField.AddIns import ExperimentSettings
from PrincetonInstruments.LightField.AddIns import SpectrometerSettings


#=============================LOGO==============================#



logo = '''



                ██╗     ██╗ ██████╗ ██╗  ██╗████████╗██╗     ██╗███████╗███████╗    ██████╗     ██████╗ 
                ██║     ██║██╔════╝ ██║  ██║╚══██╔══╝██║     ██║██╔════╝██╔════╝    ╚════██╗   ██╔═████╗
                ██║     ██║██║  ███╗███████║   ██║   ██║     ██║█████╗  █████╗       █████╔╝   ██║██╔██║
                ██║     ██║██║   ██║██╔══██║   ██║   ██║     ██║██╔══╝  ██╔══╝       ╚═══██╗   ████╔╝██║
                ███████╗██║╚██████╔╝██║  ██║   ██║   ███████╗██║██║     ███████╗    ██████╔╝██╗╚██████╔╝
                ╚══════╝╚═╝ ╚═════╝ ╚═╝  ╚═╝   ╚═╝   ╚══════╝╚═╝╚═╝     ╚══════╝    ╚═════╝ ╚═╝ ╚═════╝ 
                                                                                                        

                
                '''


print(colored(logo, "blue"))
print("\n")
print(colored("                     Adpated by Tarek Atyia at the University of Tennessee, Knoxville from code", 'red'))
print(colored("     by Sabbir Liakat at Princeton Teledyne Instruments and from code by Diego Alonso Alvarez at Prior Scientific", 'red'))

print("\n")

#==================================END LOGO===================================#


#==========================PRINCETON FUNCTIONS=================================#

def add_devices():
    for device in experiment.AvailableDevices:
        print("Adding device...")
        experiment.Add(device)
        return device


def set_value(setting, value):    
    # Check for existence before setting
    # gain, adc rate, or adc quality
    if experiment.Exists(setting):
        experiment.SetValue(setting, value)

def find_spec():
    # Find connected device
    for device in experiment.ExperimentDevices:
        if (device.Type == DeviceType.Spectrometer):
            print("SPECTROMETER LOCATED")
            return True
    
    print("SPECTROMETER NOT LOCATED")
    return False

def find_cam():
    for device in experiment.ExperimentDevices:
        if device.Type == DeviceType.Camera:
            print("CAMERA LOCATED")
            return True
    print("CAMERA NOT LOCATED")
    return False



def save_file(filename):    
    # Set the base file name
    experiment.SetValue(
        ExperimentSettings.FileNameGenerationBaseFileName,
        Path.GetFileName(filename))
    
    # Option to Increment, set to false will not increment
    experiment.SetValue(
        ExperimentSettings.FileNameGenerationAttachIncrement,
        False)

    # Option to add date
    experiment.SetValue(
        ExperimentSettings.FileNameGenerationAttachDate,
        False)

    # Option to add time
    experiment.SetValue(
        ExperimentSettings.FileNameGenerationAttachTime,
        False)
#===========================END PRINCETON FUNCTIONS==========================#




#==============================START STAGE MAPPING==============================#
from ctypes import WinDLL, create_string_buffer
import os
import sys

path = "C:\\Users\\Argon\\Downloads\\PriorSDK-1.7.0\\PriorSDK 1.7.0\\x64\\PriorScientificSDK.dll"
#===========================ERROR CHECKING======================#
if os.path.exists(path): 
    SDKPrior = WinDLL(path)
else:
    raise RuntimeError("DLL could not be loaded.")

rx = create_string_buffer(1000)

def cmd(msg):
    print(msg)
    ret = SDKPrior.PriorScientificSDK_cmd(
        sessionID, create_string_buffer(msg.encode()), rx
    )
    if ret:
        print(f"API ERROR {ret}")
    else:
        print(f"SUCCESSFUL COMMAND {rx.value.decode()}")
        print("\n")

    return ret, rx.value.decode()


ret = SDKPrior.PriorScientificSDK_Initialise()
if ret:
    print(f"!INITIALIZATION ERROR! || return_val={ret} DECODE")
    sys.exit()
else:
    print(f"INITIALIZING CONNECTION TO API... {ret} DECODE")


ret = SDKPrior.PriorScientificSDK_Version(rx)
print(f"DLL API VERSION || return_val={ret}, version={rx.value.decode()}")


sessionID = SDKPrior.PriorScientificSDK_OpenNewSession()
if sessionID < 0:
    print(f"ERROR RETRIEVING SESSION ID || return_val={ret}")
else:
    print(f"SessionID = {sessionID}")


ret = SDKPrior.PriorScientificSDK_cmd(
    sessionID, create_string_buffer(b"dll.apitest 33 SUCCESSFUL CORRESPONDENCE"), rx
)
print(f"api response {ret}, rx = {rx.value.decode()}")
if(rx.value.decode() == "goodresponse"):
    print("DLL API LOAD SUCCESSFUL")
    print("\n")

ret = SDKPrior.PriorScientificSDK_cmd(
    sessionID, create_string_buffer(b"dll.apitest -300 stillgoodrespsonse"), rx
)
print(f"api response {ret}, rx = {rx.value.decode()}")
if(rx.value.decode() == "stillgoodresponse"):
    print("DLL API TEST SUCCESSFUL")
    print("\n")
#=================================END ERROR CHECKING==========================#






#================================START CONNECTION INTERFACE====================#
print('\n')
print("Connecting to Controller...")
cmd("controller.connect 76")
cmd("controller.stage.position.get")
cmd("controller.stage.goto-position 0 0")
print("\n")
text1 = "CONTROLLER SET TO DEFAULT POSITION"
text1cen = text1.center(100, '=')
print(colored(text1cen, 'green'))
text2 = "GO TO 'TASK MANAGER' AND ENSURE THAT YOU KILL ANY 'AddInProcess.exe'"
text2cen = text2.center(100, '=')
print("\n")
print(colored(text2cen, 'yellow'))
print('\n')
print("If LightField still crashes after killing the process, reload task manager and you should see an 'AddInProcess.exe'")
print("\n")
text3 = "Press ENTER to continue"
text3cen = text3.center(100, '=')
input(colored(text3cen, 'cyan'))
print('\n')
text4 = "INPUT DIRECTIONAL COMMANDS"
text4cen = text4.center(100, '=')
print(colored(text4cen, 'light_magenta'))
print('\n')
print("Press 'q' to QUIT")
print('\n')
le_input = input("Input desired shape (square, circle, spiral, figure 8, random, raster): ")
print('\n')

#=============================ERROR CHECKING==================================#
while le_input not in ['square', 'circle', 'spiral', 'figure 8', 'random', 'raster']:
    print("UNRECOGNIZED COMMAND")
    le_input = input("Please enter a valid command: ")
if le_input == 'q':
    exit
#======================================END ERROR CHECKING=======================#

#==============================================SQUARE===================================#
if le_input == "square":
    print("Attempting to Communicate with LightField...")
    auto = Automation(True, List[String]())
    experiment = auto.LightFieldApplication.Experiment
    find_spec()
    find_cam() 
    add_devices()
    exposureTime = input("Enter the exposure time for the camera in milliseconds (ENTER THIS VALUE AS A FLOAT i.e. 23.4): ")   
    set_value(CameraSettings.ShutterTimingExposureTime, exposureTime)
    print('\n')
    print("----INITIATING SQUARE----")
    print('\n')
    print("FOR REFERENCE: THE DIAMETER OF A NICKEL IS 21000 MICRONS")
    print('\n')
    x = int(input("X-Direction (in microns): "))
    y = int(input("Y-Direction (in microns): "))

    cmd("controller.stage.move-relative " + str(x) + " 0")
    experiment.Acquire()
    time.sleep(2)
    cmd("controller.stage.move-relative 0 " + str(y))
    experiment.Acquire()
    time.sleep(2)
    cmd("controller.stage.move-relative " + "-" + str(x) + " 0")
    experiment.Acquire()
    time.sleep(2)
    cmd("controller.stage.move-relative 0 " + "-" + str(y))
    experiment.Acquire()
    time.sleep(2)
    cmd("controller.stage.position.get")
    print('\n')
    input("Press ENTER to exit")
#===========================================END SQUARE=====================================#







#=========================================CIRCLE=========================================#
if le_input == "circle":
    print("----INITIATING CIRCLE----")
    print("\n")
    print("FOR REFERENCE: THE DIAMETER OF A NICKEL IS 21000 MICRONS")
    print('\n')
    radius = float(input("Enter the radius of the circle in microns: "))
    cen_x = float(input("Enter the x coordinate for the center of the circle: "))
    cen_y = float(input("Enter the y coordinate for the center of the circle: "))
    num_measurements = int(input("enter number of measurements: "))
    print("Attempting to Communicate with LightField...")
    auto = Automation(True, List[String]())
    experiment = auto.LightFieldApplication.Experiment
    add_devices()
    find_spec()
    find_cam() 
    exposureTime = input("Enter the exposure time for the camera in milliseconds (ENTER THIS VALUE AS A FLOAT i.e. 23.4): ")   
    set_value(CameraSettings.ShutterTimingExposureTime, exposureTime)
    cmd("controller.stage.move-relative {} {}".format(cen_x, cen_y))
    points = []
    for i in range(num_measurements):
        theta = 2*math.pi*i / num_measurements
        x = cen_x + float(radius)*math.cos(theta)
        y = cen_y + float(radius)*math.sin(theta)
        points.append((x, y))
    for point in points:
        cmd("controller.stage.move-relative {} {}".format(str(point[0]), str(point[1])))
        experiment.Acquire()
        time.sleep(2)
    print('\n')

    input("Press ENTER to exit")
#============================================END CIRCLE================================# 



#==========================================S-SHAPE================================#

def generate_figure_8(num_points):
    # Calculate the step size between each point
    step_size = 2 * math.pi / (num_points - 1)

    # Generate points for the S-shape
    points = []
    for i in range(num_points):
        theta = i * step_size
        x = 5000.0 + 5000.0 * math.cos(theta)
        y = 5000.0 + 5000.0 * math.sin(2 * theta)
        points.append((x, y))
    
    return points

if le_input == "figure 8":
    print("INITIATING FIGURE 8")
    print('Attempting to communicate with LightField...')
    print('\n')
    auto = Automation(True, List[String]())
    experiment = auto.LightFieldApplication.Experiment
    add_devices()
    find_spec()
    find_cam() 
    exposureTime = input("Enter the exposure time for the camera in milliseconds (ENTER THIS VALUE AS A FLOAT i.e. 23.4): ")   
    set_value(CameraSettings.ShutterTimingExposureTime, exposureTime)
    num_points = int(input("Enter the number of measurements you want to take: "))

    points = generate_figure_8(num_points)
    for point in points:
        cmd("controller.stage.move-relative {} {}".format(str(point[0]), str(point[1])))
        experiment.Acquire()
        time.sleep(2)
    input("Press ENTER to exit")
    
if le_input == "random":
    print("----INITIATING RANDOM----")
    print("\n")
    amplitude = float(input("Enter the amplitude of the S shape in microns: "))
    cen_x = float(input("Enter the x coordinate for the center of the S shape: "))
    cen_y = float(input("Enter the y coordinate for the center of the S shape: "))
    num_measurements = int(input("Enter the number of measurements: "))
    print("Attempting to Communicate with LightField...")
    auto = Automation(True, List[String]())
    experiment = auto.LightFieldApplication.Experiment
    add_devices()
    find_spec()
    find_cam() 
    exposureTime = input("Enter the exposure time for the camera in milliseconds (ENTER THIS VALUE AS A FLOAT, e.g., 23.4): ")   
    set_value(CameraSettings.ShutterTimingExposureTime, exposureTime)
    cmd("controller.stage.move-relative {} {}".format(cen_x, cen_y))
    
    points = []
    for i in range(num_measurements):
        theta = 2 * math.pi * i / num_measurements
        x = cen_x + i * amplitude / num_measurements
        y = cen_y + amplitude * math.sin(2 * theta)
        points.append((x, y))
    
    for point in points:
        cmd("controller.stage.move-relative {} {}".format(str(point[0]), str(point[1])))
        experiment.Acquire()
        time.sleep(2)
    
    input('Press ENTER to exit')
    

if le_input == "raster":
    print("----INITIATING RASTER SCAN----")
    print("\n")
    print("FOR REFERENCE: THE DIAMETER OF A NICKEL IS 21000 MICRONS")
    print('\n')
    width = float(input("Enter the width of the raster scan in microns: "))
    height = float(input("Enter the height of the raster scan in microns: "))
    cen_x = float(input("Enter the x coordinate for the starting point of the raster scan: "))
    cen_y = float(input("Enter the y coordinate for the starting point of the raster scan: "))
    num_points_x = int(input("Enter the number of points along the x-axis: "))
    num_points_y = int(input("Enter the number of points along the y-axis: "))
    print("Attempting to Communicate with LightField...")
    auto = Automation(True, List[String]())
    experiment = auto.LightFieldApplication.Experiment
    add_devices()
    find_spec()
    find_cam() 
    exposureTime = input("Enter the exposure time for the camera in milliseconds (ENTER THIS VALUE AS A FLOAT i.e. 23.4): ")   
    set_value(CameraSettings.ShutterTimingExposureTime, exposureTime)
    cmd("controller.stage.goto-position {} {}".format(cen_x, cen_y))
    
    points = []
    for i in range(num_points_y):
        y = cen_y + i * height / (num_points_y - 1)
        if i % 2 == 0:
            for j in range(num_points_x):
                x = cen_x + j * width / (num_points_x - 1)
                points.append((x, y))
        else:
            for j in range(num_points_x - 1, -1, -1):
                x = cen_x + j * width / (num_points_x - 1)
                points.append((x, y))
    
    for point in points:
        cmd("controller.stage.goto-position {} {}".format(str(point[0]), str(point[1])))
        experiment.Acquire()
        time.sleep(2)
    
    input("Press ENTER to exit")