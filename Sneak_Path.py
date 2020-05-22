'''
This Python script streamlines the sneak path test required for DV-testing according to the regulations of CS.00054.

Sneak Path Test: Apply 13.5V to the DUT for at least one minute, disconnect power
Requirement: Voltage must fall below 1V within 1 second.

Program: Controls R&S RTB2004 Oscilloscope,R&S HMC8041 Power Supply.

Initializes power supply with 13.5V and max current. 
Allows the user to input number of channels used, set vertical scales, time-base, trigger level to initialize scope.
User chooses when to cut power to the DUT. When power is cut, infinite loop is broken, scope freezes picture, and allows user to set vertical cursors to measure 1 second. 

Python 3.7.7
PyVISA 1.10.1
'''

import pyvisa
import numpy as np
import time

#VISA Resource Addresses Configuration
rm = pyvisa.ResourceManager()
RTB = rm.open_resource('TCPIP0::172.21.120.136::inst0::INSTR') #R&S RTB2004 Scope
HMC = rm.open_resource('TCPIP0::172.21.120.111::inst0::INSTR') #R&S HMC8041 Power Supply

#Query the names of the instruments being used & RST instrument to default settings.
print(RTB.query("*IDN?"))
print(HMC.query("*IDN?"))

#Run for Five Test Cycles
for test_num in range(1,6):
    #Reset both instruments to default settings
    RTB.write("*RST")
    HMC.write('*RST')
    RTB.write("CHAN1:STAT ON") #RTB Scope Channel 1 ON.

    HMC.write("SOURce:VOLTage:LEVel:IMMediate:AMPLitude 13.5") #Set Power Supply Voltage to 13.5, as required by CS.00054 sneak path test
    HMC.write("SOURce:CURRENT:LEVel:IMMediate:AMPLitude MAX")   #Set Power Supply Current to MAX to quickly charge Caps

    #Change Scope Settings Here
    Num_Of_Channels = 1
    vertical_Scale_1 = 5 #V/div
    vertical_Scale_2 = 5 #V/div
    vertical_Scale_3 = 5 #V/div
    vertical_Scale_4 = 5 #V/div
    Time_Base = 500e-3   #seconds
    trigger_Level = 13.4

    #Turn Required Channels On
    for channel in range(1, Num_Of_Channels + 1):
        RTB.write("CHAN{}:STAT ON".format(channel))

    RTB.write("CHAN1:SCALe {}".format(vertical_Scale_1))
    RTB.write("CHAN2:SCALe {}".format(vertical_Scale_2))
    RTB.write("CHAN3:SCALe {}".format(vertical_Scale_3))
    RTB.write("CHAN4:SCALe {}".format(vertical_Scale_4))

    RTB.write("TIMebase:SCALe {}".format(Time_Base)) #Timebase Setting 
    RTB.write("TRIGger:A:LEVel {}".format(trigger_Level)) #Trigger Level Setting
    RTB.write("TRIGger:A:EDGE:SLOPe NEGative") #Set trigger to falling edge. 

    HMC.write("OUTPut ON") #Turn Power Supply output ON to send signal

    #Enable measurements to display on screen
    RTB.write("MEASurement1:ENABle ON")
    RTB.write("MEASurement2:ENABle ON")
    RTB.write("MEASurement3:ENABle ON")

    RTB.write("MEASurement1:MAIN UPEakvalue") #measure peak-positive voltage on screen
    RTB.write("MEASurement2:MAIN LPEakvalue") #measure peak-negative voltage on screen
    RTB.write("MEASurement3:MAIN FTIMe")      #measure fall-time on screen

    #Cursor settings
    RTB.write("CURSor:STATe ON") #turn cursors ON
    RTB.write("CURSor:FUNCtion HVER") #Use both Horizontal and Vertical cursors simultaneously 
    RTB.write("CURSor:Y1Position 0")   #Set horizontal cursor 1 to position y1
    RTB.write("CURSor:Y2Position 1.03")    #Set horizontal cursor 2 to position y2

    #Continuously check for power supply output state: If ON, keep running, if OFF, stop scope and break loop
    while True:
            On_Or_Off = HMC.query("OUTPut?") 
            if On_Or_Off == '0\n':
                time.sleep(2.5)
                RTB.write("STOP")
                break

    RTB.write("SEARch:CONDition EDGE")
    RTB.write("SEARch:SOURce CH1")
    RTB.write("SEARch:TRIGger:EDGE:LEVel 13")
    RTB.write("SEARch:STATe ON")

    search_result = RTB.query("SEARch:RESult1?").split(',')
    time_of_drop = float(search_result[1]) 

    RTB.write("CURSor:X1Position {}".format(time_of_drop))
    RTB.write("CURSor:X2Position {}".format(time_of_drop + 1)) #Make sure there is a delta t of 1s

    time.sleep(2)
    RTB.write("SEARch:STATe OFF")
    #Taking Screenshot of Result and saving it to front USB
    RTB.write("MMEMory:MDIRectory '/USB_FRONT/Sneak_Path' ")
    RTB.write("MMEMory:CDIRectory '/USB_FRONT/Sneak_Path' ")

    RTB.write("MMEMory:NAME 'Sneak{}'".format(test_num))
    RTB.write("HCOPy:FORMat PNG")
    RTB.write("HCOPy:IMMediate")
    HMC.write("OUTPut ON")

print("Five cycles complete")

#Writing data to excel
RTB.write("FORMat:DATA ASCii, 0")
RTB.write("CHAN1:DATA:POINts DEFault")
values = np.array(RTB.query_ascii_values("CHAN1:DATA?"))

df = pd.DataFrame(values)
filepath = 'Sneak_Path.xlsx'
df.to_excel(filepath, index = False)

