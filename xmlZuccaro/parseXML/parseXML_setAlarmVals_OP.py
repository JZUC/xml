# 12May2022 - In this script's current state, it parses the project xml file to find the ScaledMax and InputMax
#             setpoints from the original program (10178) and saves them to a dataframe. The new program (10249)
#             is then parsed to determine the location of those setpoints and replaces the existing values with the
#             values found in the old program. This includes all relevant alarm setpoints and their timer values.

# 10Jun2022 - Both scripts (with _OP as a suffix) work with the VC Still standard.

import xml.etree.ElementTree as ET
from xml.etree import ElementTree
import numpy as np
import pandas as pd


##### Enter project information #####

##### Set sysType to the number assigned to its system type
##### MP == 0
##### VC == 1
sysType = 1

operatorXML = ET.parse('C:/Users/jzuccaro/Desktop/Py_Demo/MECOStandard_L5X.L5X') # Use line 14719 for demo
operandXML = ET.parse('C:/Users/jzuccaro/Desktop/Py_Demo/MECO10248_L5X_SP.L5X') # Use line 15956 for demo

################


root = operatorXML.getroot()
root2 = operandXML.getroot()
dfOperatorSP = pd.DataFrame()
dfOperandSP = pd.DataFrame()
dfOperatorDatatypes = pd.DataFrame()
dfOperandDatatypes = pd.DataFrame()



i = 0

# Find all "Tag" elements containing the attribute DataType="Alarm_Tag"
for dataTypes in root.findall('./Controller/DataTypes/DataType'):
    # Get the value of the "Tag" attribute "Name"
    name1 = dataTypes.get('Name')
    dfOperatorDatatypes.loc[i, "Datatypes"] = name1

    i += 1

i = 0

# Find all "Tag" elements containing the attribute DataType="Alarm_Tag"
for dataTypes in root2.findall('./Controller/DataTypes/DataType'):
    # Get the value of the "Tag" attribute "Name"
    name2 = dataTypes.get('Name')
    dfOperandDatatypes.loc[i, "Datatypes"] = name2

    i += 1


i = 0
n = dfOperatorDatatypes.shape[0]

##### Set critical alarm values #####

while i < n:
    targetTag = str(dfOperatorDatatypes["Datatypes"].values[i])

###    targetTagStr = '"'+targetTag+'"'
    if targetTag == "AlmCtrl":
        operatorDatatype = '"AlmCtrl"'
    elif targetTag == "Alarm_Tag":
        operatorDatatype = '"Alarm_Tag"'
    elif targetTag == "Alm_Tag":
        operatorDatatype = '"Alm_Tag"'

    i += 1


i = 0
n = dfOperandDatatypes.shape[0]

while i < n:
    targetTag = str(dfOperandDatatypes["Datatypes"].values[i])

    ###    targetTagStr = '"'+targetTag+'"'
    if targetTag == "AlmCtrl":
        operandDatatype = '"AlmCtrl"'
    elif targetTag == "Alarm_Tag":
        operandDatatype = '"Alarm_Tag"'
    elif targetTag == "Alm_Tag":
        operandDatatype = '"Alm_Tag"'

    i += 1


if sysType == 0:
    delayPath = 'Data/Structure/StructureMember/[@Name="Alm"]/StructureMember/[@Name="DELAY"]'
    suDelayPath = 'Data/Structure/StructureMember/[@Name="Alm"]/StructureMember/[@Name="SU_DELAY"]'
    spPath = 'Data/Structure/StructureMember/[@Name="Alm"]/StructureMember/[@Name="SP"]/DataValueMember/[@Name="In"]'
elif sysType == 1:
    delayPath = 'Data/Structure/[@DataType="AlmCtrl"]/StructureMember/[@Name="Alm"]/StructureMember/[@Name="DELAY"]'
    suDelayPath = 'Data/Structure/[@DataType="AlmCtrl"]/StructureMember/[@Name="Alm"]/' \
                  'StructureMember/[@Name="SU_DELAY"]'


    spPath = 'Data/Structure/[@DataType="AlmCtrl"]/StructureMember/[@Name="Alm"]/' \
             'StructureMember/[@Name="SP"]/DataValueMember/[@Name="In"]'

operatorTagElement = "./Controller/Tags/Tag/[@DataType="+operatorDatatype+"]"
operandTagElement = "./Controller/Tags/Tag/[@DataType="+operandDatatype+"]"
# structElement = "Data/Structure/StructureMember/[@Name="+structMember+"]"

i = 0

##### Find and save critical alarm values #####

# Find all "Tag" elements containing the attribute DataType="Alarm_Tag"
for setpointTags in root.findall(operatorTagElement):
    # Get the value of the "Tag" attribute "Name"
    name = setpointTags.get('Name')
    dfOperatorSP.loc[i, "Alm Tags"] = name



    # Get the text contained in the "Description" element
#    description = setpointTags.find('Description').text

    # Find the alarm message data contained in the "Comment" element, within the "Tag" element,
    # with the attribute Operand=".ALM.MSG"
    # <Comment Operand=".ALM.MSG">
    # <![CDATA[(CF-SANIT) CF Sanitization Overdue Alarm]]>
    # </Comment>
    # for almIn in setpointTags.findall('Data/Structure/StructureMember/[@Name="SP"]/DataValueMember/[@Name="In"]'):
    #     almInSP = str(almIn.get('Value'))      ### try converting to int() like the examples below
    #     dfOperatorSP.loc[i, "In"] = almInSP

    for almIn in setpointTags.findall(spPath):
        almInSP = str(almIn.get('Value'))      ### try converting to int() like the examples below
        dfOperatorSP.loc[i, "In"] = almInSP
        print(almInSP)

    for delay in setpointTags.findall(delayPath):
        for delayIn in delay.findall('DataValueMember/[@Name="In"]'):
            delayInSP = int(delayIn.get('Value'))
            dfOperatorSP.loc[i, "Timer Delay In"] = str(delayInSP)

        for delayOut in delay.findall('DataValueMember/[@Name="UnitsSelect"]'):
            delayOutSP = int(delayOut.get('Value'))
            dfOperatorSP.loc[i, "Timer Delay Units Select"] = str(delayOutSP)

        for delayMinRem in delay.findall('DataValueMember/[@Name="Min_Rem"]'):
            delayMinRemSP = int(delayMinRem.get('Value'))
            dfOperatorSP.loc[i, "Timer Remaining Minutes"] = str(delayMinRemSP)

        for delaySecRem in delay.findall('DataValueMember/[@Name="Sec_Rem"]'):
            delaySecRemSP = int(delaySecRem.get('Value'))
            dfOperatorSP.loc[i, "Timer Remaining Seconds"] = str(delaySecRemSP)

        if sysType == 1:
            for delayTmrRem in delay.findall('DataValueMember/[@Name="TmrRemain"]'):
                delayTmrRemSP = int(delayTmrRem.get('Value'))
                dfOperatorSP.loc[i, "Remaining Timer"] = str(delayTmrRemSP)

        for timerPre in delay.findall('StructureMember/[@Name="Timer"]/DataValueMember/[@Name="PRE"]'):
            timerPreSP = int(timerPre.get('Value'))
            dfOperatorSP.loc[i, "Timer Preset"] = str(timerPreSP)

#####

        for suDelay in setpointTags.findall(suDelayPath):
            for delayIn in suDelay.findall('DataValueMember/[@Name="In"]'):
                delayInSP = str(delayIn.get('Value'))
                dfOperatorSP.loc[i, "SU Timer Delay In"] = delayInSP

            for delayOut in suDelay.findall('DataValueMember/[@Name="UnitsSelect"]'):
                delayOutSP = str(delayOut.get('Value'))
                dfOperatorSP.loc[i, "SU Timer Delay Units Select"] = delayOutSP

            for delayMinRem in suDelay.findall('DataValueMember/[@Name="Min_Rem"]'):
                delayMinRemSP = str(delayMinRem.get('Value'))
                dfOperatorSP.loc[i, "SU Timer Remaining Minutes"] = delayMinRemSP

            for delaySecRem in suDelay.findall('DataValueMember/[@Name="Sec_Rem"]'):
                delaySecRemSP = str(delaySecRem.get('Value'))
                dfOperatorSP.loc[i, "SU Timer Remaining Seconds"] = delaySecRemSP

            if sysType == 1:
                for delayTmrRem in suDelay.findall('DataValueMember/[@Name="TmrRemain"]'):
                    delayTmrRemSP = str(delayTmrRem.get('Value'))
                    dfOperatorSP.loc[i, "SU Remaining Timer"] = delayTmrRemSP

            for timerPre in suDelay.findall('StructureMember/[@Name="Timer"]'):
                for timerPreVal in timerPre.findall('DataValueMember/[@Name="PRE"]'):
                    timerPreSP = str(timerPreVal.get('Value'))
                    dfOperatorSP.loc[i, "SU Timer Preset"] = timerPreSP



    dfOperatorSP.loc[i, "Alm Tags"] = name

    i += 1

##### END #####


i = 0

# Find all "Tag" elements containing the attribute DataType="AlmCtrl"
for setpointTags in root.findall(operatorTagElement):
    # Get the value of the "Tag" attribute "Name"
    name = setpointTags.get('Name')
    dfOperandSP.loc[i, "Alm Tags"] = name

    i += 1



dfOperatorSP = dfOperatorSP.replace('\n', '', regex=True)
dfOperatorSP = dfOperatorSP.replace(np.nan, 'N/A', regex=True)
dfOperatorSP.to_csv('C:/Users/jzuccaro/Desktop/Py_Demo/almOperatorSP.csv')

dfOperandSP = dfOperandSP.replace('\n', '', regex=True)
dfOperandSP = dfOperandSP.replace(np.nan, 'N/A', regex=True)
dfOperandSP.to_csv('C:/Users/jzuccaro/Desktop/Py_Demo/almOperandSP.csv')


i = 0
n = dfOperatorSP.shape[0]

##### Set critical alarm values #####

while i < n:
    targetTag = str(dfOperatorSP["Alm Tags"].values[i])
###    targetTagStr = '"'+targetTag+'"'
    for setpointTags in root2.findall(operandTagElement):
        if setpointTags.get('Name') == targetTag:
            inVal = str(dfOperatorSP["In"].values[i])
            inPreVal = str(dfOperatorSP["Timer Preset"].values[i])
            delayInSP = str(dfOperatorSP["Timer Delay In"].values[i])
            delaySelSP = str(dfOperatorSP["Timer Delay Units Select"].values[i])
            delayMinRemSP = str(dfOperatorSP["Timer Remaining Minutes"].values[i])
            delaySecRemSP = str(dfOperatorSP["Timer Remaining Seconds"].values[i])
            suPreVal = str(dfOperatorSP["SU Timer Preset"].values[i])
            suDelayInSP = str(dfOperatorSP["SU Timer Delay In"].values[i])
            suDelaySelSP = str(dfOperatorSP["SU Timer Delay Units Select"].values[i])
            suDelayMinRemSP = str(dfOperatorSP["SU Timer Remaining Minutes"].values[i])
            suDelaySecRemSP = str(dfOperatorSP["SU Timer Remaining Seconds"].values[i])

            if sysType == 1:
                suDelayTmrRemSP = str(dfOperatorSP["SU Remaining Timer"].values[i])
                delayTmrRemSP = str(dfOperatorSP["Remaining Timer"].values[i])


            for almIn in setpointTags.findall(spPath):
                almIn.set('Value', inVal)

            for delay in setpointTags.findall(delayPath):
                for delayIn in delay.findall('DataValueMember/[@Name="In"]'):
                    delayIn.set('Value', delayInSP)

                for delayOut in delay.findall('DataValueMember/[@Name="UnitsSelect"]'):
                    delayOut.set('Value', delaySelSP)

                for delayMinRem in delay.findall('DataValueMember/[@Name="Min_Rem"]'):
                    delayMinRem.set('Value', delayMinRemSP)

                for delaySecRem in delay.findall('DataValueMember/[@Name="Sec_Rem"]'):
                    delaySecRem.set('Value', delaySecRemSP)

                if sysType == 1:
                    for delayTmrRem in delay.findall('DataValueMember/[@Name="TmrRemain"]'):
                        delayTmrRem.set('Value', delayTmrRemSP)

                for timerPre in delay.findall('StructureMember/[@Name="Timer"]/DataValueMember/[@Name="PRE"]'):
                    timerPre.set('Value', inPreVal)

            for suDelay in setpointTags.findall(suDelayPath):
                for suDelayIn in suDelay.findall('DataValueMember/[@Name="In"]'):
                    suDelayIn.set('Value', suDelayInSP)

                for suDelayOut in suDelay.findall('DataValueMember/[@Name="UnitsSelect"]'):
                    suDelayOut.set('Value', suDelaySelSP)

                for suDelayMinRem in suDelay.findall('DataValueMember/[@Name="Min_Rem"]'):
                    suDelayMinRem.set('Value', suDelayMinRemSP)

                for suDelaySecRem in suDelay.findall('DataValueMember/[@Name="Sec_Rem"]'):
                    suDelaySecRem.set('Value', suDelaySecRemSP)

                if sysType == 1:
                    for suDelayTmrRem in suDelay.findall('DataValueMember/[@Name="TmrRemain"]'):
                        suDelayTmrRem.set('Value', suDelayTmrRemSP)

                for timerPre in suDelay.findall('StructureMember/[@Name="Timer"]/DataValueMember/[@Name="PRE"]'):
                    timerPre.set('Value', suPreVal)

    i += 1

##### END #####


dfOperatorSP = dfOperatorSP.replace('\n', '', regex=True)
dfOperatorSP = dfOperatorSP.replace(np.nan, 'N/A', regex=True)
dfOperatorSP.to_csv('C:/Users/jzuccaro/Desktop/Py_Demo/almOperatorSP.csv')


ElementTree.tostring(root2, method='xml')
with open('C:/Users/jzuccaro/Desktop/Py_Demo/MECO10248_L5X_ALM_SP.L5X', 'wb') as f:
    f.write(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n');
    operandXML.write(f, xml_declaration=False, encoding='utf-8')

