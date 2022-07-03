# 11May2022 - In this script's current state, it parses the project xml file to find the ScaledMax and InputMax
#             setpoints from the original program (10178) and saves them to a dataframe. The new program (10249)
#             is then parsed to determine the location of those setpoints and replaces the existing value with the
#             value found in the old program.

import xml.etree.ElementTree as ET
from xml.etree import ElementTree
import numpy as np
import pandas as pd

##### Enter project information #####

##### Set sysType to the number assigned to its system type
##### MP == 0
##### VC == 1
sysType = 1

operatorXML = ET.parse('C:/Users/jzuccaro/Desktop/Py_Demo/MECOStandard_L5X.L5X') # Use line 14860 for demo

operandXML = ET.parse('C:/Users/jzuccaro/Desktop/Py_Demo/MECO10248_V33_11_R00_1_Jun_2022.L5X')# Use line 16126 for demo

################


root = operatorXML.getroot()
root2 = operandXML.getroot()
dfOperatorSP = pd.DataFrame()
dfOperandSP = pd.DataFrame()

if sysType == 0:
    tagDataType = '"SclTag"'
    structMember = '"FAdO"'
elif sysType == 1:
    tagDataType = '"SclTag"'
    structMember = '"AdO"'

tagElement = "./Controller/Tags/Tag/[@DataType="+tagDataType+"]"

print(tagElement)
structElement = "Data/Structure/StructureMember/[@Name="+structMember+"]"


i = 0

# Find all "Tag" elements containing the attribute DataType="SclTag"
for setpointTags in root.findall(tagElement):
    # Get the value of the "Tag" attribute "Name"
    name = setpointTags.get('Name')
    dfOperatorSP.loc[i, "Tags"] = name

    # Get the text contained in the "Description" element
#    description = setpointTags.find('Description').text

    # Find the alarm message data contained in the "Comment" element, within the "Tag" element,
    # with the attribute Operand=".ALM.MSG"
    # <Comment Operand=".ALM.MSG">
    # <![CDATA[(CF-SANIT) CF Sanitization Overdue Alarm]]>
    # </Comment>
    for structMem1 in setpointTags.findall('Data/Structure/StructureMember/[@Name="AdO"]'):
        for inMax in structMem1.findall('DataValueMember/[@Name="InputMax"]'):
            inMaxSP = float(inMax.get('Value'))
            dfOperatorSP.loc[i, "In Max"] = str(inMaxSP)
        print(inMaxSP)
        for sclMax in structMem1.findall('DataValueMember/[@Name="ScaledMax"]'):
            sclMaxSP = float(sclMax.get('Value'))
            dfOperatorSP.loc[i, "Scaled Max"] = str(sclMaxSP)

        for enIn in structMem1.findall('DataValueMember/[@Name="EnableIn"]'):
            enInVal = int(enIn.get('Value'))
            dfOperatorSP.loc[i, "Enable In"] = str(enInVal)

        # for enOut in structMem1.findall('DataValueMember/[@Name="EnableOut"]'):
        #     enOutVal = int(enOut.get('Value'))
        #     dfOperatorSP.loc[i, "Enable Out"] = str(enOutVal)

    dfOperatorSP.loc[i, "Tags"] = name

    i += 1


i = 0

# Find all "Tag" elements containing the attribute DataType="AlmCtrl"
for setpointTags in root.findall(tagElement):
    # Get the value of the "Tag" attribute "Name"
    name = setpointTags.get('Name')
    dfOperandSP.loc[i, "Tags"] = name
    i += 1


dfOperatorSP = dfOperatorSP.replace('\n', '', regex=True)
dfOperatorSP = dfOperatorSP.replace(np.nan, 'N/A', regex=True)
dfOperatorSP.to_csv('C:/Users/jzuccaro/Desktop/analogInOperatorSP.csv')

dfOperandSP = dfOperandSP.replace('\n', '', regex=True)
dfOperandSP = dfOperandSP.replace(np.nan, 'N/A', regex=True)
dfOperandSP.to_csv('C:/Users/jzuccaro/Desktop/analogInOperandSP.csv')

i = 0

n = dfOperatorSP.shape[0]


while i < n:
    targetTag = str(dfOperatorSP["Tags"].values[i])
    targetTagStr = '"'+targetTag+'"'
    for setpointTags in root2.findall(tagElement):
        if setpointTags.get('Name') == targetTag:
            sclMaxVal = str(dfOperatorSP["Scaled Max"].values[i])
            sclMax1 = setpointTags.find('Data/Structure/DataValueMember/[@Name="ScaleMax"]')
            sclMax1.set('Value', sclMaxVal)
            sclMax2 = setpointTags.find('Data/Structure/StructureMember/DataValueMember/[@Name="ScaledMax"]')
            sclMax2.set('Value', sclMaxVal)

            inMaxVal = str(dfOperatorSP["In Max"].values[i])
            inMax1 = setpointTags.find('Data/Structure/StructureMember/DataValueMember/[@Name="InputMax"]')
            inMax1.set('Value', inMaxVal)
            inMax2 = setpointTags.find('Data/Structure/DataValueMember/[@Name="InMax"]')
            inMax2.set('Value', inMaxVal)

            enInVal = str(dfOperatorSP["Enable In"].values[i])
            enIn1 = setpointTags.find('Data/Structure/StructureMember/DataValueMember/[@Name="EnableIn"]')
            enIn1.set('Value', enInVal)

            # enOutVal = str(dfOperatorSP["Enable Out"].values[i])
            # enOut1 = setpointTags.find('Data/Structure/StructureMember/DataValueMember/[@Name="EnableOut"]')
            # enOut1.set('Value', enOutVal)

    i += 1


dfOperatorSP = dfOperatorSP.replace('\n', '', regex=True)
dfOperatorSP = dfOperatorSP.replace(np.nan, 'N/A', regex=True)
dfOperatorSP.to_csv('C:/Users/jzuccaro/Desktop/Py_Demo/analogInOperatorSP.csv')


ElementTree.tostring(root2, method='xml')
with open('C:/Users/jzuccaro/Desktop/Py_Demo/MECO10248_L5X_SP.L5X', 'wb') as f:
    f.write(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n');
    operandXML.write(f, xml_declaration=False, encoding='utf-8')

