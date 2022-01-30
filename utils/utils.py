import csv
import os
import xml.etree.ElementTree as ET
import yaml


def createFolder(filepath):
    if(os.path.exists(filepath)):
        return
    os.mkdir(filepath)
    
def savetoRaml(datas, outputFileName):
    if(os.path.isfile(outputFileName)):
        os.remove(outputFileName)
    
    f = open(outputFileName, 'w', encoding='UTF-8')
    for(data) in datas:
        print(data)
        f.writelines(str(data) + '\n')
    f.close()

def getConfigInfo(configFileName):
    with open(configFileName) as file:
        configObj = yaml.safe_load(file)
    return configObj

def savetoCSV(fields, dataList, filename):
    # specifying the fields for csv file
    # fields = ['filename', 'field', 'readable', 'editable']
    # writing to csv file
    with open(filename, "a", newline="", encoding="utf_8_sig") as csvfile:
        # creating a csv dict writer object
        writer = csv.DictWriter(csvfile, fieldnames=fields)
        # writing headers (field names)
        # if isOutputHead : writer.writeheader()
        writer.writeheader()
        # writing data rows
        writer.writerows(dataList)

def savetoTxt(datas, outputFileName):
    f = open(outputFileName, 'w', encoding='UTF-8')
    for(key, value) in datas.items():
        # print(key, value)
        f.writelines(key + ': ' + value + '\n')
    f.close()

def parseXMLWithns(xmlfile, targetSubNotes, targetNote, sfdc_metadata):
    ns = {'sfdc-metadata': sfdc_metadata}

    # create element tree object
    tree = ET.parse(xmlfile)
    # get root element
    root = tree.getroot()
    # create empty list for items
    datas = []

    # iterate items
    isHavingtargetNote = False
    for item in root.findall('sfdc-metadata:' + targetNote, ns):
        isHavingtargetNote = True
        # empty employs dictionary
        notes = {'filename' : os.path.basename(xmlfile)}

        if(item.text.strip()):
            print(item.text)
            notes[targetNote] = item.text
        else:
            # iterate child elements of item
            for child in item:
                if(child.tag.split('}')[- 1] in targetSubNotes):
                    notes[child.tag.split('}')[- 1]] = child.text

        # append employs dictionary to list
        datas.append(notes)

    # print(datas)
    return datas

def parseXMLWithoutNs(xmlfile, targetSubNotes, targetNote):
    # create element tree object
    tree = ET.parse(xmlfile)
    # get root element
    root = tree.getroot()
    # create empty list for items
    datas = []

    # iterate items
    isHavingtargetNote = False
    for item in root.findall('./' + targetNote):
        isHavingtargetNote = True
        # empty employs dictionary
        notes = {'filename': os.path.basename(xmlfile)}
        # iterate child elements of item
        for child in item:
            if (child.tag in targetSubNotes):
                notes[child.tag] = child.text
        # append employs dictionary to list
        datas.append(notes)

    if(not isHavingtargetNote):
        notes = {'filename': os.path.basename(xmlfile)}
        datas.append(notes)

    # print('parseXMLWithoutNs ###################')
    # print(datas)
    # return list
    return datas

