import openpyxl
import os

from utils import utils

excel_path='C:/temp/CreateRamlTemplate-up.xlsx'
workbook = openpyxl.load_workbook(excel_path)

# シートのロード
sheet = workbook['data']

def execute():
    settingInfo = getSettingInfo()
    # print(settingInfo)
    # print('---------requestInfos----------')
    requestInfos = getDataInfo(settingInfo['reqStartLine'], settingInfo['reqEndLine'])
    # print(requestInfos)

    # print('---------responseInfos----------')
    responseInfos = getDataInfo(settingInfo['resStartLine'], settingInfo['resEndLine'])
    # print(responseInfos)

    reqDataType = getDataTypeRaml(requestInfos)
    resDataType = getDataTypeRaml(responseInfos)
    reqExample = getExapmleRaml(requestInfos)
    resExample = getExapmleRaml(responseInfos)

    outputFile(settingInfo, reqDataType, resDataType, reqExample, resExample)

def getSpace(level):
    spaceText = ''
    for i in range((level-1)*2):
        spaceText += ' '
    return spaceText

def getDataTypeRaml(dataInfos):
    dataTypeList = []
    dataTypeList.append('#%RAML 1.0 DataType')
    dataTypeList.append('')
    dataTypeList.append('uses:')
    dataTypeList.append('  xxx: xxx.raml')
    dataTypeList.append('')
    dataTypeList.append('properties:')

    propertiesSpace = ''
    for(data) in dataInfos:
        if(data['apiName'] == '-'):
            continue

        # API Name
        level = getSpace(data['level']) + propertiesSpace
        # print(str(data['level']) + ':[' + level + ']')
        dataTypeList.append(level + data['apiName'] + ':')
        # description: "SFCC顧客ナンバー"
        if(data['name'] != '-'):
            dataTypeList.append(level + '  description: "' + data['name'] + '"')
        # type: AccountBasic.customer_no
        typeText = 'xxx.xxx'
        if(data['type'] == 'object'):
            typeText = 'object'
        dataTypeList.append(level + '  type: ' + typeText)
        # required: true
        isMust = data['isMust'] == '必須'
        dataTypeList.append(level + '  required: ' + str(isMust).lower())

        propertiesText = ''
        if(data['type'] == 'object'):
            propertiesText = level + '  properties:'
            propertiesSpace = '  '
        dataTypeList.append(propertiesText)

    return dataTypeList

def getExapmleRaml(dataInfos):
    exapmleList = []
    exapmleList.append('#%RAML 1.0 NamedExample')
    exapmleList.append('value:')

    for(data) in dataInfos:
        if(data['apiName'] == '-'):
            exapmleList.append('-')
            continue

        # API Name
        level = getSpace(data['level'])
        # print(str(data['level']) + ':[' + level + ']')
        valueText = ' ""'
        if(data['type'] == 'object') or (data['type'] == 'number'):
            valueText = ''
        exapmleList.append(level + data['apiName'] + ':' + valueText)
    
    exapmleList.append('')
    return exapmleList

    
def outputFile(setting, reqDataType, resDataType, reqExample, resExample):
    rootPath = 'C:/temp'
    
    # Create Folder
    baseFolder = rootPath + '/' + setting['folderName']
    utils.createFolder(baseFolder)

    # data-types
    dataTypesFolder = baseFolder+ '/' + setting['subTypeFolderName']
    utils.createFolder(dataTypesFolder)
    # Request
    reqDataTypeFileName = dataTypesFolder + '/' + setting['fileName'] + setting['fileExt']
    utils.savetoRaml(reqDataType, reqDataTypeFileName)
    # Response
    resDataTypeFileName = dataTypesFolder + '/' + setting['fileName'] + setting['dataTypeFileRes'] + setting['fileExt']
    utils.savetoRaml(resDataType, resDataTypeFileName)

    # examples
    examplesFolder = baseFolder+ '/' + setting['subExpFolderName']
    utils.createFolder(examplesFolder)
    # Request
    reqExampleFileName = examplesFolder + '/' + setting['fileName'] + setting['dataExpFileReq'] + setting['fileExt']
    utils.savetoRaml(reqExample, reqExampleFileName)
    # Response
    resExampleFileName = examplesFolder + '/' + setting['fileName'] + setting['dataExpFileRes'] + setting['fileExt']
    utils.savetoRaml(resExample, resExampleFileName)
    
def getDataInfo(startLine, endLine):
    dataInfos = []

    for row in sheet.iter_rows(min_row=startLine, max_row=endLine):
        dataInfo = {}
        dataInfo['level'] = row[4].value
        dataInfo['apiName'] = row[7].value.strip()
        dataInfo['name'] = row[13].value.strip()
        dataInfo['type'] = row[23].value.strip()
        dataInfo['isMust'] = row[27].value.strip()
        dataInfos.append(dataInfo)

    return dataInfos

def getSettingInfo():
    setting = {}

    setting['folderName'] = sheet['B1'].value
    setting['subTypeFolderName'] = sheet['B2'].value
    setting['subExpFolderName'] = sheet['B3'].value

    setting['fileName'] = sheet['B4'].value
    setting['fileExt'] = sheet['B5'].value

    setting['dataTypeFileReq'] = sheet['B6'].value
    setting['dataTypeFileRes'] = sheet['B7'].value
    setting['dataExpFileReq'] = sheet['B8'].value
    setting['dataExpFileRes'] = sheet['B9'].value

    setting['reqStartLine'] = sheet['B10'].value
    setting['reqEndLine'] = sheet['B11'].value
    setting['resStartLine'] = sheet['B12'].value
    setting['resEndLine'] = sheet['B13'].value
    return setting

execute()
