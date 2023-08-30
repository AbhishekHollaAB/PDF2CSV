import csv
import os
from pdf2docx import Converter
import docx2txt
import pandas as pd

try:
    os.mkdir('OUTPUT_FILES')
except:
    print('')
    pass


lenOutputs = len(os.listdir('OUTPUT_FILES'))
outXLName = 'Parámetros_' + str(lenOutputs + 1)

pdf_file  = 'INPUT/Summary.pdf'
docx_file = 'OUTPUT_FILES/SummaryRTF.rtf'

cv = Converter(pdf_file)
cv.convert(docx_file, start=0, end=None)
cv.close()

print('PDF to RTF Conversion Completed')
#################################################################
#################################################################
rtf_file_path = 'OUTPUT_FILES/SummaryRTF.rtf'
txt_file_path = 'OUTPUT_FILES/SummaryTEXT.txt'

text = docx2txt.process(rtf_file_path)
with open(txt_file_path, "w") as f:
    f.write(text)

print('RTF to TXT Conversion Completed')
#################################################################
def countOccurrences(string, word):
    # split the string by spaces in a
    a = string.split(" ")
    # search for pattern in a
    count = 0
    for i in range(0, len(a)):
        # if match found increase count
        if word == a[i]:
            count = count + 1
    return count
#################################################################
txtFile = open('OUTPUT_FILES/SummaryTEXT.txt', 'r')

lineList = []
allLines = txtFile.readlines()
for line in allLines:
    line = line.replace('\n', '')
    lineList.append(line)
txtFile.close()

idxSGDStart = lineList.index('Section Geometry Data')
idxSGDEnd = lineList.index('Wire Lengths in each Span')

with open(f'OUTPUT_FILES/{outXLName}.csv', 'w', newline = '') as csvFile:
    csvWriter = csv.writer(csvFile)
    csvWriter.writerow(['APOYO INICIAL', 'APOYO FINAL', 'LONGITUD VANO',
                        'FUNCIÓN  DEL APOYO INICIAL', 'FUNCIÓN DEL APOYO FINAL', 'TEMPERATURA INICIAL (° C)',
                        'TENSIÓN INICIAL(kg)', 'RESULTANTE INICIAL (kg/m)', 'DENOMINACIÓN DEL CABLE'])
    csvFile.close()

##SECTION TABLE##
sectionTable = pd.read_excel('INPUT/Section Table.xlsx', engine = 'openpyxl')
sagHorizonTension = list(sectionTable['Sag  Horiz.  Ten. (N)'])
outputTensionList = []
for eachVal in sagHorizonTension:
    val = float(eachVal) * 0.1019
    outputTensionList.append(val)
##SECTION TABLE##

##STAKING TABLE##
stakingTable = pd.read_excel('INPUT/Staking Table.xlsx', engine = 'openpyxl')

structureNameList = list(stakingTable['Structure  Name'])
finalStructureNameList = []
for name in structureNameList:
    name = str(name)
    if name != 'nan':
        splitName = name.split('\\')
        finalStructure = splitName[-1]
        finalStructure = finalStructure.split('.')[0]
        finalStructureNameList.append(finalStructure)

structureNumberList = list(stakingTable['Structure  Number'])
finalStructureNumberList = []
for number in structureNumberList:
    number = str(number)
    if number != 'nan':
        finalStructureNumberList.append(number)
##STAKING TABLE##

##OTHER 3 COLUMNS##
temperatureInitial = input('Enter Temperatura Inicial: ')
resultanteInitial = input('Enter Resultante Inicial: ')
denominacionCable = input('Enter Denominacion Del Cable: ')
##OTHER 3 COLUMNS##

# for i in range(idxSGDStart, idxSGDEnd):
#     print(lineList[i])
#     print('++++++++++++++++++++')

mainIdx = 0
for i in range(idxSGDStart, idxSGDEnd):    
    if '.wir' in lineList[i] or '.Wir' in lineList[i]:
        if countOccurrences(lineList[i], 'Circuit') > 1 :
            CircuitlineList = lineList[i].split('Circuit')
            print('----><><><><><><><><><><><><><>><><><><----')
            print(len(CircuitlineList))
            for j in range(0, len(CircuitlineList)):
                print(CircuitlineList[j])
                if '.wir' in CircuitlineList[j]:
                    updatedSplitText = []
                    splitText = CircuitlineList[j].split(' ')
                    for item in splitText:
                        if len(item) > 0:
                            updatedSplitText.append(item)
                            
                    print(updatedSplitText)
                    # print('11111111111111111111111111111111111')
                    try:
                        apoyoInitialIdx = finalStructureNumberList.index(updatedSplitText[4])
                        apoyoFinalIdx = finalStructureNumberList.index(updatedSplitText[5])
                        apoyoInitial = str(finalStructureNameList[apoyoInitialIdx])
                        apoyoFinal = str(finalStructureNameList[apoyoFinalIdx])
    
                        apoyoIntitial = apoyoInitial.lower()
                        apoyoFinal = apoyoFinal.lower()
    
                        if 'end' in apoyoInitial:
                            apoyoInitial = 'AMARRE'
                        elif 'suspclamp' in apoyoInitial or 'susp clamp' in apoyoInitial:
                            apoyoInitial = 'SUSPENSIÓN'
                        elif 'susppost' in apoyoInitial or 'susp post' in apoyoInitial or 'post' in apoyoInitial:
                            apoyoInitial = 'POSTE DE SUSPENSIÓN'
                        else:
                            apoyoInitial = 'NO SPANISH LOOKUP FOUND'
    
                        if 'end' in apoyoFinal:
                            apoyoFinal = 'AMARRE'
                        elif 'suspclamp' in apoyoFinal or 'susp clamp' in apoyoFinal:
                            apoyoFinal = 'SUSPENSIÓN'
                        elif 'susppost' in apoyoFinal or 'susp post' in apoyoFinal or 'post' in apoyoFinal:
                            apoyoFinal = 'POSTE DE SUSPENSIÓN'
                        else:
                            apoyoFinal = 'NO SPANISH LOOKUP FOUND'
                            
                        if len(updatedSplitText) < 12:
                            addedLineList = lineList[i + 8].split(' ')
                            for eachEle in addedLineList:
                                if eachEle != '':
                                    eachEle = eachEle.replace('\t', '')
                                    # print(eachEle, 'EACH ELEEE')
                                    updatedSplitText.append(eachEle)
                                    
    
                        # print(updatedSplitText[4])
                        if '.wir' in updatedSplitText[4] or '.Wir' in updatedSplitText[4]:
                            print('IN FOURRRRRRRRRR')
                            firstColumnApoyo = updatedSplitText[5]
                            secondColumnApoyo = updatedSplitText[6]
                            avgOfThree = (float(updatedSplitText[9]) + float(updatedSplitText[10]) + float(updatedSplitText[11]))/3
                        else:
                            print('HEREEEEEEEEEEEEE MAN')
                            firstColumnApoyo = updatedSplitText[4]
                            secondColumnApoyo = updatedSplitText[5]
                            avgOfThree = (float(updatedSplitText[8]) + float(updatedSplitText[9]) + float(updatedSplitText[10]))/3
                        outputElements = [firstColumnApoyo, secondColumnApoyo, str(avgOfThree),
                                          apoyoInitial, apoyoFinal, str(temperatureInitial),
                                          outputTensionList[mainIdx], str(resultanteInitial), str(denominacionCable)]
                        mainIdx += 1
    
                        with open(f'OUTPUT_FILES/{outXLName}.csv', 'a', newline = '') as csvFile:
                            csvWriter = csv.writer(csvFile)
                            csvWriter.writerow(outputElements)
                            csvFile.close()
                    except Exception as e:
                        print(e)
        else:
            updatedSplitText = []
            splitText = lineList[i].split(' ')
            for item in splitText:
                if len(item) > 0:
                    updatedSplitText.append(item)

            # print(updatedSplitText)
            # print('22222222222222222222222222222222222222222')
            
            try:
                apoyoInitialIdx = finalStructureNumberList.index(updatedSplitText[5])
                apoyoFinalIdx = finalStructureNumberList.index(updatedSplitText[6])
                apoyoInitial = str(finalStructureNameList[apoyoInitialIdx])
                apoyoFinal = str(finalStructureNameList[apoyoFinalIdx])
    
                apoyoIntitial = apoyoInitial.lower()
                apoyoFinal = apoyoFinal.lower()
    
                if 'end' in apoyoInitial:
                    apoyoInitial = 'AMARRE'
                elif 'suspclamp' in apoyoInitial or 'susp clamp' in apoyoInitial:
                    apoyoInitial = 'SUSPENSIÓN'
                elif 'susppost' in apoyoInitial or 'susp post' in apoyoInitial or 'post' in apoyoInitial:
                    apoyoInitial = 'POSTE DE SUSPENSIÓN'
                else:
                    apoyoInitial = 'NO SPANISH LOOKUP FOUND'
    
                if 'end' in apoyoFinal:
                    apoyoFinal = 'AMARRE'
                elif 'suspclamp' in apoyoFinal or 'susp clamp' in apoyoFinal:
                    apoyoFinal = 'SUSPENSIÓN'
                elif 'susppost' in apoyoFinal or 'susp post' in apoyoFinal or 'post' in apoyoFinal:
                    apoyoFinal = 'POSTE DE SUSPENSIÓN'
                else:
                    apoyoFinal = 'NO SPANISH LOOKUP FOUND'
                    
                if len(updatedSplitText) < 12:
                    addedLineList = lineList[i + 8].split(' ')
                    for eachEle in addedLineList:
                        if eachEle != '':
                            eachEle = eachEle.replace('\t', '')
                            # print(eachEle, 'EACH ELEEE')
                            updatedSplitText.append(eachEle)
                    # print(addedLineList, 'AAAAAAAAAAAAAAAAAAAAAA')
                    
    
                if '.wir' in updatedSplitText[4] or '.Wir' in updatedSplitText[4]:
                    # print('IN FOURRRRRRRRRR')
                    firstColumnApoyo = updatedSplitText[5]
                    secondColumnApoyo = updatedSplitText[6]
                    avgOfThree = (float(updatedSplitText[9]) + float(updatedSplitText[10]) + float(updatedSplitText[11]))/3
                else:
                    firstColumnApoyo = updatedSplitText[4]
                    secondColumnApoyo = updatedSplitText[5]
                    avgOfThree = (float(updatedSplitText[8]) + float(updatedSplitText[9]) + float(updatedSplitText[10]))/3
                outputElements = [firstColumnApoyo, secondColumnApoyo, str(avgOfThree),
                                  apoyoInitial, apoyoFinal, str(temperatureInitial),
                                  outputTensionList[mainIdx], str(resultanteInitial), str(denominacionCable)]
                mainIdx += 1
    
                with open(f'OUTPUT_FILES/{outXLName}.csv', 'a', newline = '') as csvFile:
                    csvWriter = csv.writer(csvFile)
                    csvWriter.writerow(outputElements)
                    csvFile.close()
            except:
                pass
print('TXT to CSV Conversion Completed')

df = pd.read_csv(f'OUTPUT_FILES/{outXLName}.csv', encoding='latin-1')
df.to_excel(f'OUTPUT_FILES/{outXLName}.xlsx', sheet_name = 'Hoja1', index = False, engine = 'openpyxl')

os.remove(f'OUTPUT_FILES/{outXLName}.csv')
# os.remove('OUTPUT_FILES/SummaryRTF.rtf')
# os.remove('OUTPUT_FILES/SummaryTEXT.txt')

print('All process completed. Outputs are stored in OUTPUT_FILES')
#################################################################
#################################################################