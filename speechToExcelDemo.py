import os
import csv
import json
import speech_recognition as sr
import pathlib
import pandas
import numpy as np
import xlsxwriter
import sys


totalWAVFiles = 0
r = sr.Recognizer()

#Get actual subtitle file from user
validInput = False
while validInput == False:
	inCurrentDir = input("Is subtitle file in current Directory? (Y/N)")
	if (inCurrentDir == "y") or (inCurrentDir == "Y"):
		subtitleInput = input("Enter csv file containing file names and subtitles: ")
		actualCSVFile = os.getcwd() + '\\' + subtitleInput
	elif (inCurrentDir == "n") or (inCurrentDir == "N"):
		actualCSVFile = input("Enter csv file path including csv file name: ")
	else:
		print ("Type Y or N")
	if (os.path.exists(actualCSVFile)) and (".csv" in actualCSVFile):
		validInput = True
	else:
		print("Enter Valid Path or File Name (must include .csv)")

#get path from user
validInput = False
while validInput == False:
	inCurrentDir = input("does current directory and/or subdirectories contain .wav files to transcribe (Y/N)")
	if (inCurrentDir == "y") or (inCurrentDir == "Y"):
		wavFolder = os.getcwd()
	elif (inCurrentDir == "n") or (inCurrentDir == "N"):
		wavFolder = input("Enter path containing .wav files")
	else:
		print ("Type Y or N")
	if (os.path.exists(wavFolder)):
		validInput = True
	else:
		print("Enter Valid Path")

#option to hardcode by defining wavFolder and actualCSVFile ,just comment out user input sections and uncomment below
# wavFolder = os.getcwd()
# actualCSVFile = wavFolder + r"\ActualSubtitles.csv"

expectedSubtitles = pandas.read_csv(actualCSVFile)
expectedSubtitles.sort_values('File Name', inplace=True)
expectedSubtitles.reset_index(inplace=True, drop=True)
#print(expectedSubtitles)
generatedSubtitleDataList = []
filesNotListed = []
for dirname, dirs, files in os.walk(wavFolder):
	for fileName in files:
		if (fileName.endswith(".wav")) and (expectedSubtitles['File Name'].str.contains(fileName).any()):
			print(fileName)
			totalWAVFiles += 1
			filePathInit = os.path.join(dirname, fileName)
			# get transcript from google speech-to-text API as a JSON
			with sr.AudioFile(filePathInit) as source:
				audio = r.record(source)
				jsonDict = r.recognize_google(audio, show_all=True)
				# making sure
				if (len(jsonDict) > 0) and ('confidence' in jsonDict['alternative'][0]):
					transcriptString = jsonDict['alternative'][0]['transcript']
					confidenceString = jsonDict['alternative'][0]['confidence']
				else:
					transcriptString = "Engine Could Not Transcribe"
					confidenceString = "0"
				print(transcriptString)
				print(confidenceString)
				workingList = [fileName, transcriptString,confidenceString, filePathInit]
				generatedSubtitleDataList.append(workingList)
		elif fileName.endswith(".wav"):
			filesNotListed.append(fileName)

generatedSubtitleData = pandas.DataFrame(generatedSubtitleDataList, columns=['File Name', 'Generated Subtitle', 'Confidence', 'File Path'])
generatedSubtitleData.sort_values('File Name', inplace=True)
generatedSubtitleData.reset_index(inplace=True, drop=True)

listOfMissingExpectedFiles = []
listOfPresentFileData = []

currentIndex = 0
for entry in expectedSubtitles['File Name']:
	if generatedSubtitleData['File Name'].str.contains(entry).any():
		presentFileData = [entry, expectedSubtitles.at[currentIndex,'Actual Subtitles']]
		listOfPresentFileData.append(presentFileData)
	else:
		listOfMissingExpectedFiles.append(entry)
	currentIndex += 1

checkedExpectedSubtitles = pandas.DataFrame(listOfPresentFileData, columns=['File Name', 'Actual Subtitles'])
# print(checkedExpectedSubtitles)
# print(generatedSubtitleData)

finalDataFrame = pandas.DataFrame()
finalDataFrame['Matching'] = np.where(generatedSubtitleData['Generated Subtitle'] == checkedExpectedSubtitles['Actual Subtitles'], True, False)
finalDataFrame['File Name'] = checkedExpectedSubtitles['File Name'].copy()
finalDataFrame["Expected Subtitle"] = checkedExpectedSubtitles["Actual Subtitles"].copy()
finalDataFrame['Generated Subtitle'] = generatedSubtitleData['Generated Subtitle'].copy()
finalDataFrame['Confidence'] = generatedSubtitleData['Confidence'].copy().astype(float)
finalDataFrame['Recommendation'] = np.where((finalDataFrame['Confidence'] < 0.98) | (finalDataFrame['Matching'] == "False"), "Review" , "-" )
finalDataFrame['File Path'] = generatedSubtitleData['File Path'].copy()

number_rows = len(finalDataFrame.index)

finalDataFrame = finalDataFrame[['Recommendation', 'File Name', 'Expected Subtitle', 'Generated Subtitle', 'Confidence', 'Matching', 'File Path']]
#could output to csv and be done here, with no formatting and no missing files

writer = pandas.ExcelWriter('ReviewThisFile.xlsx', engine='xlsxwriter')
finalDataFrame.to_excel(writer, index=False, sheet_name='Data')
workbook = writer.book
worksheet = writer.sheets['Data']

#pretty-fying the excel sheet format for readability
reviewFormat = workbook.add_format({'bg_color': '#BE2528',
									'font_color': '#FFFFFF'})

ignoreFormat = workbook.add_format({'bg_color': '#099726',
									'font_color': '#FFFFFF'})

centerFormat = workbook.add_format({'align': 'center'})

percentFormat = workbook.add_format({'num_format': '0.00%',
									 'align': 'center'})

worksheet.conditional_format('A1:A{}'.format(number_rows+1),{'type': 'cell',
															 'criteria': 'equal to',
															 'value': '"Review"',
															 'format': reviewFormat})

worksheet.conditional_format('A1:A{}'.format(number_rows+1),{'type': 'cell',
															 'criteria': 'equal to',
															 'value': '"-"',
															 'format': ignoreFormat})

worksheet.set_column('A1:A{}'.format(number_rows+1), 16, centerFormat)
worksheet.set_column('B1:B{}'.format(number_rows+1), 40)
worksheet.set_column('C1:D{}'.format(number_rows+1), 90)
worksheet.set_column('E2:E{}'.format(number_rows+1), 14, percentFormat)
worksheet.set_column('F1:F{}'.format(number_rows+1), 12)
worksheet.set_column('G1:G{}'.format(number_rows+1), 100)
worksheet.autofilter('A1:G{}'.format(number_rows+1))

worksheet2 = workbook.add_worksheet("Missing Files and Subtitles")
worksheet2.write(0,0,"Files Not Listed:")
worksheet2.write(0,1,"Files Expected but not Found")

for row_num, fileNames in enumerate(filesNotListed):
	worksheet2.write(row_num + 1, 0, fileNames)
worksheet2.set_column('A1:A{}'.format(len(filesNotListed)), 40)

for row_num, expectedFile in enumerate(listOfMissingExpectedFiles):
	worksheet2.write(row_num + 1, 1, expectedFile)
worksheet2.set_column('B1:B{}'.format(len(listOfMissingExpectedFiles)), 60)

writer.save()
print("Number of .WAV Files Read: " + str(totalWAVFiles))
print('Excel File Created')
