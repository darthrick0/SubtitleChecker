import os
import csv
import json
import speech_recognition as sr
import pathlib
import pandas
import numpy as np
import xlsxwriter

APP_FOLDER = os.getcwd()
totalWAVFiles = 0
r = sr.Recognizer()
print(APP_FOLDER)
expectedSubtitles = pandas.read_csv(APP_FOLDER + r"\ActualSubtitles.csv")
expectedSubtitles.sort_values('File Name', inplace=True)
expectedSubtitles.reset_index(inplace=True, drop=True)

generatedSubtitleDataList = []

for dirname, dirs, files in os.walk(APP_FOLDER):
	for fileName in files:
		if fileName.endswith(".wav"):
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

print(expectedSubtitles)
generatedSubtitleData = pandas.DataFrame(generatedSubtitleDataList, columns=['File Name', 'Generated Subtitle', 'Confidence', 'File Path'])
generatedSubtitleData.sort_values('File Name', inplace=True)
generatedSubtitleData.reset_index(inplace=True, drop=True)
print(generatedSubtitleData)

finalDataFrame = pandas.DataFrame()

finalDataFrame['Matching'] = np.where(generatedSubtitleData['Generated Subtitle'] == expectedSubtitles['Actual Subtitles'], True, False)
finalDataFrame['File Name'] = expectedSubtitles['File Name'].copy()
finalDataFrame["Expected Subtitle"] = expectedSubtitles["Actual Subtitles"].copy()
finalDataFrame['Generated Subtitle'] = generatedSubtitleData['Generated Subtitle'].copy()
finalDataFrame['Confidence'] = generatedSubtitleData['Confidence'].copy().astype(float)
finalDataFrame['Recommendation'] = np.where((finalDataFrame['Confidence'] < 0.98) | (finalDataFrame['Matching'] == "False"), "Review" , "-" )
finalDataFrame['File Path'] = generatedSubtitleData['File Path'].copy()

number_rows = len(finalDataFrame.index)

finalDataFrame = finalDataFrame[['Recommendation', 'File Name', 'Expected Subtitle', 'Generated Subtitle', 'Confidence', 'Matching', 'File Path']]
#!!!could output to csv and be done here!!!
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
worksheet.set_column('E2:E{}'.format(number_rows+1), 14,percentFormat)
worksheet.set_column('F1:F{}'.format(number_rows+1),12)
worksheet.set_column('G1:G{}'.format(number_rows+1),100)
worksheet.autofilter('A1:G{}'.format(number_rows+1))

writer.save()
print("Number of .WAV Files Read: " + str(totalWAVFiles))
print('Excel File Created!!!')
