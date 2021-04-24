To use:

This program searches all folders in current directory and subdirectories for .wav files and attempts to translate them with the google developers API. Details in developer API can be found in the speechrecognition module. 

Note on functionality: As of right now the files in directories searched must all be present in the Actual Subtitles CSV otherwise it will throw an error

Example audio files are included in the zip file. About half the files are ripped from movies and have poor quality audio, the other are clear dialogue samples.

See Comments.docx for more details on the following comments (regarding performance with test audio)
Weird issues (at 95% cuttoff):
• High confidence when but a mismatch in subtitles
	o Issue specifically with abrupt audio beginnings 
	o Adding some silence to the beginning helps
		 Increases translation accuracy
		 Modifies confidence rating to below cutoff
	o Also reversed logic in npy.where comparison
• Solution: 
	o Increase confidence threshold
		 Where 98.76% is highest recieved
	o Insert 150ms padding at beginning of troubled files to increase accuracy or decrease confidence
		 150ms is barely noticeable if inserted and sounds much more natural

Example: BreakMale.wav
•	Starts abruptly (sounds like part of the first word is missing)
•	Without Silence-Padding: Confidence 98.76%.
•	With Silence-Padding (150ms): Confidence 94.48%
•	Solution Padding corrected the translation and decreased confidence to below 95% threshold

Example: you-call-that-fun.wav 
•	Starts Very abruptly 
•	Without Silence-Padding: Confidence: 96.41%
•	With Silence-Padding (150ms): Confidence: 96.24%
•	Solution: increase confidence threshold to 98%

Example: bad-taste.wav
•	Noisy file altogether
•	Starts abruptly
•	Without Silence-Padding: Confidence: 98.76%
•	With Silence-Padding(150ms): Confidence 98.76%
•	Solution: When padding was added, translation became correct
