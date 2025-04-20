'''
advisorEXtracTOR
AdvisorTrac Report Extractor (version 3)

Author: MothsInTheMachine on GitHub
Last Modified: Dec. 17, 2024
'''

import sys
import os, os.path
from datetime import datetime
from process_files import csv_process, excel_process


# Define MAIN
def main():
	labLabels = ['DC Math Lab','DC Writing Lab','SV Math Lab','SV Writing Lab']
	keywords = defineKeywords(labLabels)
	inputPath,outputPath = verifyArguments(len(sys.argv))

	printHeader()
	
	if not (verifyOverwrite(outputPath)):
		print("Cancelling process. Please run command again later.")
		printFooter()
		exit()

	
	processFiles(inputPath, outputPath, keywords, labLabels, [])
	printFooter()
	exit()
# End MAIN


# FUNCTION DEFINITIONS

def printHeader():
	print()
	print('/*******************************************/')
	print('/* AdvisorTrac Report Extractor in process */')
	print('/*******************************************/')
	print()


def printFooter():
	print()
	print('/*******************************************/')
	print('/*  Qutting AdvisorTrac Report Extractor.  */')
	print('/*******************************************/')
	print()


def quitProcess(reason='unknown'):
	print(f'Sorry, I could not understand that command. The reason is {reason}.')
	# print('Syntax:\n\n\t> python at-report-extract.py [optional to input directory] [optional output file name]\n')
	printFooter()
	quit()



def defineKeywords(_labLabels=[]):
	keywords = []
	if (len(_labLabels) > 0):
		keywords.extend(_labLabels)
	keywords.extend(['DC Math Lab Total:','DC Writing Lab Total:','SV Math Lab Total:','SV Writing Lab Total:','Grand Total:'])
	keywords.extend(['Chemistry Total:','CIS/Programming Total:','Math Total:','Other Total:','Other Course Tutoring Total:','Physics Total:'])
	keywords.extend(['Languages Total:','MS Office Total:','Reading Total:','Social Sciences Total:','Workshop Total:','Writing Total:'])
	return keywords


def verifyArguments(n):
	inputPath  = '.\\input'
	outputPath = '.\\output.csv'

	# Verify input and output paths
	if (n == 1):
		# Use default paths for processing
		pass
	
	if (n == 2  or  n == 3):
		# Use a different input path, but default output path
		inputPath = sys.argv[1]
		if not os.path.exists(inputPath):
			quitProcess(f'the extractor cannot find \"{inputPath}\"')
	
	if (n == 3):
		# Use a different input and output path
		outputPath = sys.argv[2]
		# Edit the formatting of the output path to CSV files
		pass
	
	if (n > 3):
		# Too many arguments. Exit extractor
		quitProcess(f'there are too many arguments to use')

	return inputPath,outputPath


def verifyOverwrite(outputPath):
	if (os.path.exists(outputPath)):
		# Verify overwriting of existing output file
		
		choice = ''
		while (choice != 'n') and (choice != 'y'):
			print(f'The output file \"{outputPath}\" already exists.')
			inputStr = input('Do you want to overwrite it? [y/n]: ')
			choice = inputStr[0].lower()
			# Gets to this line if invalid choice is made
			if (choice != 'n' and choice != 'y'):
				print("\nInvalid choice. Please enter Y for 'Yes' and N for 'No'...")
		
		if (choice == 'n'):
			return False
		
		if (choice == 'y'):
			print(f'\nConfirmed overwriting previous file \"{outputPath}\".')
			return True
		
	else:
		# Path does not exists
		return True

def processFiles():
	while (output_file_choice != 'c') and (output_file_choice != 'x'):
		output_file_choice=input('File output: type c for csv file or x for excel file')
		if output_file_choice.lower().strip()=='c':
			csv_process(inputPath,outputPath,keywords,labLabels,totals)
		elif output_file_choice.lower().strip()=='x':
			excel_process()
		else:
			print("\nInvalid choice. Please enter c for 'csv' and x for 'excel'...")
		
		

# Invoke MAIN
main()