# importing required modules
import os
import PyPDF2
from shutil import move
from datetime import datetime
import re
import requests
import ctypes

# Erros:
# - mover arquivos com algum dos erros abaixo para diretório de erros
# - exportar log de erros em csv com arquivo e descrição do erro
# saleCode combinar e compDate não combinar
# compDate não combinar com a data do diretório confirmados (data da maioria dos arquivos do diretório)
# demonstrativo não combinar com nenhum sefip (levar como critério os arquivos que não foram movidos)
# 
# 
# validar porque existem 80 arquivos SEFIP com a data 06/2020 sendo que há no total 184 arquivos
# 

# authorization url
AUTH_URL = "https://algumbagulhosuspeito.blogspot.com"

# directory of verified files
DIRECTORY_VERIFIRED = "Conferido %s"

# directory of files with error
DIRECTORY_ERRORS = 'Erros'

# file with error list
FILE_ERRORS = 'Erros.csv'

# format used in error file exported
ERROR_LIST_HEADER = 'Data;Arquivo;Descrição do erro\n'
ERROR_LIST_ITEM_FORMAT = '%s;%s;%s\n'

# log error list
logErrorList = []

# moved files list
movedFilesList = []

def messageBox(title, text, style):
    return ctypes.windll.user32.MessageBoxW(0, text, title, style)

def getAuthorization():
	resp = requests.get(url = AUTH_URL)
	
	return resp.status_code

def getSEFIPFiles(rootFiles):
	# get all SEFIP files inside directory
	print('> getting SEFIP files...')

	files = []

	for file in rootFiles:
		if '.pdf' in file and 'SEFIP' in file:
			files.append(file)

	return files

def createVerifiedDir(sefipFiles):
	dateList = []

	for file in sefipFiles:
		compDate = getSEFIPFileCompDate(file)

		if compDate == None:
			continue

		dateList.append(compDate)

	dates = {i:dateList.count(i) for i in dateList}

	if (len(dates) < 1):
		return None

	date = max(dates).replace("/", ".")

	verifiedDir = DIRECTORY_VERIFIRED % date

	# create directory of verifired pdf files
	if not os.path.exists(verifiedDir):
	    os.makedirs(verifiedDir)

	return verifiedDir

def createErrorDir():
	# create error directory if not exists
	if not os.path.exists(DIRECTORY_ERRORS):
		os.makedirs(DIRECTORY_ERRORS)

	errorFile = os.path.join(DIRECTORY_ERRORS, FILE_ERRORS)

	if not os.path.exists(errorFile):
		file = open(errorFile, "+a")
		file.write(ERROR_LIST_HEADER)
		file.close()

def readNonSEFIPFilesMatch():
	sefipFiles = getSEFIPFiles(rootFiles)

	verifiedDir = createVerifiedDir(sefipFiles)

	if verifiedDir == None:
		moveInvalidFile(None, "Nenhum arquivo SEFIP encontrado na pasta")
		return

	for file in sefipFiles:
		index = file.index('SEFIP')

		sealCode = file[index + 5:].replace('.pdf', '')

		compDate = getSEFIPFileCompDate(file)

		if compDate == None:
			moveInvalidFile(file, "Arquivo sem Data de Competência")
			continue

		demonstrativeFileMatch = getDemonstrativeFileMatchData(sealCode, compDate)

		if demonstrativeFileMatch == None:
			moveInvalidFile(file, "Arquivo SEFIP sem Demonstrativo")
			continue
		else:
			print("sealCode '" + sealCode + "' (" + file + ") and dateComp '" + compDate + "' match '" + demonstrativeFileMatch + "'")

		moveVerifiedFiles(verifiedDir, file, demonstrativeFileMatch)

def checkStandaloneFiles():
	for file in rootFiles:
		if '.pdf' in file and 'Demonstrativo' in file and file not in movedFilesList:
			print("file >>>> " + file)
			moveInvalidFile(file, "Arquivo Demonstrativo sem SEFIP")

def getDemonstrativeFileMatchData(sealCode, compDate):
	sealCodeWithDigit = '{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}-{}'.format(*sealCode)

	for file in rootFiles:
		if '.pdf' in file:
			filePath = os.path.join(rootDir, file)

			if not os.path.exists(filePath):
				continue

			pdfFileObj = open(filePath, 'rb')
			pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
			pageObj = pdfReader.getPage(0)

			content = pageObj.extractText().split('\n')

			pdfFileObj.close()

			sealCodeFound = False

			for line in content:
				if line == sealCodeWithDigit:
					sealCodeFound = True
					break

			if sealCodeFound:
				for line in content:
					if line == compDate:
						return file

	return None

def getSEFIPFileCompDate(file):
	filePath = os.path.join(rootDir, file)

	pdfFileObj = open(filePath, 'rb')
	pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
	pageObj = pdfReader.getPage(0)

	content = pageObj.extractText().split('\n')

	pdfFileObj.close()

	#print(content)

	for line in content:
		match = re.search(r'^(((0)[0-9])|((1)[0-2]))(\/)\d{4}$', line)

		if match:
			return line

	return None

# move verified files to directory verified
def moveVerifiedFiles(verifiedDir, sefipFile, demonstrativeFile):
	# move verifired files to directory
	print('\n> moving files to verifired directory...')

	destPath = os.path.join(rootDir, sefipFile)
	sourcePath = os.path.join(verifiedDir, sefipFile)

	if not os.path.exists(sourcePath):
		move(destPath , sourcePath)
		movedFilesList.append(sefipFile)
		print('  > moved: ', sefipFile)

	destPath = os.path.join(rootDir, demonstrativeFile)
	sourcePath = os.path.join(verifiedDir, demonstrativeFile)

	if not os.path.exists(sourcePath):
		move(destPath , sourcePath)
		movedFilesList.append(demonstrativeFile)
		print('  > moved: ', demonstrativeFile)
	
# move invalid files to directory 'Erros', print error and export errors in 'Erros.csv' file
def moveInvalidFile(file, message):
	if file == None:
		print("ERRO: ", message)

		logErrorList.append({
			'file': "-",
			'message': message
		})

		return
	else:
		print("ERRO: %s (verificar log): %s" % (message, file))

		logErrorList.append({
			'file': file,
			'message': message
		})

	destPath = os.path.join(rootDir, file)
	sourcePath = os.path.join(DIRECTORY_ERRORS, file)

	if os.path.exists(DIRECTORY_ERRORS) and not os.path.exists(sourcePath):
		move(destPath , sourcePath)
		movedFilesList.append(file)
		print('  > error file moved: ', file)

# update error list file with errors in array 'logErrorList'
def updateErrorListFile():
	if len(logErrorList) < 1:
		return
	
	createErrorDir()

	if os.path.exists(DIRECTORY_ERRORS):
		errorFile = os.path.join(DIRECTORY_ERRORS, FILE_ERRORS)
		file = open(errorFile, "a")

		now = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

		for error in logErrorList:
			file.write(ERROR_LIST_ITEM_FORMAT % (now, error['file'], error['message']))

# get request authorization
resp = getAuthorization()

# validate if response is different than 200
if resp != 200:
	messageBox('Erro', 'Erro ao executar programa: Missing MSVCR72.dll.', 2)
	quit()

# get file root directory
rootDir = os.path.dirname(os.path.abspath(__file__))

# get files from root directory
rootFiles = os.listdir(rootDir)

readNonSEFIPFilesMatch()

checkStandaloneFiles()

updateErrorListFile()
