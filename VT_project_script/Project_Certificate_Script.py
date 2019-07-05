
import os
import tempfile
import shutil
try:
	import pandas as pd
except ImportError:
	modules=['pandas' ,'xlrd']
	for mod in modules:
		os.system("pip install " + mod)
	import pandas as pd


def fixMonth(num):
	if len(num)==1:
		date = '0' + str(num)
	else:
		date = num
	return date

def fixYear(year):
	if len(year) ==2:
		yr = '20'+ year
	else:
		yr = year
	return yr

def fixDate(num):
	if len(num)==1:
		date = '0' + str(num)
	else:
		date = num
	return date



def validateCourse(crsName):
	if crsName == "":
		raise Exception("Course Name is Blank!")

	elif len(crsName) > 32:
		raise Exception("Course Name is too long for the course : " + crsName)

	else:
		return crsName

def validateCollege(clgName):
	
	if len(clgName) > 43:
		if "\\r" not in clgName:
			raise Exception("College Name is too long for : " + clgName + "\nEnter a '\\r' to Make the words in next line.")
		else:
			clgArr = clgName.split("\\r")
			return "\r".join(clgArr)

	else:
		return clgName



def handleExcel(excelFile):

	layerContents = []

	for key in excelFile:
		checkKey = key.lower().replace(" ","")
		arrItems = [] 
		if  checkKey == "startdate" or checkKey == "start" or checkKey == "date" :
			yr = []
			day = []
			mon =  []
			for startDate in excelFile[key]:
				if '.' in startDate:
					checker = startDate.split(".")
					if len(checker) == 3:
						dd , mm , yy = startDate.split(".")
						month = fixMonth(mm)
						year = fixYear(yy)
						date = fixDate(dd)

						day.append(date)
						mon.append(month)
						yr.append(year)
						
					else:
						raise ValueError("Invalid Date format detected !!! Expected '.' got something else instead : ", startDate, " in " , key)
			
				else:
					raise SyntaxError("Invalid Date format detected !!! Expected '.' got something else instead : " , startDate , " in " , key)
			layerContents.append(day)
			layerContents.append(mon)
			layerContents.append(yr)
		if checkKey == "enddate" or checkKey == "end" or checkKey == "date" :
			yr = []
			day = []
			mon =  []
			for startDate in excelFile[key]:
				if '.' in startDate:
					checker = startDate.split(".")
					if len(checker) == 3:
						dd , mm , yy = startDate.split(".")
						month = fixMonth(mm)
						year = fixYear(yy)
						date = fixDate(dd)

						day.append(date)
						mon.append(month)
						yr.append(year)
						
					else:
						raise ValueError("Invalid Date format detected !!! Expected '.' got something else instead : ", startDate, " in " , key)
			
				else:
					raise SyntaxError("Invalid Date format detected !!! Expected '.' got something else instead : " , startDate , " in " , key)
			layerContents.append(day)
			layerContents.append(mon)
			layerContents.append(yr)
		if checkKey == "studentname" or checkKey == "student" or checkKey == "name" :
			stdName = []
			for student in excelFile[key]:
				stdName.append(student)
						
			layerContents.append(stdName)

		if checkKey == "coursename" or checkKey == "course" or checkKey == "courses" :
			courseArray = []
			for course in excelFile[key]:
				courseArray.append(validateCourse(course))
			layerContents.append(courseArray)

		if checkKey == "collegename" or checkKey == "college" or checkKey == "colleges" :
			collegeArray = []
			for colege in excelFile[key]:
				collegeArray.append(validateCollege(colege))
			layerContents.append(collegeArray)

		if checkKey == "projectname" or checkKey == "project":
			projectArray = []
			for project in excelFile[key]:
				projectArray.append(project)
			layerContents.append(projectArray)
		
	return layerContents


def generateJS(layers, contents, psdPath, jpgPath, pathForJS='' , nameOfJS = 'psScript'):

	mycontent = '['
	for arr in contents:
		mycontent = mycontent + str(arr) +","+ "\n"

	mycontent = mycontent[:-2] + "]"

	js_code = '''var textLayers = ''' + str(layers) + ''';
	var values =  ''' + mycontent + ''';
	var pathPsd="'''+psdPath.replace("\\","\\\\")+'''".replace("\\\\","\\\\\\\\");
	var pathJpg ="'''+jpgPath.replace("\\","\\\\")+'''".replace("\\\\","\\\\\\\\");
	var fixedPath = pathJpg;
	var psdFile = new File(pathPsd);  
	var saveName = "";
	var layerName = "";
	var val = "";
	if(psdFile.exists){

	    app.open(psdFile);
	    function textLayrr (layers,changes) {

	    	var names = changes[0]
	    	var course = changes[1]
	    	var college = changes[2]

	    	var start_day = changes[3]
	    	var start_spr = changes[4]
	    	
	    	var start_mnth = changes[5]
	    	

			
	    	var end_day = changes[6]
	    	var end_spr = changes[7]
	    	var end_mnth = changes[8]

	    	var regID = changes[9]

	      for (i = 0 ; i< names.length ; i++){

	      	val = []

	      	val.push(names[i])
	      	val.push(course[i])
	      	val.push(college[i])

	      	val.push(start_day[i])
	      	val.push(start_spr[i])
	      	val.push(start_mnth[i])

	      	val.push(end_day[i])
	      	val.push(end_spr[i])
	      	val.push(end_mnth[i])

	      	val.push(regID[i])

	      	for (var j = 0; j < layers.length; j++) {
		      		app.activeDocument.layer = activeDocument.artLayers.getByName(layers[j]);
	                if(app.activeDocument.layer.visible == true){
	                    app.activeDocument.layer.textItem.contents= val[j];
				}
	      	}
	      	var jpgFile = new File(pathJpg+"\\\\"+val[0]+".jpg");

			jpgSaveOptions = new JPEGSaveOptions();
			jpgSaveOptions.quality = 12;

			if (jpgFile.exists){
				var dupFile = new File(pathJpg+"\\\\" + val[0] + i.toString() +".jpg");
				activeDocument.saveAs(dupFile, jpgSaveOptions, true, Extension.LOWERCASE);
				pathJpg = fixedPath
	      	}
	      	else{
	      		activeDocument.saveAs(jpgFile, jpgSaveOptions, true, Extension.LOWERCASE);
				pathJpg = fixedPath
	      	}

			
			var psdPath = pathJpg + "\\\\psd"
			var psdFile = new File(psdPath+"\\\\"+val[0]+".psd");
	      	if (psdFile.exists){
				var dupFile = new File(psdPath+"\\\\" + val[0] + i.toString() +".psd");
				activeDocument.saveAs(dupFile);
	      	}
	      	else{
	      		activeDocument.saveAs(psdFile);
	      	}
			
		}
		
		}

	textLayrr(textLayers,values)
	    
	app.activeDocument.close(SaveOptions.DONOTSAVECHANGES);  // SAVECHANGES to save file
	}
	else{
	    alert("Give full path");
	}'''

	nameOfJS = nameOfJS + ".jsx"

	fullPath = os.path.join(pathForJS, nameOfJS)

	try:
		jsFile= open(fullPath, "w")

		jsFile.write(js_code)

		jsFile.close()

		return True

	except Exception as e:

		return False
	return False

def checkEnvironment():
	
	eachPath = os.environ['PATH'].split(";")

	for path in eachPath:

		if "Adobe\\Adobe Photoshop" in path:
			return True

	return False

def validatePath(path):
	
	if path == "":
		return False

	else:
		if os.path.exists(path):
			return True
		elif os.path.isdir(path):
			return True
		else:
			return False
	return False


def main():

	files = input("\nPath of Excel File : ")

	psd_path = input("\nPath of PSD File : ")

	jpg_path = input("\nPath of JPG Files to be stored : ")

	if validatePath(files) and validatePath(psd_path) and validatePath(jpg_path):

		psdFilePath = os.path.join(jpg_path, "psd")

		if os.path.exists(psdFilePath):
			shutil.rmtree(psdFilePath, ignore_errors=True)
			os.makedirs(psdFilePath)
		else:
			os.makedirs(psdFilePath)

		data = ''

		try:
			data = pd.read_excel(files)

		except Exception as e:
			print("\nError in reading Excel!")
			print(e)

		global layerNames, layerContents

		print("\nCertificate Script is in default completion mode.")

		ch = input("Want to change any extra layer than default (y/n) ? : ").lower()

		if ch == 'y':

			print("\nThis feature is coming soon!")

		elif ch == 'n':

			#~~~~~~ These layer names must be present in PSD FILE

			layerNames = ['student_name' , 'technology_name' , 'clg_name','start_date','start_month','start_year','end_date','end_month','end_year','project_name']


			print("\n\n0. Generate only JSX File")
			print("1. Generate JSX File and Run using Photoshop")

			choice = input("\nEnter your choice : ")

			if choice == "0":

				multi_array_layer_contents = handleExcel(data)

				if generateJS(layerNames, multi_array_layer_contents, psd_path, jpg_path, pathForJS=tempfile.gettempdir()):

					shutil.copy(os.path.join(tempfile.gettempdir(),"psScript.jsx"), "psScript.jsx")
					print("\nJSX Generated! - 'psScript.jsx'")
				else:
					print("Error in Generating JSX file...")

			elif choice == "1":

				multi_array_layer_contents = handleExcel(data)

				if generateJS(layerNames, multi_array_layer_contents, psd_path, jpg_path, pathForJS=tempfile.gettempdir()):

					print("\nJSX Generated!")
					print("Running script...\n")

					if checkEnvironment():

						os.system("photoshop  --args " + os.path.join(tempfile.gettempdir(),"psScript.jsx"))

					else:
						print("\nPhotoshop is not in environment, set it in PATH variable first!")
						print("\nGenerating only JSX file...")
						shutil.copy(os.path.join(tempfile.gettempdir(),"psScript.jsx"), "psScript.jsx")
						print("\nDone! - 'psScript.jsx'\n")

				else:
					print("\nError in Generating script file!")

			else:
				print("\nBad Choice!")

		else:
			print("\nEnter a valid option!")

	else:
		print("\nError! Enter a proper PATH!")


if __name__ == '__main__':
	main()
