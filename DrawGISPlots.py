import arcpy
import xlrd
import array
from array import *
counterSHP = 0
counterXLS = 0
check = 0 
list = []
list2 = []

A = arcpy.GetParameterAsText(0) # Import 
arcpy.AddMessage("Params")
arcpy.AddMessage(A)

outputFolderPath = arcpy.GetParameterAsText(2)
arcpy.AddMessage("outputFolderPath ")
arcpy.AddMessage(outputFolderPath)

anyValue = arcpy.GetParameterAsText(3)
arcpy.AddMessage("Any Value ")
arcpy.AddMessage(anyValue)

myList  = [arcpy.GetParameterAsText(0)]
arcpy.AddMessage("params length")
arcpy.env.workspace = "F:\Survey\Nainital\March Visit\RiskEvaluation\ShapeFilesAndExcels_All"
arcpy.AddMessage(arcpy.ListFiles("*.xls")[0])	   # returns the name of the .xls file at the zeroeth location 	
#arcpy.AddMessage(arcpy.ListFiles("*.shp")[0])      # returns the name of the .shp file at the zeroeth location
#arcpy.AddMessage(len(arcpy.ListFiles("*.shp")[0])) # returns the length of the name of the file at the zeroeth location
#arcpy.AddMessage(arcpy.ListFiles("*.shp")[0][:2])  # returns the first 2 characters of the name of the file
for i in arcpy.ListFiles("*.xls"): 
	#a= array(i)
	counterXLS = counterXLS + 1 
	arcpy.AddMessage("IN FOR")
	arcpy.AddMessage(i)
	book = xlrd.open_workbook(i)
	sheet = book.sheet_by_index(0) # Get the first sheet
	arcpy.AddMessage(counterSHP)
	arcpy.AddMessage("Value in Cellllllllllllllllllllllllllllllllllllllllllllllllllllllll 15,7 is ::::")
	list.append([sheet.cell(15,7)])
	arcpy.AddMessage (sheet.cell(15,7).value) 
	#arcpy.AddMessage(arcpy.ListFiles("*.shp")[counterXLS-1])
	arcpy.AddMessage ("final OutPut Path454")
	#arcpy.AddMessage (outputFolderPath+str(sheet.cell(17,7).value)+arcpy.ListFiles("*.shp")[counterXLS-1])
	#arcpy.CopyFeatures_management(arcpy.ListFiles("*.shp")[counterXLS-1], outputFolderPath+"_"+anyValue+"_"+str(int(sheet.cell(15,7).value)))
	#arcpy.Merge_management(["majorrds.shp", "Habitat_Analysis.gdb/futrds"], "C:/output/Output.gdb/allroads")

#arcpy.AddMessage("list len")
#arcpy.AddMessage(len(list))
for j in arcpy.ListFiles("*.shp"):
	counterSHP = counterSHP + 1

#arcpy.AddMessage("counterSHP value is ")
#arcpy.AddMessage(counterSHP)
#arcpy.AddMessage ("(counterXLS  value is ")
#arcpy.AddMessage(counterXLS )

arcpy.AddMessage ("lenLIST") 
arcpy.AddMessage (len(list))
upto = 0
for count in range(0,len(list)):
	arcpy.AddMessage("count b4")
	arcpy.AddMessage(count)
	arcpy.AddMessage("LIST")
	arcpy.AddMessage(list[count])
	arcpy.AddMessage("considering INT")
	arcpy.AddMessage(list[count][9:10])
	arcpy.AddMessage("STRING")
	arcpy.AddMessage(str(list[count])[:2])
	if len(str(list[count]))== 12:
		upto = 9
		arcpy.AddMessage("in 12")
		arcpy.AddMessage(str(list[count])[8:upto])
		list2.append(str(list[count])[8:upto])
	elif len(str(list[count]))== 13:
		upto = 10
		arcpy.AddMessage("in 13")
		arcpy.AddMessage(str(list[count])[8:upto])
		list2.append(str(list[count])[8:upto])
	elif len(str(list[count]))== 14:
		upto = 11
		arcpy.AddMessage("in 14")
		arcpy.AddMessage(str(list[count])[8:upto])
		list2.append(str(list[count])[8:upto])
	elif len(str(list[count]))== 15:
		upto = 12
		arcpy.AddMessage("in 15")
		arcpy.AddMessage(str(list[count])[8:upto])
		list2.append(str(list[count])[8:upto])
	elif len(str(list[count]))== 16:
		upto = 13
		arcpy.AddMessage("in 16")
		arcpy.AddMessage(str(list[count])[8:upto])
		list2.append(str(list[count])[8:upto])
b=0
for seeValues in range(0, len(list2)):
	arcpy.AddMessage("SEE VALUES")
	arcpy.AddMessage(arcpy.ListFiles("*.shp")[seeValues][:2])
	arcpy.AddMessage(list2[seeValues])


for countList2 in range(0,len(list2)):
	arcpy.AddMessage("DUP FOR")
	arcpy.AddMessage(arcpy.ListFiles("*.shp")[countList2][:2])
	b=countList2+1
	for count1List2 in range(b,len(list2)):
		#arcpy.AddMessage("countList2")
		#arcpy.AddMessage(countList2)
		#arcpy.AddMessage("count1List2")
		#arcpy.AddMessage(count1List2)
		if list2[countList2]  == list2[count1List2]:
			arcpy.AddMessage("DUP WITH")
			arcpy.AddMessage(arcpy.ListFiles("*.shp")[count1List2][:2])

	arcpy.AddMessage("---END---")
			
	
######## for removing duplicates in 
afterRemovingDups = []
for i in list2:
       if i not in afterRemovingDups:
          afterRemovingDups.append(i)


arcpy.AddMessage("Actual List")
arcpy.AddMessage(list)


arcpy.AddMessage("With DUPS")
arcpy.AddMessage(list2)

arcpy.AddMessage("afterRemovingDups")
arcpy.AddMessage(afterRemovingDups)

m = 0 # for length of templist also for increment in the value of arcpy.ListFiles("*.shp")[j][:2]if 2 or more clusters with common values are found
#n  = 0 # 
matchCounter = 0 
tempList = []
for i in range(0,len(afterRemovingDups)):
	matchCounter = 0 
	tempListIndex = []
	tempListName = []
	tempListGroupNameString = ""
	
	for j in range(0,len(list2)):
		if afterRemovingDups[i] == list2[j]:
			arcpy.AddMessage("afterRemovingDups[i]")
			arcpy.AddMessage(afterRemovingDups[i])
			arcpy.AddMessage("list2[j]")
			arcpy.AddMessage(list2[j])
			tempListIndex.append(j)
			arcpy.AddMessage("len(tempListIndex)")
			arcpy.AddMessage(len(tempListIndex))
			tempListName.append(arcpy.ListFiles("*.shp")[j])
			tempListGroupNameString = tempListGroupNameString + "_"+ arcpy.ListFiles("*.shp")[j][:2]

	arcpy.AddMessage("Exit LOOP")
	arcpy.AddMessage("Templist INDEX")
	arcpy.AddMessage(tempListIndex)
	arcpy.AddMessage("Templist Name")
	arcpy.AddMessage(tempListName)
	arcpy.AddMessage("CHECKITOUTTTTTTTTTTTTTTTTT")
	arcpy.AddMessage('"'+';'.join(tempListName)+'"')
	arcpy.AddMessage("VALAUE")
	arcpy.AddMessage(list2[i])
	if len(tempListIndex)==1:
		arcpy.AddMessage("only one value")
		arcpy.CopyFeatures_management(arcpy.ListFiles("*.shp")[tempListIndex[0]], outputFolderPath+"_"+anyValue+"_"+tempListGroupNameString+"_"+afterRemovingDups[i])
	elif len(tempListIndex)>1:
		m =0
		m = len(tempListIndex)
		
		arcpy.AddMessage("> one value")
		arcpy.Merge_management('"'+';'.join(tempListName)+'"', outputFolderPath+"_"+anyValue+"_"+tempListGroupNameString+"_"+afterRemovingDups[i])
		arcpy.AddMessage("After Merge")



	


		