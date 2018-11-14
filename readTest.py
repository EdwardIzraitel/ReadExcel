import xlrd

#opens excel sheet at sheet 2
workbook = xlrd.open_workbook("test.xlsx")
worksheet = workbook.sheet_by_index(1)

#finds total rols of excel sheet
rows = worksheet.nrows

material = list()
thic = list()
lDict = list()
strMaterial=""
strThic=""
i=0
z=0
keepGoing=True
listThic = list()
materialTrack=0


#loops for how many rows there are in the excel sheet
for i in range (rows):
    #if the material is not in the material goes into this
    if (str(worksheet.cell(i,0).value) not in material):
        #gets the material
        material.append(str(worksheet.cell(i,0).value))
        #creates a list and adds a tuple of thickness and whatever value to it
        y=list()
        y.append(((float(worksheet.cell(i,2).value)),float(worksheet.cell(i,3).value)))
        #appends it to the material dictionary
        lDict.append(y)
        #if the thickness is not in the thickness list goes into this
        if (float(worksheet.cell(i,1).value) not in thic):
            #Appends the thickness of the current material to the thickness list
            thic.append(float(worksheet.cell(i,1).value))
        #Adds to the string that prints the matieral
        strMaterial +=str(materialTrack)+ ":"+ str(worksheet.cell(i,0).value)+ "    "
        #Creates a temp list of thickness and adds it to the main list of thicknesses
        t=list()
        t.append(str(worksheet.cell(i,1).value))
        listThic.append(t)
        materialTrack+=1
    #if the material has been seen goes into this
    else:
        #gets the index of the current material in the material list
        index = material.index(str(worksheet.cell(i,0).value))
        #appends a tuple into the list of the material list
        lDict[index].append((float(worksheet.cell(i,2).value),float(worksheet.cell(i,3).value)))
        thic.append(float(worksheet.cell(i,1).value))
        listThic[index].append(str(worksheet.cell(i,1).value))
#print(listThic)
while(keepGoing):
    strThic=""
    print ("Type to select the following material")
    typeM=int(input(strMaterial+"\n"))
    for i in range(len(listThic[typeM])):
        strThic += str(i) + ":" + listThic[typeM][i]+ "    "
    print ("Type to select the following thickness")
    typeT = int(input(strThic+"\n"))
    typeV = float(input("Input the value\n"))

    m=lDict[typeM][typeT][0]
    b=lDict[typeM][typeT][1]
    print("Your value is: " +str(m*typeV+b))
    print("\n\n-----------------------------------------")

workbook= xlrd.close_workbook("test.xlsx")
