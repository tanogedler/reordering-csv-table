import csv
import xlrd 

# We import the file where the categories are
# we save the excel in the xcl variable to refer to when we go to access the sheets
xcl = xlrd.open_workbook('./originalData/categories.xlsx') 

# We keep each Excel sheet by categories
# In category i is the sheet xcl.sheet_by_index (i) 
category1 = xcl.sheet_by_index(1) 
category2 = xcl.sheet_by_index(2) 
category3 = xcl.sheet_by_index(3) 
 
############################
# Now we are going to place each matrix of the categories in 
# two-dimensional arrays. The names of the arrangements indicate 
# the order in which they are presented in the original excel. 
#For example the first matrix of category 1 is called Array1Category1
###########################
## category 1


245/5000
# We first define a temporary arrangement where we save the rows only.
# This part is specially adapted to the given excel, 
#so if any detail is changed in the numbering of the rows, 
#this is where the changes should be made.
Array1Category1tmp = [category1.row_values(i) for i in range(33,39)]  
# Then we take the columns from 1 to 12. This is the arrangement that will be used
Array1Category1 = [Array1Category1tmp[i][1:12] for i in range(0,len(Array1Category1tmp))]


# We repeat the same procedure for each arrangement of each category.
# It will be so, 12 arrangements

Array2Category1tmp = [category1.row_values(i) for i in range(41,47)]  
Array2Category1=[Array2Category1tmp[i][1:12] for i in range(0,len(Array2Category1tmp))]

Array3Category1tmp = [category1.row_values(i) for i in range(49,55)]  
Array3Category1=[Array3Category1tmp[i][1:12] for i in range(0,len(Array3Category1tmp))]

Array4Category1tmp = [category1.row_values(i) for i in range(57,63)]  
Array4Category1=[Array4Category1tmp[i][1:12] for i in range(0,len(Array4Category1tmp))]

############################
###########################
## category 2

Array1Category2tmp = [category2.row_values(i) for i in range(32,38)]  
Array1Category2=[Array1Category2tmp[i][1:12] for i in range(0,len(Array1Category2tmp))]

Array2Category2tmp = [category2.row_values(i) for i in range(40,46)]  
Array2Category2=[Array2Category2tmp[i][1:12] for i in range(0,len(Array2Category2tmp))]

Array3Category2tmp = [category2.row_values(i) for i in range(48,54)]  
Array3Category2=[Array3Category2tmp[i][1:12] for i in range(0,len(Array3Category2tmp))]

Array4Category2tmp = [category2.row_values(i) for i in range(57,63)]  
Array4Category2=[Array4Category2tmp[i][1:12] for i in range(0,len(Array4Category2tmp))]

############################
###########################
## category 3

Array1Category3tmp = [category3.row_values(i) for i in range(35,41)]  
Array1Category3=[Array1Category3tmp[i][1:12] for i in range(0,len(Array1Category3tmp))]

Array2Category3tmp = [category3.row_values(i) for i in range(43,49)]  
Array2Category3=[Array2Category3tmp[i][1:12] for i in range(0,len(Array2Category3tmp))]

Array3Category3tmp = [category3.row_values(i) for i in range(52,58)]  
Array3Category3=[Array3Category3tmp[i][1:12] for i in range(0,len(Array3Category3tmp))]

Array4Category3tmp = [category3.row_values(i) for i in range(61,67)]  
Array4Category3=[Array4Category3tmp[i][1:12] for i in range(0,len(Array4Category3tmp))]

####################################################################

# Now we take the other file, the csv where the data is


with open ('./originalData/data.csv','r') as csv_file:
    dialect = csv.Sniffer().sniff(csv_file.read())
    csv_file.seek(0)
    reader = csv.reader(csv_file, dialect)
    # the variable data contains an arrangement of the entire document
    data=list(reader)
    # with rowcount we have the number of rows in the document
    rowcount=len(data)

    # Temporary empty arrangement where we will do the first reordering
    tmpArray = []

    # In each row we rearrange the columns, so that the last ones
    # from 4231 to the end are between 200 and those that we will modify
    for row in data:
    	new_row = row[:200] + row[4231:] + row[200:4231]
    	tmpArray.append(new_row)

    # We save the first 3 rows of tmpArray in new_array. 
    # Here we will write the following with the reordered columns of 14 in 14
    new_array = tmpArray[:3]
    # We delete the data from the first 3 rows, in the columns from 354 onwards
    new_array[0][354:] = []
    new_array[1][354:] = []
    new_array[2][354:] = []

    # We run through each row of data from the third to the end
    for file in range(3,len(tmpArray)):
    	# This index reaches up to 288 because it is the number of times 
    	# we have the 14 columns
	    for index in range(0,288):
	    	# we define a temporary row that contains the first 340 columns that are fixed
	    	tmp_row = tmpArray[file][:340]
	    	# This row is extended with the data according to this 
	    	# index count in the second variable.
	    	# for index = 0, are the columns from 340 to 354
	    	# for index = 1, columns 354 to 368
	    	# etc
	    	tmp_row.extend(tmpArray[file][340 + 14*index :340 +  14*index + 14])
	    	# This extended row is added to the array that we defined
	    	# and that already has the first three rows.
	    	new_array.append(tmp_row)


# Here we join the first part with the second. 
# That is, we put the data of the excel, the arrangements that we did, 
#in the new_array that we have just defined

# We iterate further in the number of original rows, 
# since this process we will do once for each original row of data, 
#so we have here the count on tmpArray
# Inside, we will multiply this k by 288, 
#since it is the number of rows that grew after the reordering
for k in range(0,len(tmpArray)-3):
	# This range of data indicates the 4 arrangements we have for each category
	for i in range(0,4):
		# Here iterating over the number of lines in each arrangement
		for j in range(3,9):
			# note that each array is of the form new_array [file] [column]
			# ie in the first variable we define the row and in the second the column
			# leaving the columns fixed, because they are all where same, we make
			# each line of the category arrangement is included in the new_array line
			# So, for the triplet k = 0, i = 0, j = 3, becomes
			# new_array [3] [340: 351] = Array1Category1 [0]
			# that is, the first row of array 1 of category 1 is saved in the
			# row 3, columns 340 to 351.
			# etc...
			new_array[(j + i * 6 )      + k * 288][340:351] = Array1Category1[j-3]
			new_array[(j + i * 6 + 24)  + k * 288][340:351] = Array2Category1[j-3]
			new_array[(j + i * 6 + 48)  + k * 288][340:351] = Array3Category1[j-3]
			new_array[(j + i * 6 + 72)  + k * 288][340:351] = Array4Category1[j-3]
			new_array[(j + i * 6 + 96)  + k * 288][340:351] = Array1Category2[j-3]
			new_array[(j + i * 6 + 120) + k * 288][340:351] = Array2Category2[j-3]
			new_array[(j + i * 6 + 144) + k * 288][340:351] = Array3Category2[j-3]
			new_array[(j + i * 6 + 168) + k * 288][340:351] = Array4Category2[j-3]
			new_array[(j + i * 6 + 192) + k * 288][340:351] = Array1Category3[j-3]
			new_array[(j + i * 6 + 216) + k * 288][340:351] = Array2Category3[j-3]
			new_array[(j + i * 6 + 240) + k * 288][340:351] = Array3Category3[j-3]
			new_array[(j + i * 6 + 264) + k * 288][340:351] = Array4Category3[j-3]


# You can save a first output file with this data,
# I had it here to validate the code
# If you need it, uncomment those 3 lines below

#wtr = csv.writer (open ('out.csv', 'w'), delimiter = ';')
#for x in new_array:
# wtr.writerow (x)

# Now we filter to define a new arrangement, where we will append the row
# only if column 351 has data. Column 351 in the new arrangement
# is the first one that is found after placing the matrix of the categories
arrayFiltered =[]

# as a result of this iteration the arrayFiltered with the answers 
# is contained in column 351.
for file in new_array:
	if len(file[351]) > 0:
		arrayFiltered.append(file)


# We are going to list the Id
listId = [arrayFiltered[x][8] for x in range(3,len(arrayFiltered))]

# We define an empty set
seen = set()

# With this we can have an ordered list of the ids, but where they are only once
# Note that this list must have length equal to the number of rows of the data
# as they come initially (without the first 3, of course)
listIdUnique = [x for x in listId if not (x in seen or seen.add(x))]

# Now let's go over the list of unique ids and we'll count the
# number of times repeated
for i in range(0,len(listIdUnique)):
	# if in the original id list, this first id is repeated 9 times,
	# then we will look for the repeated rows
	if listId.count(listIdUnique[i]) == 9:	
		# This firstRowIndex tells us where the block of 9 starts. 
		firstRowIndex = listId.index(listIdUnique[i])
		# The following is actually the arrangement as we have above the categories,
		# but with 9 rows, where the repeated ones are
		listTmp = [arrayFiltered[firstRowIndex + j + 3][340:351] for j in range(0,9)]
		# Knowing that the last three will be those that are repeated among the first 6,
		# we look for the repetition index between the listTmp list and keep it
		# in the following 3 variables
		firstRepeated = listTmp.index(listTmp[6])
		secondRepeated = listTmp.index(listTmp[7])
		thirdRepeated = listTmp.index(listTmp[8])
		# Then that index helps us to add the informative message 
		# in the additional column
		arrayFiltered[firstRowIndex + 6 + 3].append(firstRepeated + 1)
		arrayFiltered[firstRowIndex + 7 + 3].append(secondRepeated + 1)
		arrayFiltered[firstRowIndex + 8 + 3].append(thirdRepeated + 1)

# The name column is located in column 255
# We compare columns 76 to 78 (both inclusive) with 255
# If in any case, it exactly matches the name in 255,
# the column with the string 'inner' is added,
for file in arrayFiltered:
	if file[76] == file[255] or file[77] == file[255] or file[78] == file[255] :
		file.append('inner')
	# If, on the contrary, the one who matches is column 79 with the name in 255,
	# 'outer' is added
	elif file[79] == file[255]:
		file.append('outer')


# We finally write arrayFiltered in the output file.
wtr = csv.writer(open ('outFiltered.csv', 'w'), delimiter=';')
for x in arrayFiltered:
	wtr.writerow (x)
