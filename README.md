i Have an Excel file . i Need to apply python code for the whole process. 
In the Excel file , i want to Hide some columns , i want to apply Formula in some columns starting from a particular row and ending at a particular row . 

The exact process is like this , 
open all the Excel files in the current given Directory one after the other except the file starting with Kostenstelle because that file is used for v-lookup Purpose only. 
In the Excel file which is opened,

first identify whether the sheet Name is written in green or yellow . 
it can be any green or any type of yellow. 
So if the program identifies the place where we write the sheet Name is filled with yellow , then do the following process. 

at first identify the end row , the end row is identified by the following logic , 

in column index A , starting from row 7 , in which cell it find the Keyword Summe written in bold  . 
then mark that row as the end row . If it couldnt find the Summe written in bold in column index A anywhere then search from row 7 onwards where in which row in column index A we can find exact match with the sheet Name . For that extract the sheet Name and comapre it with the values starting from row 7. If it find any matching value with the sheet Name (case insensitive) then consider that row as the end row. If it couldnt find that also then the 
next fall back logic is like , starting from row 7 onwards , check in which cell it finds no value . so if it finds a cell with no value in column index A , then consider the cell Above and fix that as the end row. 

After Fixing the end row then the next step is that , in row 3 or  row 4 , in which column  it identifies the Keyword Veränderung(case insensitive). After identifying that column or  column index , add two new columns to the left of that identified column or column index . the column Right of the identified column , identified column itself, and the two newly added two columns shouldnt be hidden.
in the column or column index  , which is direct left of the identified column index , please write in row 3 Plan and in row 4 current year + 1 . For example in row 3 write Plan and in row 4 write 2026 if the current year is 2025. 
Both should be written in bold and should be placed in the Center of the cell. 
then in the current  Directory , take the file that starts with the Name Kostenstelle.
Then next step is that we are Performing a v-look up and taking values from  the file whose  Name starts with Kostenstelle .
for each row in the current working sheet starting from row 5 till the end row check if the value in column index AB Matches with the value in column index A in the Kostenstelle file, then copy the corresponding
value from column index D to the corresponding row in  column or column index direct left to the identified column index.
please note that the value in rows in column index AB are like 4557 775 67575, 47647648686,897598757959  in a single row. In other case row  value are like J7799 that means only one value.
In the first case if the value in the row is like 4557 775 67575, 47647648686,897598757959  here take three values for lookup and if it found a match in column index A , then we will have three different values from column index d 
in Kostenstelle file , add that together and write a single value in row in  current column index in the current sheet. then the next case if the lookup value is like J7799 , then search this value in column index A 
of the kostenstelle file such that exact 1 to 1 match may be there, or else the value in column index A if ist like t666654/J7799 , even though ist not a 1 to 1 match , but still its a match , then also v.look up should function without any Problem. 


then in the column index, which is not direct left of the identified column index , write IST in row 3 and write current year e in row 4. For example in row 3 write IST and in row 4 write 2025e if the current year is 2025. 
Both should be written in bold and should be placed in the Center of the cell. 
then in the current  Directory , take the file that starts with the Name Kostenstelle.
Then next step is that we are Performing a v-look up and taking values from  the file whose  Name starts with Kostenstelle .
for each row in the current working sheet starting from row 5 till the end row check if the value in column index AB Matches with the value in column index A in the Kostenstelle file, then copy the corresponding
value from column index C to the corresponding row in  column index which is not direct left to the identified column index.
please note that the value in rows in column index AB are like 4557 775 67575, 47647648686,897598757959  in some rows. In other rows the value are like J7799 .
In the first case if the value in the row is like 4557 775 67575, 47647648686,897598757959  here take three values for lookup and if it found a match in column index A , then we will have three different values from column index C
in Kostenstelle file , add that together and write a single value in row in  current column index in the current sheet. then the next case if the lookup value is like J7799 , then search this value in column index A 
of the kostenstelle file such that exact 1 to 1 match may be there, or else the value in column index A if ist like t666654/J7799 , even though ist not a 1 to 1 match , but still ist a match , then also v.look up should function without any Problem. 

if for both column Indices if there is no value in column index AB, then do Nothing.

then the next step is that , identify the column which has Keyword Plan in row 3 and the current year in row 4. this column shouldnt be hidden. 
then the next step is that , identify the column which has Keyword IST in row 3 and the previous year in row 4 . for example IST in row 3 and 2024e or 2024 in row 4. this column shouldnt be hidden. 
then the next step is that , identify the column which has Keyword IST in row 3 and the current year-2 in row 4 . for example IST in row 3 and 2023e or 2023 in row 4. this column shouldnt be hidden. 
the column index A also shouldnt be hidden and all other other columns in the current sheet should be hidden.











