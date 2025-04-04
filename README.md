
There will be a gui using tkinter library , user will be asked to give two files .
First file is called input file . Second file is called Kostenstelle. 
Both the files will be selected by user from the directory using the gui mask . 
As the next step the program should create a copy of the input file which is file 1 in the same directory where the file 1 lies.
Then open the copy file  ( or copy of the input file ) then delete all the values in the active worksheet(there will be only work sheet in the file ) starting from row 16.
Then the next step is , look in the original file 1 or original input file , in column index C , starting from row 16. 
So next step is to find till which row it has continuous data in column index C. Then mark the last row where it has continuous data in column index C as the end row. 
The next step is that , copy the values in column index C,D,E,F,H from the original file to the column indices C,D,E,F,H in  copy file starting from row 16 till the end row . 
Then the next step is that , perform a vlook up between the copy file and the Kostenstelle(file2) , such that copy the value from column index E from the file Kostenstelle to column index B in copy file , if the value in column index H in copy file matches with the value in column index A in Kostenstelle file. 
Then the next step is that , write in column index G in copy file the value V0 if the value in column index C start with 705,706or707   in copy file and the value in column index B is 1001. This is case 1. 
Write in column index G the value U0 , if the value in column index C start with 705,706or707   in copy file and the value in column index B is 1002 in copy file . This is case 2. 
write in column index G in copy file the value A0 if the value in column index C start with 704 in copy file and the value in column index B is 1001. This is case 3.
write in column index G in copy file the value D0 if the value in column index C start with 705,706or707   in copy file and the value in column index B is 1002. This is case 4.
If the value in column index C starts with any other number , other than in the above mentioned case, then please don’t write anything on the column index G. this is case 5 .
Please make sure that all the processing of data in copy file are done from row 16 till the end row . 
Then the next step is that , In the column index L in copy file , write the length of text in column index D with spaces in between excluded. And check if the text length is greater than or equal to 50.
IF its 50 or greater than 50 , then mark the cell as red, other wise mark the cell as light green. 
Then the next step is that , we need to perform a v-look up between copy file and Kostenstelle file if the value in column index H in copy file matches the value in column index A in kostenstele file .
IF yes then the look in the column index F in kostenstelle file whether the value is written aktiv or inaktiv( it should be case insensitive). If its aktiv then write in the column index column index M in copyfile for that corresponding value in column index H in copyfile, 
But if the value in column index F is inaktiv for a particular value in column index A in Kostenstelle file , then the next step is that copy the value from column index I from Kostenstelle file to the column index M in copy file after checking  an extra condition. 
The extra condition is that , in the kostenstelle file for a value in column index A , if the value in column index F is inaktiv , then first take its corresponding value in column index I . Bofore writing this value to the column index M in the copy file the check is like this ,    identify the cells which are coloured light green in column index A of the kostenstelle file. Lets call that cells as green cells,. 
Now we check the value in column index I with values in green cells in column index A in kostenstelle file. 
IF both the values matches , then for that particular row , take the value from column index I of the kostenstelle file and write it to the column index M in the copyfile . IF the values in column index I doesn’t match with the value in green cells then we take that value from column index I directly from kostenstelle file to the column index M to the copy file. 

Then the next step is that , look In column index G of the copy file . IF the values are A0 or D0 , then appy this formula =WERT("100000"&RECHTS("00"&H23;3))       and write the value to column index K in copy file . after writing the values in column index K , for those row values in column index H  for which we have written the value  in the column index K , after writing the value in column index K , delete the value in column index H. 
