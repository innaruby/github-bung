



Current process is like , i need to create a powerpoint presentation from the data from excel.
For that I select each slide , and then in the settings I add a connection or a shortcut link to the 
Excel , where I need to get the data from. It will be as a table. And when the table is 
Aktualised or updated with new data , that gets automatically reflected in the powerpoint slide. 
In order to create this for each slide is really time consuming because each time for each slide I need 
To set the setting in paste option and I need to select the area or range from the excel sheet manually.

The new automated process will like following
, In an excel file look in each sheet if in cell T1 , Whether somelike in the format PPT Folie Number for example PPT Folie 60 is written or not. IF this pattern is recognized in the cell T1 , 
The number in the cell correspond to the sheet 60 in the powerpoint file. All the files that means both the excel and powerpoint are lying in the same folder. 
Only we will do the processing for the sheets in excel where in cell T1 , there is a value like PPT Folie 58  like this is there. The number in the cell can change. 
So the process is like the following , 
In the excel sheet we need to identify a table like grouped data at first then we need to copy its content to the powerpoint slide with a good structure. 
So inorder to identify the table like structure from the powerpoint , we have to identify the four boundaries. 

Lets assume the shape of the boundary as a rectangle . 

Rectangle has 4 edges. Above , below , or upper , lower and then left and right . 
Below edge we can find by following logic , look in column index A , from row 7 onwards 
In which cell the word Summe is written bold. IF it fails to find the word Summe written bold , 

then look in the same column index A in the row starting from row 5 where a value is not there or if the value is same as of sheet name . The comparison with the sheet name should be case insensitive. 
For example in the column index A , till row 9 there are continuous data and in row 10 there is no data. 
In this case we define the row 9 as the below edge of the rectangle. 
Left edge is the left most pane in excel sheet. 
In the cell A1 is the title , which should be write in the title bar in powerpoint slide. 
The value in cell A1 should be excluded from the boundary. 
The top edge or the boundary can be identified by the following logic. 
Identify in which row starting from column index B to F there is a value .That row will be the upper edge. 
Then in order to identify the right edge, we can apply the following logic , 
Look from the column index A , in which column index there is no value. 
For example , in all the rows till column index T there will be a value. In column index U in no row there is a value , then we define the column index T in this example as the right most edge.
Thus we identified the boundaries and thus the rectangular shape, 
After getting this tabular data from a sheet in excel , then identify the value from cell T1 in that sheet. Extract the number from that value, that number corresponds to slide number in powerpoint . So please go to powerpoint and delete all the data in that particular slide number except the data which is in placeholder left or more closest  to the slide number which is in at bottommost right edge corner.  and then  create a title bar which should stay top leftmost most corner in the slide, The value to be written to it is the value from cell A1 from that excel sheet. 
Then create an object field or object place holder in powerpoint , and then paste the table we found from excel on to the powerpoint object field or object placeholder. 
The data or table we copied from excel data and pasted to the powerpoint , should get automatically updated as immediately when the data in the excel table changes. 



 
