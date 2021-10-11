# autoSynopsis
Reads through word document, collecting data from specified tables. Config file allows user to autogenerate sentences based on data.

Config file is a word document. Please note that the column headers must be exactly the same. The column number and description have no effect on code. To access a piece of data, you must put the table's name, a word/phrase in the column the data piece is in, and a word/phrase that appears in the same row.

Important: The logic must be set up as seen in the screen shot. This version only deals with a few logic cases. If the data point is a number, you can compare it to another number (>, <, =, >=, <=). If the data is a yes or no (boolean), it can compare that. You can also do operations on multiple data values & numbers together to make a comparison (+, -, *, /)

Data values are denoted by {3} (two brackets and a number in the middle). For tables where you want to collect 2 data points with the same column name and same consistent row word/phrase, denote the two distinct data points with {3} and {3^} respectively (3 being an example number). If the two data points are in different columns, denote them the same manner but put in two column names in the table. See below in the photos for clarification. 

If all conditions in the logic section pass, the statement is written to the docx. 

Note that there are multiple logic/statement tables. You can have as many of these tables as you want, but they must all have the same column names (i.e. Logic and Statment). This is to help organization of your statements. Statements are outputted sequentially in the output file. 

Note that the script runs through all the documents in a folder are processed sequentially. The script outputs different output word docx for each input file. The configuration file should be set up by the user to handle every possible condition the user wants to cover - the individual input docx files will probably meet only some of those conditions (and thus only some statements from the config file are outputted).

Photos below.

![Config File](images/config2.png?raw=true "Config")
![Config File](images/config22.png?raw=true "Config")
 
 Once the config file is chosen, text should appear stating that the lists were made. If this is the case, click the generate synopsis button. From there you will be prompted to select a folder. Select a folder containing files/tables you want to parse through (must be docx files). These files should look something like the following.
 

 ![Config File](images/inputdocx2.png?raw=true "Config")
 
 From there, you will be prompted to select a folder to save outputted docx files. After selecting that folder, you should get a message saying the script has completed its job. Output files will look like the following.
 
 ![Config File](images/output.png?raw=true "Config")
  ![Config File](images/output3.png?raw=true "Config")
 
 Feel free to alter how the data comes out and such (i.e. make the docx somewhat more presentable). I left it very simple as my friend's use for it is very basic and the docx file is not being presented anywhere.
