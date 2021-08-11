# autoSynopsis
Reads through word document, collecting data from specified tables. Config file allows user to autogenerate sentences based on data.

Config file is a word document. Please note that the column headers must be exactly the same. The column number and description have no effect on code. To access a piece of data, you must put that data's column name, table number (not 0 indexed to be non-coder friendly), and row number (not 0 indexed). After running the script, click the configuration button to select your config docx.

Important: The logic must be set up as seen in the screen shot. This version only deals with a few logic cases. If the data point is a number, you can compare it to another number (>, <, =). If the data is a yes or no (boolean), it can compare that. Finally, if you want to display the data in the sentence, you must denote that with brackets and the column number. 

For example, if I want to show the second column in my sentence (i.e. Life Expectancy as seen in the screen shot), I must put "{2}" as seen in the screen shot. 

If all conditions in the logic section pass, the statement is written to the docx. 

Note that there are multiple logic/statement tables. You can have as many of these tables as you want, but they must all have the same column names (i.e. Logic and Statment). This is to help organization of paragraphs. Paragraphs are generated based on these tables. So if you have 3 tables and statments pass in all three tables, your docx will have 3 paragraphs. This can be seen in the output file.

![Config File](images/config.png?raw=true "Config")
 
 Once the config file is chosen, text should appear stating that the lists were made. If this is the case, click the generate synopsis button. From there you will be prompted to select a folder. Select a folder containing files/tables you want to parse through (must be docx files). These files should look something like the following.
 
 ![Config File](images/rawdata.png?raw=true "Config")
 
 
 From there, you will be prompted to select a folder to save outputted docx files. After selecting that folder, you should get a message saying the script has completed its job. Output files will look like the following.
 
 ![Config File](images/output.png?raw=true "Config")
 
 Feel free to alter how the data comes out and such (i.e. make the docx somewhat more presentable). I left it very simple as my friend's use for it is very basic and the docx file is not being presented anywhere.
