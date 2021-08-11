# autoSynopsis
Reads through word document, collecting data from specified tables. Config file allows user to autogenerate sentences based on data.

Config file is a word document. Please note that the column headers must be exactly the same. The column number and description have no effect on code. To access a piece of data, you must put that data's column name, table number (not 0 indexed to be non-coder friendly), and row number (not 0 indexed). After running the script, click the configuration button to select your config docx.

![Config File](images/config.png?raw=true "Config")
 
 Once the config file is chosen, text should appear stating that the lists were made. If this is the case, click the generate synopsis button. From there you will be prompted to select a folder. Select a folder containing files/tables you want to parse through (must be docx files). These files should look something like the following.
 
 ![Config File](images/rawdata.png?raw=true "Config")
 
 
 From there, you will be prompted to select a folder to save outputted docx files. After selecting that folder, you should get a message saying the script has completed its job. Output files will look like the following.
 
 ![Config File](images/output.png?raw=true "Config")
 
 Feel free to alter how the data comes out and such (i.e. make the docx somewhat more presentable). I left it very simple as my friend's use for it is very basic and the docx file is not being presented anywhere.
