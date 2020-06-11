Model Auto Build is a framework that dynamically creates a tabular model based on an Excel template. This framework is compatible for all destinations of tabular models - SQL Server Analysis Services, Azure Analysis Services, and Power BI Premium.

Instructions:

1.) Download the following files and save them to a single folder on your computer.

      ModelAutoBuild.xlsx
      ExcelToTextMaster.exe
      ModelAutoBuild.cs
      BlankModelTemplate.bim

2.) Open the ModelAutoBuild.xlsx file.

3.) Populate the columns in each of the tabs, following the instructions within the notes shown on the header rows.

4.) Open the ModelAutoBuild.cs in a Text Editor (i.e. Notepad, Notepad ++, Sublime).

5.) On the first line of code, change the folderName parameter to the folder that contains all of the files in Step 1. 
    
    Here is an example:
    
    var folderName = @"C:\Documents\ModelAutoBuild\";
    
6.) Save and close the ModelAutoBuild.cs file.

7.) Open the Command Prompt.

8.) Run the ExcelToTextMaster program as shown below. The folder used for each of the two parameters below should be the same folder used in Step 1.

    Here is an example:
    
    "C:\Documents\ModelAutoBuild\ExcelToTextMaster.exe" "C:\Documents\ModelAutoBuild\"

9.) Make sure you have Tabular Editor installed on your computer. Here is a link to download it: https://github.com/otykier/TabularEditor/releases

10.) Run the following in the Command Prompt. Ensure that the location of Tabular Editor matches where it is stored on your computer. Also ensure that the folder used in the -S and -B paramters is the same as the folder from Step 1.

    start /wait /d "C:\Program Files (x86)\Tabular Editor" TabularEditor.exe "C:\Documents\ModelAutoBuild\BlankModelTemplate.bim" -S "C:\Documents\ModelAutoBuild\ModelAutoBuild.cs" -B "C:\Documents\ModelAutoBuild\NewModel.bim"
    
After completing Step 10, your new .bim file is ready. It is in the location specified under the -B parameter in Step 10. It may be opened in Tabular Editor.

If you want to deploy the model to SQL Server Analysis Services of Azure Analysis Services, view the instructions here:
https://github.com/otykier/TabularEditor/wiki/Command-line-Options

If you want to deploy the model to Power BI Premium, view the instructions here:
https://github.com/TabularEditor/tabulareditor.github.io/blob/master/_posts/2020-06-02-PBI-SP-Access.md
