# [Model Auto Build](https://www.elegantbi.com/post/modelautobuild "Model Auto Build")

Model Auto Build is a framework that dynamically creates a tabular model based on an Excel template. This framework is compatible for all destinations of tabular models - [SQL Server Analysis Services](https://docs.microsoft.com/analysis-services/ssas-overview?view=asallproducts-allversions "SSAS"), [Azure Analysis Services](https://azure.microsoft.com/services/analysis-services/ "Azure AS"), and [Power BI Premium](https://powerbi.microsoft.com/power-bi-premium/ "Power BI Premium") (using the [XMLA R/W endpoint](https://docs.microsoft.com/power-bi/admin/service-premium-connect-tools "XMLA R/W endpoint")). This framework is also viable for both in-memory and direct query models.

![](https://github.com/m-kovalsky/ModelAutoBuild/blob/master/Images/ExcelTemplate.png)

## Purpose

To provide a framework for business stakeholders and developers when initially outlining a model. When completed, the Excel template serves as a blueprint for the tabular model - akin to a blueprint for designing a building. 

This framework speeds up the development time of the model once the blueprint has been laid out. Development time can be spent on more advanced tasks such as solving DAX challenges or complex business logic requirements.

Lastly, many people who are new to Power BI are more familiar with Excel. Since the framework is based in Excel it provides a familiar environment for such folks. 

## Instructions

1.) Download the following files from the ModelAutoBuild folder and save them to a single folder on your computer.

      ModelAutoBuild.xlsx
      ModelAutoBuild.cs
      ModelAutoBuild_Example.xlsx (this file shows an example of a properly filled out ModelAutoBuild.xlsx file)

2.) Open the ModelAutoBuild.xlsx file.

3.) Populate the columns in each of the tabs, following the instructions within the notes shown on the header rows. Close the Excel file when finished.

4.) Open [Tabular Editor](https://tabulareditor.com/ "Tabular Editor") and create a new model (File -> New Model).

5.) Paste the ModelAutoBuild.cs into the [Advanced Scripting](https://docs.tabulareditor.com/Advanced-Scripting.html#working-with-the-model-object "Advanced Scripting") window within Tabular Editor.

6.) Update the fileName parameter (on the 7th line of code) to be the location and file name of your saved ModelAutoBuild.xlsx file (see the example below).
    
```C#    
string fileName = @"C:\Desktop\ModelAutoBuild";
```

7.) Click the 'Play' button (or press F5).
  
After completing Step 7, your model has been created within Tabular Editor.

If you want to deploy the model to SQL Server Analysis Services or Azure Analysis Services, view Tabular Editor's [Command Line Options](https://github.com/otykier/TabularEditor/wiki/Command-line-Options "Command Line Options").

If you want to deploy the model to Power BI Premium, view the instructions on this [post](https://github.com/TabularEditor/tabulareditor.github.io/blob/master/_posts/2020-06-02-PBI-SP-Access.md "post").

## Additional Notes

* It is not necessary to fill in all the details of the model. For example, the Expression (DAX) and other such elements may be created afterwards. The goal of this framework is not to create a completed model per say but to quickly and intelligently build the foundation.

* If you want a column to be a calculated column, simply add in the DAX in the Expression column. Columns that have DAX expressions will automatically become calculated columns. If there is no expression they will default to a data column (where you must enter a source column). Note of caution: try your best to avoid calculated columns. If in doubt, view Best Practice #6 within this post: https://www.elegantbi.com/post/top10bestpractices.

* The partition queries generated by this framework are in the following format (example below is of a fact table). This is a best practice and ensures no logic is housed within the partition query.
     
```SQL
SELECT * FROM [SchemaName].[FACT_TableName]
```

## Requirements

* [Tabular Editor](https://tabulareditor.com/ "Tabular Editor") version 2.13.0 or higher


## Version History

* 2021-05-27 [Version 1.4.2](https://github.com/m-kovalsky/ModelAutoBuild/releases/tag/1.4.2) released
* 2021-04-30 [Version 1.4.1](https://github.com/m-kovalsky/ModelAutoBuild/releases/tag/1.4.1) released
* 2021-04-14 [Version 1.4.0](https://github.com/m-kovalsky/ModelAutoBuild/releases/tag/1.4.0) released (complete code overhaul; simplified script to be executed in Tabular Editor and pull directly from Excel)
* 2020-07-06 [Version 1.3.0](https://github.com/m-kovalsky/ModelAutoBuild/releases/tag/1.3.0) released (added support for Hierarchies)
* 2020-06-24 [Version 1.2.0](https://github.com/m-kovalsky/ModelAutoBuild/releases/tag/1.2.0) released (added support for Calculated Columns)
* 2020-06-16 [Version 1.1.0](https://github.com/m-kovalsky/ModelAutoBuild/releases/tag/1.1.0) released (added Roles and Row Level Security)
* 2020-06-11 [Version 1.0.0](https://github.com/m-kovalsky/ModelAutoBuild/releases/tag/1.0.0) released on GitHub.com
