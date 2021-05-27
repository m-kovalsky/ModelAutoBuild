#r "System.IO"
#r "Microsoft.Office.Interop.Excel"

using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

string fileName = @"C:\Desktop\ModelAutoBuild"; // Enter Model Auto Build Excel file
string myWorkbook = fileName + ".xlsx";
var excelApp = new Excel.Application();
excelApp.Visible = false;
excelApp.DisplayAlerts = false;
Excel.Workbook wb = excelApp.Workbooks.Open(myWorkbook);

string[] tabs = {"DataSources", "Tables", "MeasuresColumns", "Relationships", "Model", "Roles", "RLS", "Hierarchies"};
int tabCount = tabs.Count();
string[] tableTypes = {"FACT","DIM","BRIDGE","SEC","META"};
string[] objectTypes = {"column","measure"};
bool autoGenRel = false;

for (int i=0; i<tabCount; i++)
{
    string wsName = tabs[i];
    Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[wsName];
    Excel.Range xlRange = (Excel.Range)ws.UsedRange;

    int rowCount = xlRange.Rows.Count;
	int colCount = xlRange.Columns.Count;

	// If no relationships defined, enable relationship auto-gen
    if (i==3 && rowCount <= 1)
    {
    	autoGenRel = true;
    }

	for (int r=2; r<=rowCount; r++)
	{
		// Data Sources
		if (i==0)
		{
		    string srvr = (string)(ws.Cells[r,1] as Excel.Range).Text.ToString();
		    string ic = (string)(ws.Cells[r,2] as Excel.Range).Text.ToString();
		    string provider = (string)(ws.Cells[r,3] as Excel.Range).Text.ToString().ToLower();
		    string userID = (string)(ws.Cells[r,4] as Excel.Range).Text.ToString();
		    string pwd = (string)(ws.Cells[r,5] as Excel.Range).Text.ToString();
		    string dsName = srvr + " " + ic;
		    string conn = "Data Source=" + srvr + ";Initial Catalog=" + ic;

		    if (userID != "")
		    {
		        conn = conn + ";User ID=" + userID + ";Password=" + pwd;
		    }
		    
		    if (!Model.DataSources.Any(a => a.Name == dsName))
		    {
		    	var obj = Model.AddDataSource(dsName);
		    }

			var ds = (Model.DataSources[dsName] as ProviderDataSource);

		    if (provider == "sql")
		    {
		        ds.Provider = "System.Data.SqlClient";
		        ds.ConnectionString = conn + ";Integrated Security=True;Persist Security Info=True";
		    }
		    else if (provider == "databricks")
		    {
		        ds.Provider = "";
		        ds.ConnectionString = "Provider=MSDASQL;DSN=" + srvr;
		    }
			
		}

		// Tables
		else if (i==1)
		{			
			string tableName = (string)(ws.Cells[r,1] as Excel.Range).Text.ToString();
			string dataSource = (string)(ws.Cells[r,2] as Excel.Range).Text.ToString();
			string tableType = (string)(ws.Cells[r,3] as Excel.Range).Text.ToString().ToUpper();
			string schema = (string)(ws.Cells[r,4] as Excel.Range).Text.ToString();
			string dc = (string)(ws.Cells[r,5] as Excel.Range).Text.ToString().ToLower();
			string m = (string)(ws.Cells[r,6] as Excel.Range).Text.ToString().ToLower();
			string desc = (string)(ws.Cells[r,7] as Excel.Range).Text.ToString();			

			// Delete table if it exists
			foreach (var t in Model.Tables.Where(a => a.Name == tableName).ToList())
			{
				t.Delete();
			}

			if (!Model.DataSources.Any(a => a.Name == dataSource))
			{
				Error("Unable to create the '"+tableName+"' as it is assigned an invalid data source: '"+dataSource+"'.");
				return;
			}

			if (!tableTypes.Contains(tableType))
			{
				Error("Unable to create the '"+tableName+"' as it is assigned an invalid table type: '"+tableType+"'.");
				return;
			}

			var obj = Model.AddTable(tableName);
			obj.Partitions[0].DataSource = Model.DataSources[dataSource];
			obj.Partitions[0].Query = "SELECT * FROM [" + schema + "].[" + tableType + "_" + tableName + "]";
			obj.Description = desc;

			if (dc == "date")
			{
				obj.DataCategory = "Time";
			}

			if (m.StartsWith("direct") || m == "dq")
			{
				obj.Partitions[0].Mode = ModeType.DirectQuery;
			}			
		}

		// Measures & columns
		else if (i==2)
		{			
			string objectName = (string)(ws.Cells[r,1] as Excel.Range).Text.ToString();
			string tableName = (string)(ws.Cells[r,2] as Excel.Range).Text.ToString();
			string objectType = (string)(ws.Cells[r,3] as Excel.Range).Text.ToString().ToLower();
			string sourceColumn = (string)(ws.Cells[r,4] as Excel.Range).Text.ToString();
			string dt = (string)(ws.Cells[r,5] as Excel.Range).Text.ToString().ToLower();
			string expr = (string)(ws.Cells[r,6] as Excel.Range).Text.ToString();
			string hide = (string)(ws.Cells[r,7] as Excel.Range).Text.ToString().ToLower();
			string fmt = (string)(ws.Cells[r,8] as Excel.Range).Text.ToString();
			string key = (string)(ws.Cells[r,9] as Excel.Range).Text.ToString().ToLower();
			string summ = (string)(ws.Cells[r,10] as Excel.Range).Text.ToString().ToLower();
			string displayFolder = (string)(ws.Cells[r,11] as Excel.Range).Text.ToString();
			string dataCategory = (string)(ws.Cells[r,12] as Excel.Range).Text.ToString();
			string sortByCol = (string)(ws.Cells[r,13] as Excel.Range).Text.ToString();
			string desc = (string)(ws.Cells[r,14] as Excel.Range).Text.ToString();			

			if (!Model.Tables.Any(a => a.Name == tableName))
			{
				Error("'" + tableName + "' is not a valid table in this model. Attempt for '" + objectType + "' " + objectName);
				return;
			}

			if (!objectTypes.Contains(objectType))
			{
				Error("Error generating a measure/column. Object type must be 'Column' or 'Measure'.");
				return;
			}

			// Add column properties
		    if (objectType == "column")
		    {
		        if (String.IsNullOrEmpty(expr))
		        {
		            var obj = Model.Tables[tableName].AddDataColumn(objectName);
		            obj.SourceColumn = sourceColumn;
		        }
		        else
		        {
		            var obj = Model.Tables[tableName].AddCalculatedColumn(objectName);
		            obj.Expression = expr;
		        }

		        var col = Model.Tables[tableName].Columns[objectName];
		        string colName = col.Name;

		        if (dt.StartsWith("int"))
		        {
		            col.DataType = DataType.Int64;
		        } 
		        else if (dt == "string")
		        {
		            col.DataType = DataType.String;
		        }
		        else if (dt == "datetime")
		        {
		            col.DataType = DataType.DateTime;
		        }
		        else if (dt == "double")
		        {
		            col.DataType = DataType.Double;
		        }
		        if (hide.StartsWith("y"))
		        {
		            col.IsHidden = true;
		        }
		        if (key.StartsWith("y"))
		        {
		            col.IsKey = true;
		        }
		        if (summ == "none") 
		        {
		            col.SummarizeBy = AggregateFunction.None;
		        }

		        col.DisplayFolder = displayFolder;
		        col.DataCategory = dataCategory;
		        col.Description = desc;

		        if (sortByCol != "")
		        {
		        	if (!Model.Tables[tableName].Columns.Any(a => a.Name == sortByCol))
		        	{
		        		Error("The column '"+colName+"' cannot be sorted by the column '"+sortByCol+"' as the sort-by column does not exist in the '"+tableName+"' table.");
		        		return;
		        	}
		            col.SortByColumn = Model.Tables[tableName].Columns[sortByCol]; 
		        }
		        if (fmt == "Whole Number")
		        {
		            col.FormatString = "#,0";
		        }
		        else if (fmt == "Percentage")
		        {
		            col.FormatString = "#,0.0%;-#,0.0%;#,0.0%";
		        }
		        else if (fmt == "Month Year")
		        {
		            col.FormatString = "mmmm YYYY";
		        }
		        else if (fmt == "Currency")
		        {
		            col.FormatString = "$#,0;$#,0;$#,0";
		        }
		        else if (fmt == "Decimal")
		        {
		            col.FormatString = "#,0.0";
		        }
		    }

		    // Add measure properties
		    if (objectType == "measure")
		    {
			    var obj = Model.Tables[tableName].AddMeasure(objectName);
				obj.Expression = expr;
				obj.DisplayFolder = displayFolder;
				obj.Description = desc;

				if (hide.StartsWith("y"))
				{
					obj.IsHidden = true;
				}
				if (fmt == "Whole Number")
				{
					obj.FormatString = "#,0";
				}
				else if (fmt == "Percentage")
				{
					obj.FormatString = "#,0.0%;-#,0.0%;#,0.0%";
				}
				else if (fmt == "Currency")
				{
					obj.FormatString = "$#,0;$#,0;$#,0";
				}
				else if (fmt == "Month Year")
				{
					obj.FormatString = "mmmm YYYY";  
				}
		    }
		}

		// Relationships
		else if (i==3)
		{
			string fromTable = (string)(ws.Cells[r,1] as Excel.Range).Text.ToString();
			string fromColumn = (string)(ws.Cells[r,2] as Excel.Range).Text.ToString();
			string toTable = (string)(ws.Cells[r,3] as Excel.Range).Text.ToString();
			string toColumn = (string)(ws.Cells[r,4] as Excel.Range).Text.ToString();
			string act = (string)(ws.Cells[r,5] as Excel.Range).Text.ToString().ToLower();
			string cfb = (string)(ws.Cells[r,6] as Excel.Range).Text.ToString().ToLower();

			if (!Model.Tables.Any(a => a.Name == fromTable))
			{
				Error("The relationship cannot be created because the table '"+fromTable+"' does not exist.");
				return;
			}

			if (!Model.Tables.Any(a => a.Name == toTable))
			{
				Error("The relationship cannot be created because the table '"+toTable+"' does not exist.");
				return;
			}

			if (!Model.Tables[fromTable].Columns.Any(a => a.Name == fromColumn))
			{
				Error("The relationship cannot be created because the column '"+fromColumn+"' does not exist in the "+fromTable+".");
				return;
			}

			if (!Model.Tables[toTable].Columns.Any(a => a.Name == toColumn))
			{
				Error("The relationship cannot be created because the column '"+toColumn+"' does not exist in the "+toTable+".");
				return;
			}

			// This assumes that the model does not already contain a measure with the same name (if it does, the new measure will get a numeric suffix):
		    var obj = Model.AddRelationship();
            obj.FromColumn = Model.Tables[fromTable].Columns[fromColumn];
            obj.ToColumn = Model.Tables[toTable].Columns[toColumn];
           
            if (act == "no")
            {
                obj.IsActive = false;
            }

            if (cfb == "bi")
            {
                obj.CrossFilteringBehavior = CrossFilteringBehavior.BothDirections;
            }         
		}

		else if (i==4)
		{
			string name = (string)(ws.Cells[r,1] as Excel.Range).Text.ToString();
			string m = (string)(ws.Cells[r,2] as Excel.Range).Text.ToString().ToLower();
			string prem = (string)(ws.Cells[r,3] as Excel.Range).Text.ToString().ToLower();

			if (name == "")
			{
				Error("Model name must be provided on the 'Model' tab.");
				return;
			}
			else
			{
				Model.Database.Name = name;
    			Model.Database.ID = name;
			}
			
    		 // Update model mode
		    if (m.StartsWith("direct") || m == "dq")
		    {
		        Model.DefaultMode = ModeType.DirectQuery;
		    }

		    // This enables overwriting deployments to Power BI Premium
		    if (prem.StartsWith("y"))
		    {
		        Model.DefaultPowerBIDataSourceVersion = PowerBIDataSourceVersion.PowerBI_V3;
		        Model.Database.CompatibilityMode = CompatibilityMode.PowerBI;
		        Model.Database.CompatibilityLevel = 1560;
		    }
		}

		// Roles
		else if (i==5)
		{
			string ro = (string)(ws.Cells[r,1] as Excel.Range).Text.ToString();
			string rm = (string)(ws.Cells[r,2] as Excel.Range).Text.ToString();
			string mp = (string)(ws.Cells[r,3] as Excel.Range).Text.ToString().ToLower();

			// Add Roles and do not duplicate
		    if (!Model.Roles.ToList().Any(x => x.Name == ro))
		    {
			    var obj = Model.AddRole(ro);
			    obj.RoleMembers = rm;

			    if (mp == "read")
			    {
			        obj.ModelPermission = ModelPermission.Read;
			    }
			    else if (mp == "admin")
			    {
			        obj.ModelPermission = ModelPermission.Administrator;
			    }
			    else if (mp == "refresh")
			    {
			        obj.ModelPermission = ModelPermission.Refresh;
			    }
			    else if (mp == "readrefresh")
			    {
			        obj.ModelPermission = ModelPermission.ReadRefresh;
			    }
			    else if (mp == "none")
			    {
			        obj.ModelPermission = ModelPermission.None;
			    }
			}
		}

		// RLS
		else if (i==6)
		{
			string ro = (string)(ws.Cells[r,1] as Excel.Range).Text.ToString();
			string tableName = (string)(ws.Cells[r,2] as Excel.Range).Text.ToString();
			string rls = (string)(ws.Cells[r,3] as Excel.Range).Text.ToString();
		    int rlsLength = rls.Length;

		    if (!Model.Tables.Any(a => a.Name == tableName))
		    {
		    	Error("Row level security for the '"+ro+"' role cannot be created since the '"+tableName+"' table does not exist.");
		    	return;
		    }

		    if (!Model.Roles.Any(a => a.Name == ro))
		    {
		    	Error("Row level security for the '"+ro+"' role cannot be created since the role does not exist.");
		    	return;
		    }

		    if (rls[0] == '"')
	        {
				rls = rls.Substring(1,rlsLength - 2);
	        }
		    
		    rls = rls.Replace("\"\"","\"");    
		    
		    Model.Tables[tableName].RowLevelSecurity[ro] = rls;  
		}

		// Hierarchies
		else if (i==7)
		{
			try
			{
				string hierarchyName = (string)(ws.Cells[r,1] as Excel.Range).Text.ToString();
				string tableName = (string)(ws.Cells[r,2] as Excel.Range).Text.ToString();
				string columnName = (string)(ws.Cells[r,3] as Excel.Range).Text.ToString();

				if (!Model.Tables.Any(a => a.Name == tableName))
				{
					Error("The hierarchy '"+hierarchyName+"' cannot be created because the table '"+tableName+"' does not exist.");
					return;
				}

				if (!Model.Tables[tableName].Columns.Any(a => a.Name == columnName))
				{
					Error("The hierarchy '"+hierarchyName+"' cannot be created because the column '"+columnName+"' does not exist in the '"+tableName+"' table.");
					return;
				}

			    if (!Model.AllHierarchies.ToList().Any(x => x.Name == hierarchyName))
		        {
		            // Add the hierarchy
		            var obj = Model.Tables[tableName].AddHierarchy(hierarchyName);
		        }
			    
			    // Add each level of each hierarchy
			    Model.Tables[tableName].Hierarchies[hierarchyName].AddLevel(columnName);
			}
			catch
			{	    
			}
		}
	}
}

// Remove quotes from Expression and FormatString (format all DAX measures)
foreach(var o in Model.AllMeasures.ToList())
{
    string expr = o.Expression;
    int exprLength = expr.Length;
    
    // Remove quotes from Expressions
    if (expr[0] == '"')
    {
        o.Expression = expr.Substring(1,exprLength - 2);
    }
    
    o.Expression = o.Expression.Replace("\"\"","\"");
     
    // Remove quotes from Format Strings
    o.FormatString = o.FormatString.Trim('"');

    // Uncomment the line below if you want the DAX to be formatted.
    //FormatDax(o);
}

 // Remove quotes from Expression and FormatString (format all DAX calculated columns)
foreach(var o in Model.AllColumns.Where(a => a.Type.ToString() == "Calculated").ToList())
{
	var obj = (Model.Tables[o.Table.Name].Columns[o.Name] as CalculatedColumn);
    string expr = obj.Expression;
    int exprLength = expr.Length;
    
    // Remove quotes from Format Strings
    o.FormatString = o.FormatString.Trim('"');
    
    // Remove quotes from Expressions
    if (expr[0] == '"')
    {
        obj.Expression = expr.Substring(1,exprLength - 2);
    }
    
    obj.Expression = obj.Expression.Replace("\"\"","\"");

    // Uncomment the line below if you want the DAX to be formatted.
    //FormatDax(obj);
}

// Replaces \n with new line (measures)
foreach(var o in Model.AllMeasures.ToList())
{
    // Replaces \n with new line
    o.Expression = o.Expression.Replace("\\n", "\r\n");
}

// Replaces \n with new line (calculated columns)
foreach(var o in Model.AllColumns.Where(a => a.Type.ToString() == "Calculated").ToList())
{
	var obj = (Model.Tables[o.Table.Name].Columns[o.Name] as CalculatedColumn);
    obj.Expression = obj.Expression.Replace("\\n", "\r\n");
}

// Auto-generate relationships
if (autoGenRel)
{
	string keySuffix = "Id";
    // Loop through all rows but skip the first one:
	foreach (var tbl in Model.Tables.Where(a => a.Partitions[0].Query.Contains("FACT_")))
    {
        foreach (var factColumn in tbl.Columns.Where(a => a.Name.EndsWith(keySuffix)))
        {
            var dim = Model.Tables.FirstOrDefault(t => factColumn.Name == t.Name + keySuffix);

            if (dim != null)
            {
                var dimColumn = dim.Columns.FirstOrDefault(c => factColumn.Name == c.Name);
                if (dimColumn != null)
                {
                    var rel = Model.AddRelationship();
                    rel.FromColumn = factColumn;
                    rel.ToColumn = dimColumn;
                }
            }
        }
    }
}

wb.Close();
excelApp.Quit();
System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);