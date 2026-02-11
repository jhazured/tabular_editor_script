#r "System.IO"
#r "Microsoft.Office.Interop.Excel"

using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

// ============================================================================
// CONFIGURATION
// ============================================================================
string filePath = @"C:\Users\jharker\Documents\DataDictionary";
bool dataSourceM = false; // Set this to true if you want the data source to use M
string excelFilePath = filePath + ".xlsx"; 
string textFilePath = filePath + ".txt";
string modelName = Model.Database.Name;

// ============================================================================
// VALIDATION
// ============================================================================
if (modelName == "SemanticModel")
{
    Error("Please name your model in the properties window: Model -> Database -> Name");
    return;
}

// ============================================================================
// SHEET 1: TABLES, COLUMNS, MEASURES, HIERARCHIES
// ============================================================================
var sbMain = new System.Text.StringBuilder();
string[] colNameMain = { "Model","Table","Object Type","Object","Hidden Flag","Description","Display Folder","Formula/Expression","Format String","Data Type" };
int colNameMainCount = colNameMain.Length;
string newline = Environment.NewLine;

// Add headers
for (int i=0; i < colNameMainCount; i++)
{
    if (i<colNameMainCount-1)
    {
        sbMain.Append(colNameMain[i] + '\t');
    }
    else
    {
        sbMain.Append(colNameMain[i] + newline);
    }
}

// Extract model metadata
foreach (var t in Model.Tables.Where(a => a.ObjectType.ToString() != "CalculationGroupTable").OrderBy(a => a.Name).ToList())
{
    string tableName = t.Name;
    string tableDesc = (t.Description ?? "").Replace("'","''");
    string objectType = "Table";
    string hiddenFlag;                 
    string expr;
    string formatStr = "";
    string dataType = "";

    if (t.IsHidden)
    {
        hiddenFlag = "Yes";
    }
    else
    {
        hiddenFlag = "No";
    }
    
    if (t.SourceType.ToString() == "Calculated")
    {
        expr = (Model.Tables[tableName] as CalculatedTable).Expression;
        expr = expr.Replace("\n"," ").Replace("\r"," ").Replace("\t"," ");
        sbMain.Append(modelName + '\t' + tableName + '\t' + objectType + '\t' + tableName + '\t' + hiddenFlag + '\t' + tableDesc + '\t' + " " + '\t' + expr + '\t' + formatStr + '\t' + dataType + newline);
    }
    else
    {
        sbMain.Append(modelName + '\t' + tableName + '\t' + objectType + '\t' + tableName + '\t' + hiddenFlag + '\t' + tableDesc + '\t' + " " + '\t' + "***N/A***" + '\t' + formatStr + '\t' + dataType + newline);
    }
    
    // Columns
    foreach (var o in t.Columns.OrderBy(a => a.Name).ToList())
    {
        string objectName = o.Name;
        string objectDesc = (o.Description ?? "").Replace("'","''");
        string objectDF = o.DisplayFolder ?? "";
        objectType = "Attribute";
        dataType = o.DataType.ToString();
        
        if (o.IsHidden)
        {
            hiddenFlag = "Yes";
        }
        else
        {
            hiddenFlag = "No";
        }
        
        if (o.Type.ToString() == "Calculated")
        {
            expr = (Model.Tables[tableName].Columns[objectName] as CalculatedColumn).Expression;
            expr = expr.Replace("\n"," ").Replace("\r"," ").Replace("\t"," ");
            sbMain.Append(modelName + '\t' + tableName + '\t' + objectType + '\t' + objectName + '\t' + hiddenFlag + '\t' + objectDesc + '\t' + objectDF + '\t' + expr + '\t' + formatStr + '\t' + dataType + newline);        
        }
        else
        {
            sbMain.Append(modelName + '\t' + tableName + '\t' + objectType + '\t' + objectName + '\t' + hiddenFlag + '\t' + objectDesc + '\t' + objectDF + '\t' + "***N/A***" + '\t' + formatStr + '\t' + dataType + newline); 
        }
    }
    
    // Measures
    foreach (var o in t.Measures.OrderBy(a => a.Name).ToList())
    {
        string objectName = o.Name;
        string objectDesc = (o.Description ?? "").Replace("'","''");
        string objectDF = o.DisplayFolder ?? "";
        objectType = "Measure";
        expr = o.Expression;
        formatStr = (o.FormatString ?? "").Replace("\t"," ");
        
        expr = expr.Replace("\n"," ").Replace("\r"," ").Replace("\t"," ");
        
        if (o.IsHidden)
        {
            hiddenFlag = "Yes";
        }
        else
        {
            hiddenFlag = "No";
        }
        
        sbMain.Append(modelName + '\t' + tableName + '\t' + objectType + '\t' + objectName + '\t' + hiddenFlag + '\t' + objectDesc + '\t' + objectDF + '\t' + expr + '\t' + formatStr + '\t' + dataType + newline);
    }
    
    // Hierarchies
    foreach (var o in t.Hierarchies.OrderBy(a => a.Name).ToList())
    {
        string objectName = o.Name;
        string objectDesc = (o.Description ?? "").Replace("'","''");
        string objectDF = o.DisplayFolder ?? "";
        objectType = "Hierarchy";
        
        // Build hierarchy levels
        var levels = string.Join(" > ", o.Levels.Select(l => l.Name));
        expr = "Levels: " + levels;
        
        if (o.IsHidden)
        {
            hiddenFlag = "Yes";
        }
        else
        {
            hiddenFlag = "No";
        }
        
        sbMain.Append(modelName + '\t' + tableName + '\t' + objectType + '\t' + objectName + '\t' + hiddenFlag + '\t' + objectDesc + '\t' + objectDF + '\t' + expr + '\t' + formatStr + '\t' + dataType + newline);
    }
}

// Calculation Groups
foreach (var o in Model.CalculationGroups.ToList())
{
    string tableName = o.Name;
    string tableDesc = (o.Description ?? "").Replace("'","''");
    string hiddenFlag;
    string objectType = "Calculation Group";
    string formatStr = "";
    string dataType = "";
    
    if (o.IsHidden)
    {
        hiddenFlag = "Yes";
    }
    else
    {
        hiddenFlag = "No";
    }
    
    sbMain.Append(modelName + '\t' + tableName + '\t' + objectType + '\t' + tableName + '\t' + hiddenFlag + '\t' + tableDesc + '\t' + "" + '\t' + "***N/A***" + '\t' + formatStr + '\t' + dataType + newline);    
    
    foreach (var i in o.CalculationItems.ToList())
    {        
        string objectName = i.Name;
        string objectDesc = (i.Description ?? "").Replace("'","''");
        string expr = i.Expression;
        objectType = "Calculation Item";
        formatStr = (i.FormatStringExpression ?? "").Replace("\t"," ");
        
        expr = expr.Replace("\n"," ").Replace("\r"," ").Replace("\t"," ");
        
        sbMain.Append(modelName + '\t' + tableName + '\t' + objectType + '\t' + objectName + '\t' + hiddenFlag + '\t' + objectDesc + '\t' + "" + '\t' + expr + '\t' + formatStr + '\t' + dataType + newline);    
    }
} 

// ============================================================================
// SHEET 2: RELATIONSHIPS
// ============================================================================
var sbRel = new System.Text.StringBuilder();
string[] colNameRel = { "From Table","From Column","To Table","To Column","Cardinality","Cross Filter","Active","Security Filtering" };
int colNameRelCount = colNameRel.Length;

// Add headers
for (int i=0; i < colNameRelCount; i++)
{
    if (i<colNameRelCount-1)
    {
        sbRel.Append(colNameRel[i] + '\t');
    }
    else
    {
        sbRel.Append(colNameRel[i] + newline);
    }
}

// Extract relationships
foreach(var r in Model.Relationships.OrderBy(a => a.FromTable.Name))
{
    string fromTable = r.FromTable.Name;
    string fromCol = r.FromColumn.Name;
    string toTable = r.ToTable.Name;
    string toCol = r.ToColumn.Name;
    string card = r.FromCardinality.ToString() + ":" + r.ToCardinality.ToString();
    string cross = r.CrossFilteringBehavior.ToString();
    string active = r.IsActive.ToString();
    string security = r.SecurityFilteringBehavior.ToString();
    
    sbRel.Append(fromTable + '\t' + fromCol + '\t' + toTable + '\t' + toCol + '\t' + card + '\t' + cross + '\t' + active + '\t' + security + newline);
}

// ============================================================================
// SHEET 3: RLS RULES
// ============================================================================
var sbRls = new System.Text.StringBuilder();
string[] colNameRls = { "Role","Table","Filter Expression","Description" };
int colNameRlsCount = colNameRls.Length;

// Add headers
for (int i=0; i < colNameRlsCount; i++)
{
    if (i<colNameRlsCount-1)
    {
        sbRls.Append(colNameRls[i] + '\t');
    }
    else
    {
        sbRls.Append(colNameRls[i] + newline);
    }
}

// Extract RLS rules
foreach(var role in Model.Roles.OrderBy(a => a.Name))
{
    string roleName = role.Name;
    string roleDesc = (role.Description ?? "").Replace("'","''");
    
    if (role.TablePermissions.Count == 0)
    {
        sbRls.Append(roleName + '\t' + "***No table permissions***" + '\t' + "" + '\t' + roleDesc + newline);
    }
    else
    {
        foreach(var perm in role.TablePermissions.OrderBy(a => a.Table.Name))
        {
            string table = perm.Table.Name;
            string filter = (perm.FilterExpression ?? "").Replace("\n"," ").Replace("\r"," ").Replace("\t"," ");
            
            sbRls.Append(roleName + '\t' + table + '\t' + filter + '\t' + roleDesc + newline);
        }
    }
}

// ============================================================================
// SHEET 4: DATA SOURCES
// ============================================================================
var sbDs = new System.Text.StringBuilder();
string[] colNameDs = { "Data Source","Type","Connection String","Description" };
int colNameDsCount = colNameDs.Length;

// Add headers
for (int i=0; i < colNameDsCount; i++)
{
    if (i<colNameDsCount-1)
    {
        sbDs.Append(colNameDs[i] + '\t');
    }
    else
    {
        sbDs.Append(colNameDs[i] + newline);
    }
}

// Extract data sources
foreach(var ds in Model.DataSources.OrderBy(a => a.Name))
{
    string dsName = ds.Name;
    string dsType = ds.Type.ToString();
    string connStr = "";
    string dsDesc = (ds.Description ?? "").Replace("'","''");
    
    if (ds.Type.ToString() == "Structured")
    {
        var sds = ds as StructuredDataSource;
        connStr = "Protocol: " + (sds.Protocol ?? "") + " | Path: " + (sds.Path ?? "");
    }
    else
    {
        var lds = ds as ProviderDataSource;
        connStr = (lds.ConnectionString ?? "").Replace("\t"," ");
    }
    
    sbDs.Append(dsName + '\t' + dsType + '\t' + connStr + '\t' + dsDesc + newline);
}

// ============================================================================
// SHEET 5: TABLE PARTITIONS & STORAGE MODES
// ============================================================================
var sbPart = new System.Text.StringBuilder();
string[] colNamePart = { "Table","Partition Name","Mode","Source Type","Query/Expression" };
int colNamePartCount = colNamePart.Length;

// Add headers
for (int i=0; i < colNamePartCount; i++)
{
    if (i<colNamePartCount-1)
    {
        sbPart.Append(colNamePart[i] + '\t');
    }
    else
    {
        sbPart.Append(colNamePart[i] + newline);
    }
}

// Extract partition info
foreach(var t in Model.Tables.Where(a => a.ObjectType.ToString() != "CalculationGroupTable").OrderBy(a => a.Name))
{
    foreach(var p in t.Partitions)
    {
        string tableName = t.Name;
        string partName = p.Name;
        string mode = p.Mode.ToString();
        string sourceType = p.SourceType.ToString();
        string query = "";
        
        if (p is Partition)
        {
            query = ((Partition)p).Query ?? "";
        }
        else if (p is MPartition)
        {
            query = ((MPartition)p).Expression ?? "";
        }
        
        query = query.Replace("\n"," ").Replace("\r"," ").Replace("\t"," ");
        
        // Truncate very long queries
        if (query.Length > 500)
        {
            query = query.Substring(0, 497) + "...";
        }
        
        sbPart.Append(tableName + '\t' + partName + '\t' + mode + '\t' + sourceType + '\t' + query + newline);
    }
}

// ============================================================================
// CREATE EXCEL FILE
// ============================================================================

// Delete existing files
try
{
    File.Delete(textFilePath);
    File.Delete(excelFilePath);
}
catch
{
}

// Create combined text file (will be parsed into sheets)
var sbCombined = new System.Text.StringBuilder();
sbCombined.Append("SHEET:Model Objects" + newline);
sbCombined.Append(sbMain.ToString());
sbCombined.Append(newline + "SHEET:Relationships" + newline);
sbCombined.Append(sbRel.ToString());
sbCombined.Append(newline + "SHEET:RLS Rules" + newline);
sbCombined.Append(sbRls.ToString());
sbCombined.Append(newline + "SHEET:Data Sources" + newline);
sbCombined.Append(sbDs.ToString());
sbCombined.Append(newline + "SHEET:Partitions" + newline);
sbCombined.Append(sbPart.ToString());

SaveFile(textFilePath, sbCombined.ToString());

// Create Excel workbook with multiple sheets
var excelApp = new Excel.Application();
excelApp.Visible = false;
excelApp.DisplayAlerts = false;

var wb = excelApp.Workbooks.Add();

// Create 5 sheets
string[] sheetNames = { "Model Objects", "Relationships", "RLS Rules", "Data Sources", "Partitions" };
string[] sheetData = { sbMain.ToString(), sbRel.ToString(), sbRls.ToString(), sbDs.ToString(), sbPart.ToString() };

for (int s = 0; s < 5; s++)
{
    Excel.Worksheet ws;
    
    // Use existing sheets or add new ones
    if (s < wb.Worksheets.Count)
    {
        ws = wb.Worksheets[s + 1] as Excel.Worksheet;
    }
    else
    {
        ws = wb.Worksheets.Add() as Excel.Worksheet;
    }
    
    ws.Name = sheetNames[s];
    
    // Parse tab-delimited data
    var lines = sheetData[s].Split(new[] { newline }, StringSplitOptions.RemoveEmptyEntries);
    
    for (int row = 0; row < lines.Length; row++)
    {
        var cols = lines[row].Split('\t');
        
        for (int col = 0; col < cols.Length; col++)
        {
            ws.Cells[row + 1, col + 1] = cols[col];
        }
        
        // Format header row
        if (row == 0)
        {
            var headerRange = ws.Range[ws.Cells[1, 1], ws.Cells[1, cols.Length]];
            headerRange.Font.Bold = true;
            headerRange.Interior.Color = Excel.XlRgbColor.rgbLightGray;
        }
    }
    
    // Auto-fit columns
    ws.Columns.AutoFit();
    
    // Freeze header row
    ws.Range["A2"].Select();
    excelApp.ActiveWindow.FreezePanes = true;
}

// Delete extra default sheets if any
while (wb.Worksheets.Count > 5)
{
    ((Excel.Worksheet)wb.Worksheets[wb.Worksheets.Count]).Delete();
}

// Save workbook
wb.SaveAs(excelFilePath, Excel.XlFileFormat.xlWorkbookDefault);

// Close and cleanup
wb.Close();
excelApp.Quit();
System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

// Delete temp text file
try
{
    File.Delete(textFilePath);
}
catch
{
}

// Show success message
var summary = "Complete Data Dictionary exported to: " + excelFilePath + newline + newline;
summary += "Sheet 1 - Model Objects: " + Model.Tables.Sum(t => t.Columns.Count + t.Measures.Count + t.Hierarchies.Count + 1) + " objects" + newline;
summary += "Sheet 2 - Relationships: " + Model.Relationships.Count + " relationships" + newline;
summary += "Sheet 3 - RLS Rules: " + Model.Roles.Count + " roles" + newline;
summary += "Sheet 4 - Data Sources: " + Model.DataSources.Count + " data sources" + newline;
summary += "Sheet 5 - Partitions: " + Model.Tables.Sum(t => t.Partitions.Count) + " partitions";

Info(summary);
