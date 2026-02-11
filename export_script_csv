// Export all measures to CSV
var sb = new System.Text.StringBuilder();
sb.AppendLine("Table,Measure,DAX Expression,Description");

// Complete model export to separate CSV sections
var sb = new System.Text.StringBuilder();

// 1. MEASURES
sb.AppendLine("=== MEASURES ===");
sb.AppendLine("Table,Measure,DAX,Description,FormatString");
foreach(var m in Model.AllMeasures)
{
    var table = m.Table.Name;
    var name = m.Name;
    var dax = m.Expression.Replace("\n", " ").Replace("\r", " ").Replace("\"", "\"\"");
    var desc = (m.Description ?? "").Replace("\"", "\"\"");
    var fmt = (m.FormatString ?? "").Replace("\"", "\"\"");
    
    var line = "\"" + table + "\",\"" + name + "\",\"" + dax + "\",\"" + desc + "\",\"" + fmt + "\"";
    sb.AppendLine(line);
}

// 2. CALCULATED COLUMNS
sb.AppendLine("");
sb.AppendLine("=== CALCULATED COLUMNS ===");
sb.AppendLine("Table,Column,DAX,DataType");
foreach(var t in Model.Tables)
{
    foreach(var c in t.Columns)
    {
        if(c is CalculatedColumn)
        {
            var table = t.Name;
            var colName = c.Name;
            var dax = ((CalculatedColumn)c).Expression.Replace("\n", " ").Replace("\r", " ").Replace("\"", "\"\"");
            var dtype = c.DataType.ToString();
            
            var line = "\"" + table + "\",\"" + colName + "\",\"" + dax + "\",\"" + dtype + "\"";
            sb.AppendLine(line);
        }
    }
}

// 3. CALCULATED TABLES
sb.AppendLine("");
sb.AppendLine("=== CALCULATED TABLES ===");
sb.AppendLine("Table,DAX");
foreach(var t in Model.Tables)
{
    if(t is CalculatedTable)
    {
        var table = t.Name;
        var dax = ((CalculatedTable)t).Expression.Replace("\n", " ").Replace("\r", " ").Replace("\"", "\"\"");
        
        var line = "\"" + table + "\",\"" + dax + "\"";
        sb.AppendLine(line);
    }
}

// 4. RELATIONSHIPS
sb.AppendLine("");
sb.AppendLine("=== RELATIONSHIPS ===");
sb.AppendLine("FromTable,FromColumn,ToTable,ToColumn,Cardinality,CrossFilter,Active");
foreach(var r in Model.Relationships)
{
    var fromTable = r.FromTable.Name;
    var fromCol = r.FromColumn.Name;
    var toTable = r.ToTable.Name;
    var toCol = r.ToColumn.Name;
    var card = r.FromCardinality.ToString() + ":" + r.ToCardinality.ToString();
    var cross = r.CrossFilteringBehavior.ToString();
    var active = r.IsActive.ToString();
    
    var line = "\"" + fromTable + "\",\"" + fromCol + "\",\"" + toTable + "\",\"" + toCol + "\",\"" + card + "\",\"" + cross + "\",\"" + active + "\"";
    sb.AppendLine(line);
}

// 5. RLS RULES
sb.AppendLine("");
sb.AppendLine("=== RLS RULES ===");
sb.AppendLine("Role,Table,FilterExpression");
foreach(var role in Model.Roles)
{
    foreach(var perm in role.TablePermissions)
    {
        var roleName = role.Name;
        var table = perm.Table.Name;
        var filter = (perm.FilterExpression ?? "").Replace("\n", " ").Replace("\r", " ").Replace("\"", "\"\"");
        
        var line = "\"" + roleName + "\",\"" + table + "\",\"" + filter + "\"";
        sb.AppendLine(line);
    }
}

// Save to file
var outputPath = @"C:\temp\complete_model.csv";
System.IO.File.WriteAllText(outputPath, sb.ToString());

// Copy to clipboard
System.Windows.Forms.Clipboard.SetText(sb.ToString());

// Show result
var msg = "Exported complete model to " + outputPath;
msg = msg + "\n- Measures: " + Model.AllMeasures.Count();
msg = msg + "\n- Relationships: " + Model.Relationships.Count();
msg = msg + "\n- Roles: " + Model.Roles.Count();
Info(msg);
