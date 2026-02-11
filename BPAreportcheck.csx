#r "System.IO"
using System.IO;

// ============================================================================
// CONFIGURATION
// ============================================================================
string outputPath = @"C:\Users\jharker\Documents\BPA_Report.txt";
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
// RUN BEST PRACTICE ANALYZER
// ============================================================================
var sb = new System.Text.StringBuilder();
string newline = Environment.NewLine;

// Headers
sb.Append("Model Name" + '\t' + "Rule Category" + '\t' + "Rule Name" + '\t' + "Object Name" + '\t' + "Object Type" + '\t' + "Severity" + '\t' + "Can Auto-Fix" + '\t' + "Rule ID" + '\t' + "Description" + newline);

// Get BPA analyzer
var analyzer = new TabularEditor.BestPracticeAnalyzer.Analyzer();
analyzer.SetModel(Model);

// Run analysis
var results = analyzer.AnalyzeAll().ToList();

if (results.Count == 0)
{
    Info("No BPA violations found! Model follows all best practices.");
}
else
{
    // Export results
    foreach (var result in results)
    {
        string category = result.Rule.Category ?? "";
        string ruleName = result.RuleName ?? "";
        string objectName = result.ObjectName ?? "";
        string objectType = result.ObjectType ?? "";
        string severity = result.Rule.Severity.ToString();
        string canFix = result.CanFix.ToString();
        string ruleId = result.Rule.ID ?? "";
        string description = (result.Rule.Description ?? "").Replace("\t", " ").Replace("\n", " ");
        
        sb.Append(modelName + '\t' + category + '\t' + ruleName + '\t' + objectName + '\t' + objectType + '\t' + severity + '\t' + canFix + '\t' + ruleId + '\t' + description + newline);
    }
    
    // Save to file
    System.IO.File.WriteAllText(outputPath, sb.ToString());
    
    // Summary
    var summary = "BPA Analysis Complete!" + newline + newline;
    summary += "Model: " + modelName + newline;
    summary += "Total violations: " + results.Count + newline + newline;
    summary += "By Severity:" + newline;
    summary += "- Error: " + results.Count(r => r.Rule.Severity == 1) + newline;
    summary += "- Warning: " + results.Count(r => r.Rule.Severity == 2) + newline;
    summary += "- Info: " + results.Count(r => r.Rule.Severity == 3) + newline + newline;
    summary += "Report saved to: " + outputPath;
    
    Info(summary);
}