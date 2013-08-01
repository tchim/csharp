using System;
using System.Collections.Generic;
using System.Data; //DataTable
using System.Text; //StringBuilder
public static class CSV_Converter
{
#region Constants
private const string _COMMA_SEPARATOR = ",";
#endregion
//-------------------------------------------------------------------------------------------------
private static string Generate_ArgumentException_Text(string p_Method_Name, string p_Input_Parameter_Name)
{
string s_Message = "Abort " + p_Method_Name + ". Input " + p_Input_Parameter_Name + " cannot be null.";
return s_Message;
}
//-------------------------------------------------------------------------------------------------
public static bool DataTable_Contains_All_ColumnNames(DataTable p_DataTable, string[] p_ColumnNames)
{
const string s_MethodName = "DataTable Contains All ColumnNames";
if (p_DataTable == null)
{
throw new ArgumentException(Generate_ArgumentException_Text(s_MethodName, "DataTable"));
}
if (p_ColumnNames == null)
{
throw new ArgumentException(Generate_ArgumentException_Text(s_MethodName, "ColumnNames"));
}
foreach (string s_columnName in p_ColumnNames)
{
if (!p_DataTable.Columns.Contains(s_columnName))
{
return false;
}
}
return true;
}
//-------------------------------------------------------------------------------------------------
public static string Convert_DataTable_All_Columns_To_CSV(DataTable p_DataTable)
{
const string s_MethodName = "Convert DataTable To CSV";
if (p_DataTable == null)
{
throw new ArgumentException(Generate_ArgumentException_Text(s_MethodName, "DataTable"));
}
string[] s_ColumnNames_To_Convert = new string[p_DataTable.Columns.Count];
foreach (DataColumn dc in p_DataTable.Columns)
{
s_ColumnNames_To_Convert[dc.Ordinal] = dc.ColumnName;
}
return Convert_DataTable_Columns_To_CSV(p_DataTable, s_ColumnNames_To_Convert);
}
//-------------------------------------------------------------------------------------------------
public static string Convert_DataTable_Columns_To_CSV(DataTable p_DataTable, string[] p_ColumnNames_To_Convert)
{
const string s_MethodName = "Convert DataTable To CSV";
//Validate inputs before processing
if (p_DataTable == null)
{
throw new ArgumentException(Generate_ArgumentException_Text(s_MethodName, "DataTable"));
}
if (p_ColumnNames_To_Convert == null)
{
throw new ArgumentException(Generate_ArgumentException_Text(s_MethodName, "ColumnNames_To_Convert"));
}
if (!DataTable_Contains_All_ColumnNames(p_DataTable, p_ColumnNames_To_Convert))
{
StringBuilder sb_Message = new StringBuilder();
sb_Message.Append("Abort ");
sb_Message.Append(s_MethodName);
sb_Message.Append(". ");
sb_Message.Append("At least one of the ColumnNames is not contained in the DataTable");
throw new ArgumentException(sb_Message.ToString());
}
StringBuilder sb_OutputFileData = new System.Text.StringBuilder();
foreach (DataRow dr in p_DataTable.Rows)
{
bool b_this_is_first_item_flag = true;
foreach (string s_column in p_ColumnNames_To_Convert)
{
//Append a separator ONLY if this is not the first item in the data row
if (b_this_is_first_item_flag)
{
b_this_is_first_item_flag = false;
}
else
{
sb_OutputFileData.Append(_COMMA_SEPARATOR);
}
string s_temp = dr[s_column].ToString();
sb_OutputFileData.Append(s_temp);
}
sb_OutputFileData.Append("\n");
}
return sb_OutputFileData.ToString();
}
//-------------------------------------------------------------------------------------------------
}