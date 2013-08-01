using System;
using System.Web; //needed for HttpContext
using System.Web.UI.WebControls; //needed for GridView
using OfficeOpenXml; // namespace for the ExcelPackage assembly
using System.IO; //needed for FileInfo
using System.Data; //needed for DataTable
/// <summary>
/// Summary description for Excel_Utility
/// This class is intended for use with ExcelPackage.dll
/// http://excelpackage.codeplex.com/
///
/// ExcelPackage.dll is a GPL-licensed code that allows export GridViews into native .xlsx
/// without requiring some version of MS Office 2007 installed on the server.
/// </summary>
public class ExcelPackage_Utility
{
#region Constants
private const string _DEFAULT_EXCEL_OUTPUTFILE_EXTENSION = ".xlsx";
private const string _DEFAULT_SHEETNAME = "sheet1";
#endregion
//-------------------------------------------------------------------------------------------------
private static string Generate_ArgumentException_Text(string p_Method_Name, string p_Input_Parameter_Name)
{
string s_Message = "Abort " + p_Method_Name + ". Input " + p_Input_Parameter_Name + " cannot be null.";
return s_Message;
}
//------------------------------------------------------------------------------------------
public static byte[] Generate_XLSX_As_Byte_Array(DataTable p_DataTable, string p_sheet_name)
{
const string s_MethodName = "Generate_XLSX_As_Byte_Array";
//May not have a null input DataTable
if (p_DataTable == null)
{
throw new ArgumentException(Generate_ArgumentException_Text(s_MethodName, "DataTable"));
}
//If p_sheet_name then use the default sheet name defined in the constant
string s_sheet_name = string.IsNullOrEmpty(p_sheet_name) ? _DEFAULT_SHEETNAME : p_sheet_name;
ExcelPackage xlPackage = new ExcelPackage();
//-----------------------------------------------------------
//Begin Excel Sheet Creation
//Create a new WorkSheet
ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets.Add(s_sheet_name);
const int _starting_row = 1;
const int _starting_column = 1;
int n_excel_row = _starting_row;
int n_excel_column = _starting_column;
//Create the worksheet header
foreach (DataColumn dc in p_DataTable.Columns) //Creating Headings
{
var cell = worksheet.Cells[n_excel_row, n_excel_column];
//Setting Value in cell
cell.Value = dc.ColumnName;
//Format Column
ExcelColumn column = worksheet.Column(n_excel_column);
column.AutoFit();
string s_DataType_Name = dc.DataType.Name;
switch (s_DataType_Name)
{
case "DateTime":
column.Style.Numberformat.Format = "yyyy-mm-dd";
break;
default:
break;
}
n_excel_column++;
}
//Fill up the rows below the header using two nested loops.
//Outer loop iterates down the rows
foreach (DataRow dr in p_DataTable.Rows) // Adding Data into rows
{
n_excel_row++;
n_excel_column = _starting_column;
//Inner loop iterate across a row starting at column 1
foreach (DataColumn dc in p_DataTable.Columns)
{
var cell = worksheet.Cells[n_excel_row, n_excel_column];
//Setting Value in cell
cell.Value = dr[dc.ColumnName]; ;
n_excel_column++;
}
}
//End Excel Sheet Creation
//-----------------------------------------------------------
//Return
return xlPackage.GetAsByteArray();
}
//------------------------------------------------------------------------------------------
}