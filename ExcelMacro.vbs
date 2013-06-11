Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open("C:\ACL DATA\ACLContest\Macro.xlsm")
Set objWorkbook = objExcel.Workbooks.Open("C:\ACL DATA\ACLContest\Excel_File.xlsx")

objExcel.Application.Run "Macro.xlsm!FormatTable"
objExcel.ActiveWorkbook.Save
objExcel.ActiveWorkbook.Close

objExcel.Application.Quit

WScript.Echo "Finished."
WScript.Quit