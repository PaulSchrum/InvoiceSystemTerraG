namespace TimeSheetAnalyser

open Microsoft.Office.Interop

type TimesheetAnalyst(filename) =
   member this.xlApp = new Excel.ApplicationClass()
   member this.workbook = this.xlApp.Workbooks.Open(filename)
   member this.worksheet = this.workbook.Worksheets.["Timesheet"]

