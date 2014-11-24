namespace InvoiceFS

open System
//open System.IO
//open FSharpx.TypeProviders

//#r "Microsoft.Office.Interop.Excel"
//#r "office"
//open Microsoft.Office.Interop

//type TimesheetRow = ExcelProvider<"C:\SourceModules\InvoiceSystemTerraG\TestData\RM21 Paul Schrum time and expenses.xlsm","Timesheet",true>
type TimesheetRow = {
   WorkDate: DateTime;
   WorkBegin: DateTime;
   WorkEnd: DateTime;
   Deductions: DateTime;
   Total: TimeSpan;
   WeekTotal: TimeSpan;
   JobNumber: String;
   Task: String;
   Description: String;
   NotChargeable: String;
   Comments: String;
   Invoicable: String;
   Invoiced: String
   }


