namespace TimesheetAnalyst

open System
open System.IO
open FSharpx.TypeProviders

type TimesheetRow = ExcelProvider<"C:\SourceModules\InvoiceSystemTerraG\TestData\TestDataset.xlsm", "Timesheet", true>

