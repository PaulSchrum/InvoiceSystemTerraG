using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;


namespace TimeAnalyzerino
{
    public class TimeSheetRow : JobNumberKeyRow
   {

      public TimeSheetRow(ExcelWorksheet ws, int row)
         : base(ws, row)
      {
         WorkDate = convertCellToDateTime(ws.Cells[row, 1]);
         WorkBegin = convertCellToDateTime(ws.Cells[row, 2]);
         WorkEnd = convertCellToDateTime(ws.Cells[row, 3]);
         Deductions = convertCellToTimeSpan(ws.Cells[row, 4]);
         Total = convertCellToTimeSpan(ws.Cells[row, 5]);
         WeekTotal = convertCellToTimeSpan(ws.Cells[row, 6]);
         NotChargeable = convertCellToString(ws.Cells[row, 10]);
         Comments = convertCellToString(ws.Cells[row, 11]);
         Invoicable = convertCellToString(ws.Cells[row, 12]);
         Invoice = convertCellToString(ws.Cells[row, 13]);
      }

      public DateTime WorkDate {get; internal set;}
      public DateTime WorkBegin {get; internal set;}
      public DateTime WorkEnd {get; internal set;}
      public TimeSpan Deductions {get; internal set;}
      public TimeSpan Total {get; internal set;}
      public TimeSpan WeekTotal {get; internal set;}
      public String NotChargeable {get; internal set;}
      public String Comments {get; internal set;}
      public String Invoicable {get; internal set;}
      public String Invoice { get; internal set; }

   }
}
