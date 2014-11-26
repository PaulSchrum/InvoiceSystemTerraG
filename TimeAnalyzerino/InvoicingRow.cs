using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace TimeAnalyzerino
{
   public class InvoicingRow : GeneralWorksheetRow
   {
      public InvoicingRow(ExcelWorksheet ws, int row)
         :base(ws, row)
      {
         InvoiceNumber = convertCellToString(ws.Cells[row, 1]);
         JobNumber = convertCellToInt(ws.Cells[row, 2]);
         StartDate = convertCellToDateTime(ws.Cells[row, 3]);
         EndDate = convertCellToDateTime(ws.Cells[row, 4]);
         BillableHours = convertCellToTimeSpan(ws.Cells[row, 5]);
         HourlyRate = convertCellToDouble(ws.Cells[row, 6]);
         BilledAmount = convertCellToDouble(ws.Cells[row, 7]);
         DateSent = convertCellToDateTime(ws.Cells[row, 8]);
         DatePaymentReceived = convertCellToDateTime(ws.Cells[row, 9]);
         AmountPayed = convertCellToDouble(ws.Cells[row, 10]);
         DatePaymentDeposited = convertCellToDateTime(ws.Cells[row, 9]);
         Comment = convertCellToString(ws.Cells[row, 9]);
      }

      public InvoicingRow()
      {
         base.RowInSheet = -1;
         
      }

      public String InvoiceNumber {get; set;}
      public int JobNumber {get; set;}
      public DateTime StartDate {get; set;}
      public DateTime EndDate {get; set;}
      public TimeSpan BillableHours {get; set;}
      public Double HourlyRate {get; set;}
      public Double BilledAmount {get; set;}
      public DateTime DateSent {get; set;}
      public DateTime DatePaymentReceived {get; set;}
      public Double AmountPayed {get; set;}
      public DateTime DatePaymentDeposited {get; set;}
      public String Comment { get; set; }


   }
}
