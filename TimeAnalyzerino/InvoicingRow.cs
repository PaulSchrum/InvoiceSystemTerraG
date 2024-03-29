﻿using System;
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
         InvoiceOrderNumber = determineInvoiceOrderNumber();
      }

      public InvoicingRow()
      {
         base.RowInSheet = -1;
         
      }

      public InvoicingRow
         ( int jobNumber
         , int invoiceOrderNumber
         , DateTime startDate
         , DateTime endDate
         , TimeSpan billableHours
         , Double hourlyRate
         , Double billedAmount
         )
         : base()
      {
         JobNumber = jobNumber;
         StartDate = startDate;
         EndDate = endDate;
         BillableHours = billableHours;
         HourlyRate = hourlyRate;
         BilledAmount = billedAmount;
         InvoiceOrderNumber = invoiceOrderNumber;
         InvoiceNumber = jobNumber.ToString() + "." +
            InvoiceOrderNumber.ToString("D4");
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
      public int InvoiceOrderNumber { get; set; }

      public override void WriteToRow(ExcelWorksheet ws, int row)
      {
         if (Util.AreAnyNull(ws))
            return;
         if (row < 1)
            return;
         ws.Cells[row, 1].Value = this.InvoiceNumber;
         ws.Cells[row, 2].Value = this.JobNumber;
         ws.Cells[row, 3].Value = this.StartDate;
         ws.Cells[row, 4].Value = this.EndDate;
         ws.Cells[row, 5].Value = this.BillableHours;
         ws.Cells[row, 6].Value = this.HourlyRate;
         ws.Cells[row, 7].Value = this.BilledAmount;
      }

      private int determineInvoiceOrderNumber()
      {
         var strs = InvoiceNumber.Split('.');
         if(strs.Length > 1)
         {
            return Convert.ToInt32(strs[1]);
         }
         return 0;
      }

   }
}
