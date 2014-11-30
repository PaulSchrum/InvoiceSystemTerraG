using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

using OfficeOpenXml;
using TimeAnalyzerino;


[assembly: InternalsVisibleTo("UnitTestTimeAnalyzer")]
namespace Invoicing
{
   public class InvoiceSummary
   {
      private InvoiceSummary() { }

      public List<InvoiceDay> InvoiceDays { get; protected set; }
      public InvoicingRow InvoicingRow { get; protected set; }
      public CompaniesRow Addressee { get; protected set; }
      public int JobNumber { get; protected set; }
      public int OrderNumber { get; protected set; }
      protected TSanalyst analyst { get; set; }
      public String FileName { get; protected set; }
      public String FullFileName { get; protected set; }
      public bool IsIntermediate { get; set; }
      public bool TestingMode { get; set; }
      private ExcelPackage xlPackage {get; set;}
      private ExcelWorkbook xlWorkBook { get; set; }
      private ExcelWorksheet XLTimeSheet { get; set; }

      protected String FirstWordOfCompanyName;

      private void generateFileName()
      {
         var mostRecentInvoice = this.analyst.allInvoicingRows
            .Where(row => row.Value.JobNumber == this.JobNumber)
            .Select(row => row.Value)
            .OrderBy(row => row.InvoiceOrderNumber)
            .LastOrDefault();

         this.OrderNumber = mostRecentInvoice.InvoiceOrderNumber + 1;
         try { FirstWordOfCompanyName = this.Addressee.CompanyName.Split(' ')[0] + " "; }
         catch (Exception e) { FirstWordOfCompanyName = "Unnamed "; }
         this.setFileName();
      }

      private void setFileName()
      {
         this.FileName =
            FirstWordOfCompanyName +
            JobNumber + "." +
            OrderNumber.ToString("D4") + ".xlsx";
      }

      public void SaveAsNewExcelFile(String seedFileNameOnly, String workingPath)
      {
         String seedFile = workingPath + @"\" + seedFileNameOnly;
         FullFileName = workingPath + @"\" + this.FileName;
         bool successfulWrite = false;
         while(!successfulWrite)
         {
            if(File.Exists(FullFileName))
            {
               this.incrementOrderNumber();
               FullFileName = workingPath + @"\" + this.FileName;
            }
            else
            {
               File.Copy(seedFile, FullFileName);
               successfulWrite = true;
            }
         }

         xlPackage = xlPackage = new ExcelPackage(new FileInfo(this.FullFileName));
         xlWorkBook = xlPackage.Workbook;
         XLTimeSheet = xlWorkBook.Worksheets["Service Invoice"];
         XLTimeSheet.Cells["B7"].Value = this.Addressee.ContactPerson;
         XLTimeSheet.Cells["B8"].Value = this.Addressee.CompanyName;
         XLTimeSheet.Cells["B9"].Value = this.Addressee.Address1;
         XLTimeSheet.Cells["B10"].Value = this.Addressee.Address2;
         XLTimeSheet.Cells["B11"].Value = this.Addressee.CityStateZip;
         XLTimeSheet.Cells["E8"].Value = this.getInvoiceStartDate();
         XLTimeSheet.Cells["F8"].Value = this.getInvoiceEndDate();
         XLTimeSheet.Cells["E11"].Value = DateTime.Today;
         XLTimeSheet.Cells["F11"].Value = DateTime.Today.AddDays(17.0);
         XLTimeSheet.Cells["B11"].Value = this.Addressee.CityStateZip;
         XLTimeSheet.Cells["F4"].Value = 
            this.JobNumber.ToString() + "." + this.OrderNumber.ToString("D4");
         XLTimeSheet.Cells["F5"].Value = this.Addressee.JobNumber.ToString();

         if (this.IsIntermediate)
            XLTimeSheet.Cells["E3"].Value = "Intermediate Invoice";

         int startDataRow = 14;
         int nextDataRow = startDataRow;
         foreach(var day in this.InvoiceDays)
         {
            day.WriteToExcelWorksheet(XLTimeSheet, ref nextDataRow);
         }

         foreach(int row in Enumerable.Range(startDataRow, nextDataRow - startDataRow + 2))
         {
            if(row % 2 == 0)
            {
              XLTimeSheet.ForTableRow(row, 1, 5,
                  cell => cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White)
                  );
            }
            else
            {
               XLTimeSheet.ForTableRow(row, 1, 5,
                  cell => cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.WhiteSmoke)
                  );
            }
         }

         XLTimeSheet.DeleteRow(nextDataRow - 1, 2);

         xlPackage.Save();

         var rowToWriteTo = analyst.allInvoicingRows.Count + 3;
         analyst.WriteToNextAvailableRow(
            new InvoicingRow
               ( this.JobNumber
               , this.OrderNumber
               , this.getInvoiceStartDate()
               , this.getInvoiceEndDate()
               , TimeSpan.FromHours(this.getBillableHours())
               , this.getHourlyRate()
               , this.getBilledAmount()
               )
            );
      }

      private DateTime getInvoiceStartDate()
      {
         return this.InvoiceDays
            .OrderBy(day => day.Date_)
            .FirstOrDefault().Date_;
      }

      private DateTime getInvoiceEndDate()
      {
         return this.InvoiceDays
            .OrderBy(day => day.Date_)
            .LastOrDefault().Date_;
      }

      private Double getBillableHours()
      {
         return (Double) this.InvoiceDays
            .Sum(day => 
               day.JobNumberSummaries
                  .Sum(jobForDay => 
                     jobForDay.HoursWorked))
            ;
      }

      private Double getHourlyRate()
      {
         return
            (Double)this.InvoiceDays.First().JobNumberSummaries.First().HourlyRate;
      }

      private Double getBilledAmount()
      {
         return getBillableHours() * getHourlyRate();
      }

      private void incrementOrderNumber()
      {
         this.OrderNumber++;
         this.setFileName();
      }

      internal void deleteInvoiceXLfile_forTestingCleanup()
      {
         if (File.Exists(this.FullFileName))
            File.Delete(this.FullFileName);
      }

      public static InvoiceSummary Create(TSanalyst analyst, int jobNumber)
      {
         var allInvoiceableRows =
            analyst.GetAllInvoicableRowsNotYetInvoiced(jobNumber);

         var InvoiceDays = InvoiceDay.CreateList(allInvoiceableRows);

         if (null == InvoiceDays || InvoiceDays.Count == 0) 
            return null;

         var returnValue = new InvoiceSummary();
         returnValue.InvoiceDays = InvoiceDays;
         returnValue.InvoicingRow = new InvoicingRow();
         returnValue.InvoicingRow.Comment = String.Empty;
         returnValue.InvoicingRow.DatePaymentDeposited = default(DateTime);
         returnValue.InvoicingRow.DatePaymentReceived = default(DateTime);
         returnValue.InvoicingRow.DateSent = default(DateTime);
         returnValue.InvoicingRow.AmountPayed = 0.0;

         returnValue.InvoicingRow.StartDate = InvoiceDays.First().Date_;
         returnValue.InvoicingRow.EndDate = InvoiceDays.Last().Date_;
         returnValue.InvoicingRow.JobNumber = 
            InvoiceDays.First().JobNumberSummaries.First().JobNumberInteger;
         returnValue.InvoicingRow.BillableHours = TimeSpan.FromHours(
            (Double) InvoiceDays.Sum(jobSummary => jobSummary.HoursWorked));
         returnValue.InvoicingRow.BilledAmount = 
            (Double) InvoiceDays.Sum(jobSummary => jobSummary.Pay);

         returnValue.Addressee = analyst.allCompanies
            .Where(row => row.Value.JobNumber == jobNumber)
            .OrderBy(row => row.Value.StartDate)
            .FirstOrDefault().Value
            ;

         returnValue.JobNumber = jobNumber;
         returnValue.analyst = analyst;
         returnValue.IsIntermediate = false;
         returnValue.generateFileName();

         return returnValue;
      }

   }
}
