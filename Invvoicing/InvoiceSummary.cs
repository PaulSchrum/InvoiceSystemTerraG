using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TimeAnalyzerino;


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
      public bool IsIntermediate { get; set; }
      public bool TestingMode { get; set; }

      private void generateFileName()
      {
         var mostRecentInvoice = this.analyst.allInvoicingRows
            .Where(row => row.Value.JobNumber == this.JobNumber)
            .Select(row => row.Value)
            .OrderBy(row => row.InvoiceOrderNumber)
            .LastOrDefault();

         this.OrderNumber = mostRecentInvoice.InvoiceOrderNumber + 1;
         String FirstWordOfCompanyName;
         try { FirstWordOfCompanyName = this.Addressee.CompanyName.Split(' ')[0] + " "; }
         catch (Exception e) { FirstWordOfCompanyName = "Unnamed "; }
         this.FileName = 
            FirstWordOfCompanyName + " " + 
            JobNumber + "." +
            OrderNumber.ToString("D4") + ".xlsx";
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
