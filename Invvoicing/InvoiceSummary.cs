using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TimeAnalyzerino;
//using Invoicing;

namespace Invoicing
{
   public class InvoiceSummary
   {
      private InvoiceSummary() { }

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

         return returnValue;
      }

      public List<InvoiceDay> InvoiceDays { get; protected set; }
      public InvoicingRow InvoicingRow { get; protected set; }

   }
}
