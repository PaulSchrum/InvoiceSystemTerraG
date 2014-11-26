using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TimeAnalyzerino;

namespace Invoicing
{
   public class InvoiceDay
   {
      protected InvoiceDay(List<JobNumberSummary> allTasksForDay)
      {
         this.JobNumberSummaries = allTasksForDay;
      }

      public DateTime Date_ { get { return this.JobNumberSummaries.First().Date_; } }
      public Decimal Pay { get { return this.JobNumberSummaries.Sum(row => row.PayForThisDay); } }
      public Decimal HoursWorked { get { return this.JobNumberSummaries.Sum(row => row.HoursWorked); } }
      public List<JobNumberSummary> JobNumberSummaries { get; protected set; }

      internal static List<InvoiceDay> CreateList(
         IEnumerable<TimeSheetRow> allInvoiceableRows)
      {
         if (null == allInvoiceableRows) return null;
         List<InvoiceDay> returnList = new List<InvoiceDay>();

         var tsRowsGroupedByDate = allInvoiceableRows
            .GroupBy(row => row.WorkDate)
            .OrderBy(grp => grp.Key)
            ;

         foreach(var grp in tsRowsGroupedByDate)
         {
            var v = grp.AsEnumerable()
               .GroupBy(vVal => vVal.JobNumber);

            List<JobNumberSummary> allTasksForDay = JobNumberSummary.CreateList(v);
            returnList.Add(new InvoiceDay(allTasksForDay));
         }

         return returnList.OrderBy(row => row.Date_).ToList();
      }
   }
}
