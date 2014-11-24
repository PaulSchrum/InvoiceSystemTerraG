using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

using OfficeOpenXml;

namespace TimeAnalyzerino
{
   public class TSanalyst
   {
      public TSanalyst(String pathFileName)
      {
         xlPathAndName = pathFileName;
         fileInfo = new FileInfo(xlPathAndName);

         xlPackage = new ExcelPackage(fileInfo);
         xlWorkBook = xlPackage.Workbook;

         XLTimeSheet = xlWorkBook.Worksheets["Timesheet"];

         allTimesheetRows = Enumerable.Range(2, XLTimeSheet.Dimension.End.Row)
            .Where(row => true == TimeSheetRow.HasData(XLTimeSheet, row))
            .Select(row => new TimeSheetRow(XLTimeSheet, row))
            .ToDictionary(row => row.RowInSheet, row => row)
            ;

         XLJobNumberKeySheet = xlWorkBook.Worksheets["JobNumberKey"];
         allJobNumberKeyRows = Enumerable.Range(2, XLJobNumberKeySheet.Dimension.End.Row)
            .Where(row => true == JobNumberKeyRow.HasData(XLJobNumberKeySheet, row))
            .Select(row => new JobNumberKeyRow(XLJobNumberKeySheet, row))
            .ToDictionary(row => row.RowInSheet, row => row)
            ;

         XLInvoicing = xlWorkBook.Worksheets["Invoicing"];
         allInvoicingRows = Enumerable.Range(2, XLInvoicing.Dimension.End.Row)
            .Where(row => true == InvoicingRow.HasData(XLInvoicing, row))
            .Select(row => new InvoicingRow(XLInvoicing, row))
            .ToDictionary(row => row.RowInSheet, row => row)
            ;
      }

      private String xlPathAndName { get; set; }
      private FileInfo fileInfo {get; set;}
      private ExcelPackage xlPackage {get; set;}
      private ExcelWorkbook xlWorkBook { get; set; }
      public ExcelWorksheet XLTimeSheet { get; protected set; }
      public Dictionary<int, TimeSheetRow> allTimesheetRows { get; protected set; }
      
      public ExcelWorksheet XLJobNumberKeySheet { get; protected set; }
      public Dictionary<int, JobNumberKeyRow> allJobNumberKeyRows { get; protected set; }

      public ExcelWorksheet XLInvoicing { get; protected set; }
      public Dictionary<int, InvoicingRow> allInvoicingRows { get; protected set; }

      public Dictionary<int, List<KeyValuePair<int,TimeSheetRow>>> GetJobsByDateRange(DateTime start, DateTime end)
      {
         return allTimesheetRows
            .Where(row => row.Value.WorkDate >= start && row.Value.WorkDate < end)
            .GroupBy(row => row.Value.JobNumberIntegerPart)
            .OrderBy(grp => grp.Key)
            .ToDictionary(i => i.Key, i => i.ToList());
            ;
      }


      public IEnumerable<TimeSheetRow> GetTimesheetRowsByJobOverDateRange(int jobInt, DateTime start, DateTime end)
      {
         return
            allTimesheetRows
            .Where(row => row.Value.WorkDate >= start && row.Value.WorkDate < end)
            .Where(row => row.Value.JobNumberIntegerPart == jobInt)
            .Select(row => row.Value)
            ;
      }

      public IEnumerable<JobNumberKeyRow> GetInvoiceableJobsByJobNumber(int jobNumber)
      {
         return
            allJobNumberKeyRows
            .Where(row => row.Value.JobNumberIntegerPart == jobNumber)
            .Where(row => false == String.IsNullOrEmpty(row.Value.Invoiceable))
            .Select(row => row.Value)
            ;
      }

      public IEnumerable<TimeSheetRow> GetTimesheetRowsByInvoiceableJobOverDateRange
         (int jobInt, DateTime start, DateTime end)
      {
         var billableJobs =
            GetInvoiceableJobsByJobNumber(jobInt);

         return
            GetTimesheetRowsByJobOverDateRange(jobInt, start, end)
            .Join(billableJobs
               , timeSheetRow => timeSheetRow.JobNumber
               , billableJobRow => billableJobRow.JobNumber
               , (tsh, bilJ) => tsh
            );
      }

      public DateTime GetDateOfLastInvoiceSent(int jobNum)
      {
         var lastInvoiceRow =
            this.allInvoicingRows
            .Where(row => row.Value.JobNumber == jobNum)
            .OrderBy(row => row.Value.DateSent)
            .Take(1)
            .Select(row => row.Key)
            .FirstOrDefault()
            ;

         if (lastInvoiceRow == 0) return default(DateTime);

         return this.allInvoicingRows[lastInvoiceRow].DateSent;
      }
   }
}
