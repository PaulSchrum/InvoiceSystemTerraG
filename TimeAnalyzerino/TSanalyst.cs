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

         try
         {
            xlPackage = new ExcelPackage(fileInfo);
         }
         catch(IOException ex)
         {
            throw new Exception("That excel file is probably open.  Close it or share it and try again.",ex);
         }
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

         allJobNumbersEverInvoiced =
            allInvoicingRows
            .Select(row => row.Value.JobNumber)
            .Distinct()
            ;

         XLCompanies = xlWorkBook.Worksheets["Companies"];
         allCompanies = Enumerable.Range(2, XLCompanies.Dimension.End.Row)
            .Where(row => true == CompaniesRow.HasData(XLCompanies, row))
            .Select(row => new CompaniesRow(XLCompanies, row))
            .ToDictionary(row => row.RowInSheet, row => row)
            ;

      }

      private String xlPathAndName { get; set; }
      private FileInfo fileInfo {get; set;}
      private ExcelPackage xlPackage {get; set;}
      private ExcelWorkbook xlWorkBook { get; set; }
      internal ExcelWorksheet XLTimeSheet { get; set; }
      public Dictionary<int, TimeSheetRow> allTimesheetRows { get; protected set; }
      
      internal ExcelWorksheet XLJobNumberKeySheet { get; set; }
      public Dictionary<int, JobNumberKeyRow> allJobNumberKeyRows { get; protected set; }

      public ExcelWorksheet XLInvoicing { get; protected set; }
      public Dictionary<int, InvoicingRow> allInvoicingRows { get; protected set; }
      protected IEnumerable<int> allJobNumbersEverInvoiced { get; set; }

      public ExcelWorksheet XLCompanies { get; protected set; }
      public Dictionary<int, CompaniesRow> allCompanies { get; protected set; }

      public Dictionary<int, List<KeyValuePair<int,TimeSheetRow>>> GetJobsByDateRange(DateTime start, DateTime end)
      {
         return allTimesheetRows
            .Where(row => row.Value.WorkDate >= start && row.Value.WorkDate < end)
            .GroupBy(row => row.Value.JobNumberInteger)
            .OrderBy(grp => grp.Key)
            .ToDictionary(i => i.Key, i => i.ToList());
            ;
      }


      public IEnumerable<TimeSheetRow> GetTimesheetRowsByJobOverDateRange(int jobInt, DateTime start, DateTime end)
      {
         return
            allTimesheetRows
            .Where(row => row.Value.WorkDate >= start && row.Value.WorkDate < end)
            .Where(row => row.Value.JobNumberInteger == jobInt)
            .Select(row => row.Value)
            ;
      }

      public IEnumerable<JobNumberKeyRow> GetInvoiceableJobsByJobNumber(int jobNumber)
      {
         return
            allJobNumberKeyRows
            .Where(row => row.Value.JobNumberInteger == jobNumber)
            .Where(row => false == String.IsNullOrEmpty(row.Value.Invoiceable))
            .Select(row => row.Value)
            ;
      }

      public IEnumerable<TimeSheetRow> GetTimesheetAllRowsByInvoiceableJob
         (int jobInt)
      {
         var billableJobs =
            GetInvoiceableJobsByJobNumber(jobInt);

         return
            allTimesheetRows
            .Where(row => row.Value.JobNumberInteger == jobInt)
            .Select(row => row.Value)
            .Join(billableJobs
               , timeSheetRow => timeSheetRow.JobNumber
               , billableJobRow => billableJobRow.JobNumber
               , (tsh, billJ) => tsh
            );
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

      public IEnumerable<int> GetAllInvoicableProjectNumbersIntegerParts()
      {
         return
            this.allJobNumberKeyRows
            .Where(row => !(String.IsNullOrEmpty(row.Value.Invoiceable)))
            .Select(row => row.Value.JobNumberInteger)
            .Distinct()
            ;
      }

      public IEnumerable<String> GetAllInvoicableProjectNumbers(int projectNumber)
      {
         return
            this.allJobNumberKeyRows
            .Where(row =>
               {
                  if (projectNumber == 0)
                     return true;
                  else
                     return row.Value.JobNumberInteger == projectNumber;
               }
               )
            .Where(row => !(String.IsNullOrEmpty(row.Value.Invoiceable)))
            .Select(row => row.Value.JobNumber)
            ;
      }

      public IEnumerable<IGrouping<int, TimeSheetRow>>
         GetAllBillableTimeSheetRowsByProject()
      {
         var invoicableProjNums =
            GetAllInvoicableProjectNumbers(0);

         return
            this.allTimesheetRows
            .Join(
               invoicableProjNums
               , timesheetRow => timesheetRow.Value.JobNumber
               , invoicableRow => invoicableRow
               , (tsh, inv) => tsh.Value
            )
            .GroupBy(row => row.JobNumberInteger)
            ;
      }

      public IEnumerable<TimeSheetRow> GetLastWorkedRowForEachBillableProjectNumber()
      {
         return
            GetAllBillableTimeSheetRowsByProject()
            .SelectMany(group => group
               .OrderByDescending(row => row.WorkDate)
               .Take(1)
               .Select(row => row))
            ;
      }

      public IEnumerable<int> GetAllProjectNumbersWhichHaveNeverBeenInvoiced()
      {
         var neverInvoiced =
            GetLastWorkedRowForEachBillableProjectNumber()
            .Where(
               row => allJobNumbersEverInvoiced
                  .All(invJob => invJob != row.JobNumberInteger))
            .Select(row => row.JobNumberInteger)
            ;
         return neverInvoiced;
      }

      public IEnumerable<InvoicingRow> GetMostRecentInvoiceForAllProjectsEverInvoiced()
      {
         var v = this.allInvoicingRows
            .OrderByDescending(row => row.Value.EndDate)
            .GroupBy(row => row.Value.JobNumber)
            .Select(grp => grp.AsEnumerable()
               .Take(1)
               )
            .Select(kvPair => kvPair.FirstOrDefault().Value)
            //.ToList()
            ;

         return v;
      }

      public InvoicingRow GetMostRecentInvoiceForProject(int jobNumber)
      {
         InvoicingRow v =
            GetMostRecentInvoiceForAllProjectsEverInvoiced()
            .Where(invRow => invRow.JobNumber == jobNumber)
            .FirstOrDefault()
            ;
         return v;
      }

      public IEnumerable<int> GetAllProjectNumbersWhichMayNowBeInvoiced()
      {
         var monad =
            GetLastWorkedRowForEachBillableProjectNumber()
            .Where(row => 
               GetMostRecentInvoiceForAllProjectsEverInvoiced()
               .All(invRow => row.WorkDate > invRow.EndDate)
               )
            .Select(row => row.JobNumberInteger)
            .Union(GetAllProjectNumbersWhichHaveNeverBeenInvoiced())
            //.ToList()
            ;
         return monad;
      }

      public IEnumerable<TimeSheetRow> GetAllInvoicableRowsNotYetInvoiced(int jobNumber)
      {
         if (false ==
            GetAllProjectNumbersWhichMayNowBeInvoiced().Contains(jobNumber))
            return null;

         if (true == GetAllProjectNumbersWhichHaveNeverBeenInvoiced().Contains(jobNumber))
            return GetTimesheetAllRowsByInvoiceableJob(jobNumber);

         var v =
         //return 
            GetTimesheetAllRowsByInvoiceableJob(jobNumber)
            .Where(row =>
               row.WorkDate > GetMostRecentInvoiceForProject(jobNumber).EndDate)
            ;

         return v;
      }

   }
}
