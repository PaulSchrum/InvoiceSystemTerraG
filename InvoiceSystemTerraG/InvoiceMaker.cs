using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using LinqToExcel;
using LinqToExcel.Query;
using LinqToExcel.Extensions;
using LinqToExcel.Domain;
using LinqToExcel.Attributes;

using NodaTime;
using System.Collections.ObjectModel;
using InvoiceSystemTerraG;

namespace InvoiceSystemTG
{
    public class InvoiceMaker : IDisposable
    {
        private String excelFileToQuery {get; set;}
        private String tempExcelFileToWorkOn { get; set; }
        private ExcelQueryFactory excelQueryFactory { get; set; }
        ExcelQueryable<Timesheet> TimesheetEQ { get; set; }
        List<Timesheet> TheTimeSheet { get; set; }
        ExcelQueryable<JobNumberKey> JobNumberKeyEQ { get; set; }
        List<JobNumberKey> TheJobNumberKey { get; set; }

        public InvoiceMaker(String ExcelFilePathName)
        {
            excelFileToQuery = ExcelFilePathName;
            generateTempFileName();
            File.Copy(excelFileToQuery, tempExcelFileToWorkOn, true);
            if (!File.Exists(tempExcelFileToWorkOn)) return;
            excelQueryFactory = new ExcelQueryFactory(tempExcelFileToWorkOn);
            TimesheetEQ = this.excelQueryFactory.Worksheet<Timesheet>("Timesheet");
            TheTimeSheet = TimesheetEQ.
                Where(row => row.Total != null).ToList();

            JobNumberKeyEQ = this.excelQueryFactory.Worksheet<JobNumberKey>("JobNumberKey");
            TheJobNumberKey = this.JobNumberKeyEQ.ToList();

            var jCount = JobNumberKeyEQ.Count();

            foreach (var row in TheTimeSheet)
            {
                row.postProcess();
            }
        }

        private void generateTempFileName()
        {
            var newFN = Path.GetFileNameWithoutExtension(excelFileToQuery) + "_temp";
            tempExcelFileToWorkOn = Path.GetDirectoryName(excelFileToQuery) + "\\" +
                                    newFN +
                                    Path.GetExtension(excelFileToQuery);
        }

        public void Dispose()
        {
            if (null == excelQueryFactory) return;
            excelQueryFactory.Dispose();
            File.Delete(tempExcelFileToWorkOn);
            excelQueryFactory = null;
        }

        public IEnumerable<Timesheet> queryAllByTask(IEnumerable<Timesheet> aCollection, String task)
        {
            IEnumerable<Timesheet> localCollection;
            if (aCollection == null)
                localCollection = this.TheTimeSheet;
            else
                localCollection = aCollection;

            List<Timesheet> retTS = new List<Timesheet>();
            foreach (var row in localCollection)
            {
                if (row.Task.NotNullAndEquals(task))
                    retTS.Add(row);
            }
            return retTS;

            //return from aRow in localCollection
            //       where aRow.Task.Equals(task)
            //       select aRow;
        }

        public ObservableCollection<InvoiceSummary> GetAllInvoicableProjects(
            DateTime FirstDate, DateTime LastDate)
        {
            ObservableCollection<InvoiceSummary> retCollection = new ObservableCollection<InvoiceSummary>();

            var timeSheetRowsInDateRange = (
                from row in TheTimeSheet
                where row.WorkDate >= FirstDate &&
                      row.WorkDate <= LastDate &&
                      row.NotChargeable == null
                select row).ToList();

            var projectSummaryPhase1 = (
                from row in timeSheetRowsInDateRange
                group row by row.JobNumber.LeftOfChar('.') into g
                select g
                ).ToList();

            var projectSummaryPhase2 = (
                from client in projectSummaryPhase1
                let clientName = JobNumberKey.ClientName(TheJobNumberKey, client.Key)
                where clientName.Length > 0
                select client
                ).ToList();

            foreach(var client in projectSummaryPhase2)
            {
                var newInvSummLine = new InvoiceSummary(
                    client.Key,
                    JobNumberKey.ClientName(TheJobNumberKey, client.Key));

                newInvSummLine.TotalTime = 0.0;
                newInvSummLine.TotalAmountDue = 0.0m;
                foreach( var workTimeRow in client)
                {
                    Double rate = 0.0; Double Hours = 0.0;
                    if (JobNumberKey.IsInvoicable(TheJobNumberKey, workTimeRow.JobNumber) == false)
                        continue;

                    rate = JobNumberKey.GetHourlyRate(TheJobNumberKey, workTimeRow.JobNumber);
                    Hours = workTimeRow.Total.HoursToDouble();
                    newInvSummLine.TotalTime += Hours;
                    newInvSummLine.TotalAmountDue += (Decimal)(Hours * rate);
                }
                retCollection.Add(newInvSummLine);
            }

            return retCollection;
        }

        public Object RunPrivateMethodTest(String methodName, Object[] arguments=null)
        {
            Object returnVal = null;

            var machineName = System.Environment.MachineName;
            var userName = System.Environment.UserName;

            var methodInfo = this.GetType().GetMethod(methodName);
            if (null == methodInfo) throw new MissingMethodException(methodName);
            returnVal = methodInfo.Invoke(this, arguments);

            return returnVal;
        }


    }

    public static class StringExtensions
    {
        public static bool NotNullAndEquals(this String s1, String s2)
        {
            if (s1 == null) return false;
            if (s2 == null) return false;
            return s1.Equals(s2);
        }

        public static String LeftOfChar(this String This, Char character)
        {
            if( !(This.Contains(character))) return String.Empty;

            String[] strArray = This.Split(character);
            if (strArray.Length < 1) return String.Empty;
            return strArray[0];
        }

        public static string GetParsed(this String This, Char character, int index)
        {
            String[] parsed = This.Split(character);
            if (index >= parsed.Length) return String.Empty;
            return parsed[index];
        }

        public static Double HoursToDouble(this String This)
        {
            if (This.Contains(":") == false) throw new FormatException(String.Format("{0} does not match expected format of hh:mm.", This));
            String[] parsed = This.Split(':');
            Double hours = Convert.ToDouble(parsed[0]);
            Double minutes=0.0;
            if (parsed.Length > 1)
                minutes = Convert.ToDouble(parsed[1]);
            return hours + minutes / 60.0;
        }
    }

    public class JobNumberKey
    {
        public String JobNumber { get; set; }
        public String Task { get; set; }
        public String Description { get; set; }
        public String Invoicable { get; set; }
        public Double HourlyRate { get; set; }
        public String Comment { get; set; }

        public static String ClientName(List<JobNumberKey> aList, String aJobNumber)
        {
            JobNumberKey theRow = null;
            try
            {
                theRow = (from row in aList
                          where row.JobNumber.LeftOfChar('.').Equals(aJobNumber)
                          select row).FirstOrDefault();
            }
            catch (Exception) { }
            if (theRow == null) return String.Empty;
            return theRow.Description.GetParsed(':', 1);
        }

        private static IEnumerable<JobNumberKey> getRowByJobNumber
            (List<JobNumberKey> aList, String aJobNumber)
        {
            return from r in aList where r.JobNumber.Equals(aJobNumber) select r;
        }

        public static bool IsInvoicable(List<JobNumberKey> aList, String aJobNumber)
        {
            var theRow = getRowByJobNumber(aList, aJobNumber);
            if (null == theRow) return false;
            String isInvable = theRow.FirstOrDefault().Invoicable;
            if (isInvable == null || isInvable.Length < 1) return false;
            return isInvable.ToUpper().Equals("Y");
        }

        public static Double GetHourlyRate(List<JobNumberKey> aList, String aJobNumber)
        {
            var theRow = getRowByJobNumber(aList, aJobNumber);
            return theRow.FirstOrDefault().HourlyRate;
        }
    }

    public class Timesheet : JobNumberKey
    {
        public DateTime WorkDate { get; set; }
        public DateTime WorkBegin { get; set; }
        public DateTime WorkEnd { get; set; }
        public String Deductions { get; set; }
        public String Total { get; set; }
        public NodaTime.Period TotalPeriod { get; set; }
        public String WeekTotal { get; set; }
        public String NotChargeable { get; set; }

        public void postProcess()
        {
            var pb = new PeriodBuilder();
            var stringParts = this.Total.Split(':');
            try
            {
                pb.Hours = Convert.ToInt64(stringParts[0]);
                pb.Minutes = Convert.ToInt64(stringParts[1]);
            }
            catch (FormatException)
            {
                return;
            }
            TotalPeriod = pb.Build();

            if (this.Task == null)
                this.Task = String.Empty;

        }
    }

}

