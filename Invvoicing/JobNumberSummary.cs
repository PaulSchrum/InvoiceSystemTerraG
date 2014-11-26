using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TimeAnalyzerino;

namespace Invoicing
{
    public class JobNumberSummary
    {
       public JobNumberSummary
          ( int jobNumberInt
          , String jobSubnumer
          , DateTime _date
          , Decimal hoursWorked
          , String _task
          , String description
          , Decimal hourlyRate
          , Decimal payForDay
          )
       {
          JobNumberInteger = jobNumberInt;
          JobSubnumber = jobSubnumer;
          Date_ = _date;
          HoursWorked = hoursWorked;
          Task_ = _task;
          Description = description;
          HourlyRate = hourlyRate;
          PayForThisDay = payForDay;
       }

       public JobNumberSummary(IGrouping<string, TimeSheetRow> tskGrp)
       {
          var todaysTaskGroups = tskGrp.AsEnumerable();
          var frst = todaysTaskGroups.FirstOrDefault();
          this.JobNumberInteger = frst.JobNumberInteger;
          this.JobSubnumber = frst.JobSubnumber;
          this.Date_ = frst.WorkDate;

          this.HoursWorked = todaysTaskGroups
             .Select(row => (Decimal) row.Total.TotalHours)
             .Sum();

          this.Task_ = frst.Task;
          this.Description = frst.Description;
          this.HourlyRate = 40M;
          this.PayForThisDay = this.HoursWorked * this.HourlyRate;
       }

       public int JobNumberInteger { get; protected set; }
       public String JobSubnumber { get; protected set; }
       public DateTime Date_ { get; protected set; }
       public Decimal HoursWorked { get; protected set; }
       public String Task_ { get; protected set; }  // not currently used in invoicing
       public String Description { get; protected set; }
       public Decimal HourlyRate { get; protected set; }
       public Decimal PayForThisDay { get; protected set; }

      internal static List<JobNumberSummary> CreateList(IEnumerable<IGrouping<string,TimeSheetRow>> timeSheetRows)
      {
         var returnList = new List<JobNumberSummary>();

         foreach(var tsRow in timeSheetRows)
         {
            returnList.Add(new JobNumberSummary(tsRow));
         }

         return returnList;
      }
   }
}
