using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;


namespace TimeAnalyzerino
{
   public class TimeSheetRow
   {
      public TimeSheetRow
         ( DateTime workDate
         , DateTime workBegin
         , DateTime workEnd
         , TimeSpan deductions
         , TimeSpan total
         , TimeSpan weekTotal
         , String jobNumber
         , String task
         , String description
         , String notChargeable
         , String comments
         , String invoicable
         , String invoice
         , int rowInSheet
         )
      {
         WorkDate = workDate;
         WorkBegin = workBegin;
         WorkEnd = workEnd;
         Deductions = deductions;
         Total = total;
         WeekTotal = weekTotal;
         JobNumber = jobNumber;
         getPartsFromJobNumber();
         Task = task;
         Description = description;
         NotChargeable = notChargeable;
         Comments = comments;
         Invoicable = invoicable;
         Invoice = invoice;
         RowInSheet = rowInSheet;
         MaxRowInSheet = Math.Max(RowInSheet, maxRowInSheet_);
      }

      public TimeSheetRow(ExcelWorksheet ws, int row)
      {
         WorkDate = convertCellToDateTime(ws.Cells[row, 1]);
         WorkBegin = convertCellToDateTime(ws.Cells[row, 2]);
         WorkEnd = convertCellToDateTime(ws.Cells[row, 3]);
         Deductions = convertCellToTimeSpan(ws.Cells[row, 4]);
         Total = convertCellToTimeSpan(ws.Cells[row, 5]);
         WeekTotal = convertCellToTimeSpan(ws.Cells[row, 6]);
         JobNumber = convertCellToString(ws.Cells[row, 7]);
         getPartsFromJobNumber();
         Task = convertCellToString(ws.Cells[row, 8]);
         Description = convertCellToString(ws.Cells[row, 9]);
         NotChargeable = convertCellToString(ws.Cells[row, 10]);
         Comments = convertCellToString(ws.Cells[row, 11]);
         Invoicable = convertCellToString(ws.Cells[row, 12]);
         Invoice = convertCellToString(ws.Cells[row, 13]);
         RowInSheet = row;
         MaxRowInSheet = Math.Max(RowInSheet, maxRowInSheet_);
      }

      public DateTime WorkDate {get; internal set;}
      public DateTime WorkBegin {get; internal set;}
      public DateTime WorkEnd {get; internal set;}
      public TimeSpan Deductions {get; internal set;}
      public TimeSpan Total {get; internal set;}
      public TimeSpan WeekTotal {get; internal set;}
      public String JobNumber {get; internal set;}
      private int jobNumberIntegerPart_;
      public int JobNumberIntegerPart { get { return jobNumberIntegerPart_; } }
      public String JobNumberDecimalPart { get; internal set; }
      public String Task {get; internal set;}
      public String Description {get; internal set;}
      public String NotChargeable {get; internal set;}
      public String Comments {get; internal set;}
      public String Invoicable {get; internal set;}
      public String Invoice { get; internal set; }
      public int RowInSheet { get; set; }

      private DateTime convertCellToDateTime(ExcelRange cell)
      {
         if (null == cell) return new DateTime(0L);
         if (null == cell.Value) return new DateTime(0L);
         if(String.IsNullOrEmpty(cell.Value.ToString()))
            return new DateTime(0L);

         DateTime returnValue;
         var cellContents = cell.Value.ToString();
         bool parsed;
         // DateTime.TryParse won't work, so I made my own
         parsed = parseDateTime(cellContents, out returnValue);
         if(false == parsed)
         {
            long longValue = convertDecimalDaysStringToTick(cellContents);
            returnValue = new DateTime(longValue);
         }
         return returnValue;
      }

      private bool parseDateTime(String strg, out DateTime outVal)
      {
         outVal = new DateTime();
         if (true == String.IsNullOrEmpty(strg)) return false;
         var yearStr = strg.Split(' ');
         if (yearStr.Length < 1) return false;
         var dateStr = yearStr[0].Split('/');
         if (dateStr.Length != 3) return false;
         int month; int day; int year;
         bool successState = true;

         successState |= Int32.TryParse(dateStr[0], out month);
         successState |= Int32.TryParse(dateStr[1], out day);
         successState |= Int32.TryParse(dateStr[2], out year);

         if (true == successState)
            outVal = new DateTime(year, month, day);

         return successState;
      }

      private long convertDecimalDaysStringToTick(String ddays)
      {
         Double asDouble;
         bool parsed = Double.TryParse(ddays, out asDouble);
         if (false == parsed)
            return 0L;
         return Convert.ToInt64(asDouble * 24 * 3600 * 10000000);
      }

      private TimeSpan convertCellToTimeSpan(ExcelRange cell)
      {
         if (null == cell.Value)
            return new TimeSpan(0L);

         var cellContents = cell.Value.ToString();
         long timeSpanAsTicks = convertDecimalDaysStringToTick(cellContents);
         return new TimeSpan(timeSpanAsTicks);
         //return TimeSpan.ParseExact
      }

      private String convertCellToString(ExcelRange cell)
      {
         if (null == cell) return String.Empty;
         if (String.IsNullOrEmpty(cell.Text)) return String.Empty;
         return cell.Text.ToString();
      }

      private void getPartsFromJobNumber()
      {
         if (true == String.IsNullOrEmpty(this.JobNumber)) return;
         var jobnum = this.JobNumber.Split('.');
         Int32.TryParse(this.JobNumber.Split('.').FirstOrDefault(), out jobNumberIntegerPart_);
         if (jobnum.Length > 1)
            this.JobNumberDecimalPart = jobnum[1];
      }

      private static int maxRowInSheet_ = 0;
      public static int MaxRowInSheet 
      {
         get { return maxRowInSheet_; }
         protected set { maxRowInSheet_ = value; }
      }

      public static bool HasData(ExcelWorksheet ws, int row)
      {
         return (null != ws.Cells[row, 1].Value);
      }
   }
}
