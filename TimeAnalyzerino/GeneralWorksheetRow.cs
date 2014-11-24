using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace TimeAnalyzerino
{
   public class GeneralWorksheetRow
   {
      public GeneralWorksheetRow(ExcelWorksheet ws, int row)
      {
         RowInSheet = row;
         MaxRowInSheet = Math.Max(RowInSheet, maxRowInSheet_);
      }

      public int RowInSheet { get; set; }

      protected DateTime convertCellToDateTime(ExcelRange cell)
      {
         if (null == cell) return new DateTime(0L);
         if (null == cell.Value) return new DateTime(0L);
         if (String.IsNullOrEmpty(cell.Value.ToString()))
            return new DateTime(0L);

         DateTime returnValue;
         var cellContents = cell.Value.ToString();
         bool parsed;
         // DateTime.TryParse won't work, so I made my own
         parsed = parseDateTime(cellContents, out returnValue);
         if (false == parsed)
         {
            long longValue = convertDecimalDaysStringToTick(cellContents);
            returnValue = new DateTime(longValue);
         }
         return returnValue;
      }

      protected bool parseDateTime(String strg, out DateTime outVal)
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

      protected long convertDecimalDaysStringToTick(String ddays)
      {
         Double asDouble;
         bool parsed = Double.TryParse(ddays, out asDouble);
         if (false == parsed)
            return 0L;
         return Convert.ToInt64(asDouble * 24 * 3600 * 10000000);
      }

      protected TimeSpan convertCellToTimeSpan(ExcelRange cell)
      {
         if (null == cell.Value)
            return new TimeSpan(0L);

         var cellContents = cell.Value.ToString();
         long timeSpanAsTicks = convertDecimalDaysStringToTick(cellContents);
         return new TimeSpan(timeSpanAsTicks);
         //return TimeSpan.ParseExact
      }

      protected String convertCellToString(ExcelRange cell)
      {
         if (null == cell) return String.Empty;
         if (String.IsNullOrEmpty(cell.Text)) return String.Empty;
         return cell.Text.ToString();
      }

      protected static int maxRowInSheet_ = 0;
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
