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

         allRows = Enumerable.Range(2, XLTimeSheet.Dimension.End.Row)
            .Where(row => true == TimeSheetRow.HasData(XLTimeSheet, row))
            .Select(row => new TimeSheetRow(XLTimeSheet, row))
            .ToList()
            ;
      }

      private String xlPathAndName { get; set; }
      private FileInfo fileInfo {get; set;}
      private ExcelPackage xlPackage {get; set;}
      private ExcelWorkbook xlWorkBook { get; set; }
      public ExcelWorksheet XLTimeSheet { get; protected set; }
      public List<TimeSheetRow> allRows { get; protected set; }
      protected int lastDataRow {get; set;}

   }
}
