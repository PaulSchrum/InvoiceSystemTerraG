using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace TimeAnalyzerino
{
   public class JobNumberKeyRow : GeneralWorksheetRow
   {

      public JobNumberKeyRow(ExcelWorksheet ws, int row)
         : base(ws, row)
      {
         JobNumber = convertCellToString(ws.Cells[row, 7]);
         getPartsFromJobNumber();
         Task = convertCellToString(ws.Cells[row, 8]);
         Description = convertCellToString(ws.Cells[row, 9]);
      }

      public String JobNumber { get; internal set; }
      protected int jobNumberIntegerPart_;
      public int JobNumberIntegerPart { get { return jobNumberIntegerPart_; } }
      public String JobNumberDecimalPart { get; internal set; }
      public String Task { get; internal set; }
      public String Description { get; internal set; }

      protected void getPartsFromJobNumber()
      {
         if (true == String.IsNullOrEmpty(this.JobNumber)) return;
         var jobnum = this.JobNumber.Split('.');
         Int32.TryParse(this.JobNumber.Split('.').FirstOrDefault(), out jobNumberIntegerPart_);
         if (jobnum.Length > 1)
            this.JobNumberDecimalPart = jobnum[1];
      }


   }
}
