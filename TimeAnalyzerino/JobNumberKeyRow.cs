using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace TimeAnalyzerino
{
   public class JobNumberKeyRow : JobNumberKeyBaseRow
   {
      public JobNumberKeyRow(ExcelWorksheet ws, int row)
         : base(ws, row, 1)
      {
         Invoiceable = convertCellToString(ws.Cells[row, 4]);
         Comments = convertCellToString(ws.Cells[row, 5]);
      }
      public String Invoiceable { get; internal set; }
      public String Comments { get; internal set; }

   }
}
