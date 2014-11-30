using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OfficeOpenXml;

namespace TimeAnalyzerino
{
   public static class EPPlusExtensions
   {
      public static void ForTableRow
         ( this ExcelWorksheet ws
         , int row
         , int startColumn
         , int numberOfColumns
         , Action<ExcelRange> whatToDoToThem
         )
      {
         var aCell = ws.Cells[1,2];
         var cols = Enumerable.Range(startColumn, numberOfColumns);
         foreach(var col in Enumerable.Range(startColumn, numberOfColumns))
         {
            whatToDoToThem(ws.Cells[row, col]);
         }
      }
   }
}
