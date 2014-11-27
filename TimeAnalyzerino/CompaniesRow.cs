using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace TimeAnalyzerino
{
   public class CompaniesRow : GeneralWorksheetRow
   {
      public CompaniesRow(ExcelWorksheet ws, int row)
         :base(ws, row)
      {
         JobNumber = convertCellToInt(ws.Cells[row, 1]);
         StartDate = convertCellToDateTime(ws.Cells[row, 2]);
         CompanyName = convertCellToString(ws.Cells[row, 3]);
         ContactPerson = convertCellToString(ws.Cells[row, 4]);
         Address1 = convertCellToString(ws.Cells[row, 5]);
         Address2 = convertCellToString(ws.Cells[row, 6]);
         CityStateZip = convertCellToString(ws.Cells[row, 7]);
      }

      public int JobNumber {get; protected set;}
      public DateTime StartDate {get; protected set;}
      public String CompanyName {get; protected set;}
      public String ContactPerson {get; protected set;}
      public String Address1 {get; protected set;}
      public String Address2 {get; protected set;}
      public String CityStateZip { get; protected set; }

   }
}
