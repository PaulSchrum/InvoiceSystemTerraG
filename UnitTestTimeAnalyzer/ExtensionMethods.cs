using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace UnitTestTimeAnalyzer
{
   internal static class ExtensionMethods
   {
      private static String filePathAndName { get; set; }
      private static String worksheetName { get; set; }
      private static ExcelPackage xlPackage { get; set; }
      private static ExcelWorksheet XLWorkSheet { get; set; }

      private static void OpenFileAndWorksheetIfNecessary(String xlPathAndName, String xlWorksheetName)
      {
         if(null == filePathAndName
            || !(filePathAndName.Equals(xlPathAndName)))
         {
            var fileInfo = new FileInfo(xlPathAndName);
            try
            {
               xlPackage = new ExcelPackage(fileInfo);
               worksheetName = String.Empty;
            }
            catch(IOException ex)
            {
               throw new Exception("That excel file is probably open.  Close it or share it and try again.", ex);
            }
         }
         if(null == worksheetName 
            || !(worksheetName.Equals(xlWorksheetName)))
         {
            worksheetName = xlWorksheetName;
            var wb = xlPackage.Workbook;
            XLWorkSheet = wb.Worksheets[worksheetName];
         }
      }

      private static StringBuilder composeStringInCaseItsNeeded(String expected, String actual)
      {
         var sb = new StringBuilder("Expected (");
         sb.Append(expected);
         sb.Append(") does not match Actual (");
         sb.Append(actual);
         sb.Append(").");
         return sb;
      }

      public static void Dispose()
      {
         filePathAndName = null;
         worksheetName = null;
         xlPackage = null;
         XLWorkSheet = null;
      }

      private static Object GetCellAt(String xlPathAndName, String worksheetName, int row, int column)
      {
         //OpenFileAndWorksheetIfNecessary(xlPathAndName, worksheetName);
         AssertCellIsNotEmpty(xlPathAndName, worksheetName, row, column);
         return XLWorkSheet.Cells[row, column].Value;
      }

      public static void AssertCellHasValue
         ( this String fullPathAndName
         , String WorksheetName
         , int row
         , int column
         , String expectedValue
         )
      {
         var valStr = (String) GetCellAt(fullPathAndName, WorksheetName, row, column);
         var sb = composeStringInCaseItsNeeded(expectedValue, valStr);
         if (!(expectedValue.Equals(valStr))) throw new Exception(sb.ToString());
      }

      public static void AssertCellIsEmpty
         (this String fullPathAndName_
         , String worksheetName_
         , int row
         , int column
         )
      {
         OpenFileAndWorksheetIfNecessary(fullPathAndName_, worksheetName_);
         if (XLWorkSheet.Cells[row, column].Value != null)
            throw new Exception("Cell is not empty although it was expected to be empty.");
      }

      public static void AssertCellIsNotEmpty
         (this String fullPathAndName_
         , String worksheetName_
         , int row
         , int column
         )
      {
         OpenFileAndWorksheetIfNecessary(fullPathAndName_, worksheetName_);
         if (XLWorkSheet.Cells[row, column].Value == null)
            throw new Exception("Cell is empty although it was expected to be not empty.");
      }

   }
}
