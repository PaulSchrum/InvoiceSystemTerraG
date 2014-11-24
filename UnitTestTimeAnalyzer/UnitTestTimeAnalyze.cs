using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using TimeAnalyzerino;

namespace UnitTestTimeAnalyzer
{
   [TestClass]
   public class UnitTestTimeAnalyze
   {
      private String XLfilename = @"C:\SourceModules\InvoiceSystemTerraG\TestData\TestDataset.xlsm";
      private TSanalyst analyst = null;

      [TestInitialize]
      public void TimeAnalyzerSetup()
      {
         if (null != analyst) return;
         analyst = new TSanalyst(XLfilename);
      }

      [TestCleanup]
      public void TimeAnalyzerCleanup()
      {
      }

      [TestMethod]
      public void TimeAnalyzer_Create_IsNotNull()
      {
         TimeAnalyzerSetup();
         Assert.IsNotNull(analyst);
      }

      [TestMethod]
      public void TimeAnalyzer_Create_OpensXLFileAndReadsTimesheetWorksheet()
     { 
         TimeAnalyzerSetup();
         Assert.IsNotNull(analyst.XLTimeSheet);
      }

      [TestMethod]
      public void TimeAnalyzer_TimesheetWorksheet_Has5758Rows()
      {
         TimeAnalyzerSetup();
         var rows = analyst.XLTimeSheet.Dimension.End.Row;
         Assert.AreEqual(expected: 5758, actual: rows);
      }

   }
}
