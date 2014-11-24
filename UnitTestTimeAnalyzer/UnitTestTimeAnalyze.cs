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

      [TestMethod]
      public void TimeAnalyzer_TimesheetWorksheet_Row1144Task_Is_SchrumAssignment02()
      {
         TimeAnalyzerSetup();
         var softwareTaskName = analyst.allTimesheetRows[1144].Task;
         Assert.AreEqual(
            expected: "SchrumAssignment02",
            actual: softwareTaskName);
      }

      [TestMethod]
      public void TimeAnalyzer_TimesheetWorksheet_Row1144JobInteger_Is_1100()
      {
         TimeAnalyzerSetup();
         int jobNumberIP = analyst.allTimesheetRows[1144].JobNumberIntegerPart;
         Assert.AreEqual(
            expected: 1100,
            actual: jobNumberIP);
      }

      [TestMethod]
      public void TimeAnalyzer_TimesheetWorksheet_Row1144JobDecimal_Is_1()
      {
         TimeAnalyzerSetup();
         String jobNumberDec = analyst.allTimesheetRows[1144].JobNumberDecimalPart;
         Assert.AreEqual(
            expected: "1",
            actual: jobNumberDec);
      }

      [TestMethod]
      public void TimeAnalyzer_TimesheetWorksheet_GetJobsByDateRange_Returns4Jobs()
      {
         TimeAnalyzerSetup();
         var v = analyst.GetJobsByDateRange(
            new DateTime(2014, 8, 6),
            new DateTime(2014, 8, 18)
            );
         Assert.AreEqual(expected: 4, actual: v.Count);
      }

      [TestMethod]
      public void TimeAnalyzer_TimesheetWorksheet_GetJobsByDateRange_Returns9RowsForJob1100()
      {
         TimeAnalyzerSetup();
         var v = analyst.GetJobsByDateRange(
            new DateTime(2014, 8, 6),
            new DateTime(2014, 8, 18)
            );
         Assert.AreEqual(expected: 9, actual: v[1100].Count);
      }

      [TestMethod]
      public void TimeAnalyzer_JobNumberKeyWorksheet_Row7_Has_Description_RM21()
      {
         TimeAnalyzerSetup();
         String description = analyst.allJobNumberKeyRows[7].Description;
         Assert.AreEqual(
            expected: "RM21",
            actual: description);
      }

   }
}
