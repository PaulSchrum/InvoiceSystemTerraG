using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TimeAnalyzerino;
using Invoicing;
using System.IO;

namespace UnitTestTimeAnalyzer
{
   [TestClass]
   public class UnitTestTimeAnalyze
   {
      private String XLfilename = @"C:\SourceModules\InvoiceSystemTerraG\TestData\TestDataset.xlsm";
      private TSanalyst analyst = null;
      private String invoiceDirectory =
         @"C:\SourceModules\InvoiceSystemTerraG\TestData\Invoicing";
      private String invoiceSeedXLfile = "Invoice Seed.xlsx";
      private bool SeedRelatedTestingMayProceed { get; set; }

      [TestInitialize]
      public void TimeAnalyzerSetup()
      {
         if (null != analyst) return;
         analyst = new TSanalyst(XLfilename);

         var rm21Dev = analyst.allJobNumberKeyRows
            .Where(row => row.Value.JobNumber.Equals("21.01"))
            .FirstOrDefault().Value;
         if (null != rm21Dev)
            rm21Dev.testing_changeInvoicableValueTo("y");

         SeedRelatedTestingMayProceed =
            isSeedFileAvailable();
      }

      private bool isSeedFileAvailable()
      {
         bool result = false;

         if (Directory.Exists(this.invoiceDirectory) &&
            File.Exists(this.invoiceDirectory + @"\"
               + this.invoiceSeedXLfile))
         {
            result = true;
         }

         return result;
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
         int jobNumberIP = analyst.allTimesheetRows[1144].JobNumberInteger;
         Assert.AreEqual(
            expected: 1100,
            actual: jobNumberIP);
      }

      [TestMethod]
      public void TimeAnalyzer_TimesheetWorksheet_Row1144JobDecimal_Is_1()
      {
         TimeAnalyzerSetup();
         String jobNumberDec = analyst.allTimesheetRows[1144].JobSubnumber;
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
      public void TimeAnalyzer_TimesheetWorksheet_GetAllTimesheetRowsByJobOverDateRange_Gets6Rows()
      {
         TimeAnalyzerSetup();
         var v = analyst.GetTimesheetRowsByJobOverDateRange(
            1100,
            new DateTime(2014, 7, 24),
            new DateTime(2014, 7, 27)
            )
            .ToList()
            ;

         Assert.AreEqual(expected: 6, actual: v.Count);
      }

      [TestMethod]
      public void TimeAnalyzer_TimesheetWorksheet_GetTimesheetRowsByInvoiceableJobOverDateRange_Gets3Rows()
      {
         TimeAnalyzerSetup();
         var v = analyst.GetTimesheetRowsByInvoiceableJobOverDateRange(
            1100,
            new DateTime(2014, 7, 24),
            new DateTime(2014, 7, 28)
            )
            .ToList()
            ;

         Assert.AreEqual(expected: 3, actual: v.Count);
      }

      [TestMethod]
      public void TimeAnalyzer_CompaniesCollection_Has2EntriesFor2014()
      {
         TimeAnalyzerSetup();
         var companies = analyst.allCompanies
            .Where(row => row.Value.StartDate < new DateTime(2014, 12, 31))
            .ToList();
         Assert.AreEqual(expected: 2, actual: companies.Count);
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

      [TestMethod]
      public void TimeAnalyzer_JobNumberKeyWorksheet_Job1100_Has8InvoiceableRows()
      {
         TimeAnalyzerSetup();
         var invoiceables = analyst.GetInvoiceableJobsByJobNumber(1100).ToList();
         Assert.AreEqual(
            expected: 8,
            actual: invoiceables.Count);
      }

      [TestMethod]
      public void TimeAnalyzer_InvoicingWorksheet_Has3Rows()
      {
         TimeAnalyzerSetup();
         var previousInvoices = analyst.allInvoicingRows.ToList();
         Assert.AreEqual(
            expected: 3,
            actual: previousInvoices.Count);
      }

      [TestMethod]
      public void TimeAnalyzer_InvoicingWorksheet_Job1200_MostRecentInvoiceWas25August2014()
      {
         TimeAnalyzerSetup();
         var lastInvoiceDate = analyst.GetDateOfLastInvoiceSent(1200);
         DateTime Expected = new DateTime(2014, 8, 25);
         Assert.AreEqual(
            expected: Expected,
            actual: lastInvoiceDate);
      }

      [TestMethod]
      public void TimeAnalyzer_InvoicingWorksheet_NonExistantJobNumber_ReturnsDefault()
      {
         TimeAnalyzerSetup();
         var lastInvoiceDate = analyst.GetDateOfLastInvoiceSent(73);
         DateTime Expected = default(DateTime);
         Assert.AreEqual(
            expected: Expected,
            actual: lastInvoiceDate);
      }

      [TestMethod]
      public void TimeAnalyzer_GetAllInvoicableProjectNumberIntegers_Returns3()
      {
         TimeAnalyzerSetup();
         var numberOfInvoicableProjects =
            analyst.GetAllInvoicableProjectNumbersIntegerParts()
            .ToList();
         Assert.AreEqual(
            expected: 3,
            actual: numberOfInvoicableProjects.Count);
      }

      [TestMethod]
      public void TimeAnalyzer_GetAllInvoicableProjectNumbers_Returns12()
      {
         TimeAnalyzerSetup();
         var numberOfInvoicableProjects =
            analyst.GetAllInvoicableProjectNumbers(0)
            .ToList().Count;
         Assert.AreEqual(
            expected: 12,
            actual: numberOfInvoicableProjects);
      }

      [TestMethod]
      public void TimeAnalyzer_GetAllInvoicableProjectNumberFor1200_Returns3()
      {
         TimeAnalyzerSetup();
         var numberOfInvoicableProjects =
            analyst.GetAllInvoicableProjectNumbers(1200)
            .ToList().Count;
         Assert.AreEqual(
            expected: 3,
            actual: numberOfInvoicableProjects);
      }

      [TestMethod]
      public void TimeAnalyzer_GetLastActiveDateForAllInvoicableJobNumbers_Has3Rows()
      {
         TimeAnalyzerSetup();
         var rowsOfLastDateWorked =
            analyst.GetLastWorkedRowForEachBillableProjectNumber()
            .ToList();
         Assert.AreEqual(
            expected: 3,
            actual: rowsOfLastDateWorked.Count);
      }

      [TestMethod]
      public void TimeAnalyzer_GetLastActiveInvoicableDateForJobNumber1200_IsAugust25_2014()
      {
         TimeAnalyzerSetup();
         var numberOfInvoicableProjects =
            analyst.GetLastWorkedRowForEachBillableProjectNumber()
            .Where(row => row.JobNumberInteger == 1200)
            .Select(row => row)
            .FirstOrDefault()
            ;

         var Expected = new DateTime(2014, 8, 25);
         Assert.AreEqual(expected: 1200, actual: numberOfInvoicableProjects.JobNumberInteger);
         Assert.AreEqual(
            expected: Expected,
            actual: numberOfInvoicableProjects.WorkDate);
      }

      [TestMethod]
      public void TimeAnalyzer_GetAllInvoicableProjectsNeverInvoiced_Returns1()
      {
         TimeAnalyzerSetup();
         var neverInvoicedProjects =
            analyst.GetAllProjectNumbersWhichHaveNeverBeenInvoiced()
            .ToList();
         Assert.AreEqual(
            expected: 1,
            actual: neverInvoicedProjects.Count);
      }

      [TestMethod]
      public void TimeAnalyzer_GetMostRecentInvoiceForEachProjectEverInvoiced_ReturnsCorrect()
      {
         TimeAnalyzerSetup();
         var mostRecentInvoiceForEveryProjectEverInvoiced =
            analyst.GetMostRecentInvoiceForAllProjectsEverInvoiced()
            .ToList();

         var value = mostRecentInvoiceForEveryProjectEverInvoiced;
         bool success =
            (value.Count == 2)
            && value[0].JobNumber == 1100
            && value[0].DateSent == new DateTime(2014, 11, 14)
            && value[1].JobNumber == 1200
            && value[1].DateSent == new DateTime(2014, 8, 25)
            ;

         Assert.IsTrue(success);
      }

      [TestMethod]
      public void TimeAnalyzer_GetAllInvoicableProjectsThatCouldBeInvoiced_Returns2()
      {
         TimeAnalyzerSetup();
         var couldBeInvoiced =
            analyst.GetAllProjectNumbersWhichMayNowBeInvoiced()
            .ToList();
         Assert.AreEqual(
            expected: 2,
            actual: couldBeInvoiced.Count);
      }

      [TestMethod]
      public void TimeAnalyzer_GetAllTimesheetRows_For1200_Returns9()
      {
         TimeAnalyzerSetup();
         var tsRowsFor1200 =
            analyst.GetAllBillableTimeSheetRowsByProject()
            .Where(grp => grp.Key == 1200)
            .ToList();
         Assert.IsNotNull(tsRowsFor1200);
         Assert.AreEqual(
            expected: 1,
            actual: tsRowsFor1200.Count);
         var values = tsRowsFor1200[0];
         Assert.AreEqual(
            expected: 9
            , actual: values.ToList().Count);
      }

      [TestMethod]
      public void TimeAnalyzer_GetMostRecentInvoice_For21_ReturnsNull()
      {
         TimeAnalyzerSetup();
         var mostRecentInvoiceFor21 =
            analyst.GetMostRecentInvoiceForProject(21);
         Assert.IsNull(mostRecentInvoiceFor21);
      }

      [TestMethod]
      public void TimeAnalyzer_GetAllInvoiceableTSRow_For21_ReturnsCorrect()
      {
         TimeAnalyzerSetup();
         var invoicableTSRowsFor21 =
            analyst.GetTimesheetAllRowsByInvoiceableJob(21);
         Assert.IsNotNull(invoicableTSRowsFor21);
         Assert.IsTrue(invoicableTSRowsFor21.ToList().Count > 0);
      }

      [TestMethod]
      public void TimeAnalyzer_GetMostRecentInvoice_For1200_Returns20140821()
      {
         TimeAnalyzerSetup();
         var mostRecentInvoiceFor1200 =
            analyst.GetMostRecentInvoiceForProject(1200);
         Assert.IsNotNull(mostRecentInvoiceFor1200);
         DateTime expected = new DateTime(2014, 8, 21);
         Assert.AreEqual(
            expected: expected,
            actual: mostRecentInvoiceFor1200.EndDate
            );
      }

      [TestMethod]
      public void TimeAnalyzer_GetAllInvoiceableTSrows_For21_Returns20140607Row()
      {
         TimeAnalyzerSetup();
         var invoicableTSrowsFor21 =
            analyst.GetAllInvoicableRowsNotYetInvoiced(21)
            .LastOrDefault();
         Assert.IsNotNull(invoicableTSrowsFor21);
         DateTime expected = new DateTime(2014, 6, 7);
         Assert.AreEqual(
            expected: expected,
            actual: invoicableTSrowsFor21.WorkDate
            );
      }

      [TestMethod]
      public void TimeAnalyzer_GetAllInvoiceableTSrows_For1100_ReturnsMoreThan2()
      {
         TimeAnalyzerSetup();
         var invoicableTSrowsFor1100 =
            analyst.GetAllInvoicableRowsNotYetInvoiced(1100)
            .ToList();
         Assert.IsNotNull(invoicableTSrowsFor1100);
         Assert.IsTrue
            (invoicableTSrowsFor1100.Count > 2
            );
      }

      [TestMethod]
      public void InvoiceSummary_CreateForJob12863384_ReturnsNull()
      {
         TimeAnalyzerSetup();
         var v = InvoiceSummary.Create(analyst, 12863384);
         Assert.IsNull(v);
      }

      [TestMethod]
      public void InvoiceSummary_CreateForJob1100_ReturnsNotNull()
      {
         TimeAnalyzerSetup();
         var invoice = InvoiceSummary.Create(analyst, 1100);
         Assert.IsNotNull(invoice);
      }

      [TestMethod]
      public void InvoiceSummary_CreateForJob1100_ReturnsPayeeLobbyGuard()
      {
         TimeAnalyzerSetup();
         var invoice = InvoiceSummary.Create(analyst, 1100);
         Assert.IsNotNull(invoice);
         invoice.TestingMode = true;
         invoice.IsIntermediate = true;
         Assert.AreEqual(
            expected: "LobbyGuard Solutions, LLC",
            actual: invoice.Addressee.CompanyName
            );
      }

      private void SkipTestIfNeeded_SeedInvoiceFile()
      {
         if (false == SeedRelatedTestingMayProceed)
         {
            Assert.Inconclusive(
               "Unable to run tests related to file creation." + "\n"
               + "A required directory or file does not exist."
               );
         }
      }

      [TestMethod]
      public void InvoiceSummary_CreateXLFileForJob1100_ReturnsCreateSuccessful()
      {
         TimeAnalyzerSetup();
         SkipTestIfNeeded_SeedInvoiceFile();
         var newInvoice = InvoiceSummary.Create(analyst, 1100);
         Assert.IsNotNull(newInvoice);
      }

      [TestMethod]
      public void InvoiceSummary_CreateXLFileForJob1100_FileNameIsCorrect()
      {
         TimeAnalyzerSetup();
         SkipTestIfNeeded_SeedInvoiceFile();
         var newInvoice = InvoiceSummary.Create(analyst, 1100);
         Assert.AreEqual(
            expected: "LobbyGuard 1100.0003.xlsx",
            actual: newInvoice.FileName
            );
      }

      [TestMethod]
      public void InvoiceSummary_CreateXLFileForJob1100_FileExists()
      {
         TimeAnalyzerSetup();
         SkipTestIfNeeded_SeedInvoiceFile();
         var newInvoice = InvoiceSummary.Create(analyst, 1100);
         try
         {
            newInvoice.SaveAsNewExcelFile
               (invoiceSeedXLfile
               , invoiceDirectory
               );
            Assert.IsTrue(
               File.Exists(newInvoice.FullFileName)
               );
         }
         finally
         {
            newInvoice.deleteInvoiceXLfile_forTestingCleanup();
         }
      }

   }
}
