using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TimeSheetAnalyser;

namespace TimeSheetAnalystTests
{
    [TestClass]
    public class UnitTestingOfTimeSheetAnalst
    {
        [TestMethod]
        public void TimesheetAnalysist_instantiatesNullWorkbok_withIncorrectPath()
        {
            TimeSheetAnalyser.TimesheetAnalyst analyst = new TimeSheetAnalyser.TimesheetAnalyst("x");
            Assert.IsNull(analyst.workbook);
        }

        [TestMethod]
        public void TimesheetAnalysist_instantiatesNonNullWorkbok_withCorrectPath()
        {
            TimeSheetAnalyser.TimesheetAnalyst analyst = new TimeSheetAnalyser.TimesheetAnalyst(@"C:\SourceModules\InvoiceSystemTerraG\TestData\RM21 Paul Schrum time and expenses.xlsm");
            Assert.IsNotNull(analyst.workbook);
        }

    }
}
