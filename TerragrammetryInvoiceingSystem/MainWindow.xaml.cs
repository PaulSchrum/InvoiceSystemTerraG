using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Invoicing;
using TimeAnalyzerino;

namespace TerragrammetryInvoiceingSystem
{
   /// <summary>
   /// Interaction logic for MainWindow.xaml
   /// </summary>
   public partial class MainWindow : Window
   {
      public MainWindow()
      {
         InitializeComponent();
         this.TimesheetName.Text = @"C:\Users\Paul\Documents\RM21\Expenses\RM21 Paul Schrum time and expenses.xlsm";
         this.DestinationFolder.Text = @"C:\Users\Paul\Documents\Life\Business\Terragrammetry\Invoicing";
         this.seedFileFullName = @"C:\Users\Paul\Documents\Life\Business\Terragrammetry\Invoicing\Invoice Seed.xlsx";

         setComboBoxValues();
      }

      private TSanalyst analyst { get; set; }
      private InvoiceSummary invoiceSummary { get; set; }
      private String seedFileFullName { get; set; }
      private List<String> invoicableJobNumbers { get; set; }

      private void setComboBoxValues()
      {
         analyst = new TSanalyst(this.TimesheetName.Text);
         invoiceSummary = null;
         this.invoicableJobNumbers = analyst
            .GetAllProjectNumbersWhichMayNowBeInvoiced()
            .Select(anInteger => anInteger.ToString()).ToList();
         this.cmb_invoiceableProjects.ItemsSource = this.invoicableJobNumbers;
         if (this.invoicableJobNumbers.Count > 0)
         {
            this.cmb_invoiceableProjects.SelectedItem = this.invoicableJobNumbers.FirstOrDefault();
            this.btn_generateInvoice.IsEnabled = true;
         }
         else
         {
            this.cmb_invoiceableProjects.SelectedItem = "No Jobs Can be Invoiced.";
            this.btn_generateInvoice.IsEnabled = false;
         }
      }

      private void btn_generateInvoice_Click(object sender, RoutedEventArgs e)
      {
         int job = Convert.ToInt32(this.cmb_invoiceableProjects.SelectedItem);
         this.invoiceSummary = InvoiceSummary.Create(analyst, job);
         this.invoiceSummary.SaveAsNewExcelFile
            (seedFileFullName, this.DestinationFolder.Text);
         setComboBoxValues();
      }
   }
}
