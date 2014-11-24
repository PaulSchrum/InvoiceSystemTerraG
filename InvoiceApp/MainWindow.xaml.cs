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

using InvoiceApp.ViewModel;

namespace InvoiceApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private InvoiceViewModel vm;
        public MainWindow()
        {
            InitializeComponent();
            vm = new InvoiceViewModel();
        }

        private void wnd_main_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            vm.Dispose();
        }

        private void btn_getInvoicableSummary_Click(object sender, RoutedEventArgs e)
        {
            vm.GetInvoicableSummary(
                new DateTime(2014, 7, 1),
                DateTime.Now);
        }
    }
}
