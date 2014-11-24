using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using InvoiceSystemTerraG;
using InvoiceSystemTG;

namespace InvoiceApp.ViewModel
{
    public class InvoiceViewModel : INotifyPropertyChanged, IDisposable
    {
        private InvoiceMaker model_invoiceMaker;
        private List<Timesheet> ts;
        public ObservableCollection<InvoiceSummary> InvoiceableSummary { get; set; }

        public InvoiceViewModel()
        {
            model_invoiceMaker = new InvoiceMaker(@"C:\Users\Paul\Documents\RM21\Expenses\RM21 Paul Schrum time and expenses.xlsm");
        }

        public void Dispose()
        {
            if (null == model_invoiceMaker) return;
            model_invoiceMaker.Dispose();
        }

        public void GetInvoicableSummary(
            DateTime FirstDate, DateTime LastDate)
        {
            InvoiceableSummary = model_invoiceMaker.GetAllInvoicableProjects(FirstDate, LastDate);
        }

        private DateTime firstDate;
        public DateTime FirstDate
        {
            get { return firstDate; }
            set { firstDate = value; RaisePropertyChanged("FirstDate"); }
        }

        private DateTime lastDate;
        public DateTime LastDate
        {
            get { return lastDate; }
            set { lastDate = value; RaisePropertyChanged("LastDate"); }
        }

        private ObservableCollection<InvoiceSummary> invoiceables;

        public void QueryInvoicables()
        {

        }

        public void GenerateInvoices()
        {

        }

        public event PropertyChangedEventHandler PropertyChanged;
        public void RaisePropertyChanged(String prop)
        {
            if (null != PropertyChanged)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
            }
        }

    }
}
