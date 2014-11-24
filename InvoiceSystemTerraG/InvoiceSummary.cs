using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InvoiceSystemTerraG
{
    public class InvoiceSummary
    {
        public String ProjectNumber { get; set; }
        public String ClientName { get; set; }
        public Double TotalTime { get; set; }
        public Decimal TotalAmountDue { get; set; }

        public InvoiceSummary(String projectNumber, String clientName)
        {
            ProjectNumber = projectNumber;
            ClientName = clientName;
        }

        public static InvoiceSummary GetByProjectNumber(IEnumerable<InvoiceSummary> aList, String projectNumber)
        {
            return aList.Where(x => x.ProjectNumber.Equals(projectNumber)).FirstOrDefault();
        }
    }
}
