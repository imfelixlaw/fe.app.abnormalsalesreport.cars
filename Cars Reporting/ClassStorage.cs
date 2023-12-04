using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Cars_Reporting
{
    class jstorage
    {
        public string jobno { get; set; }
        public string site { get; set; }
    }

    class jsdata
    {
        public string Centre { get; set; }
        public string Site { get; set; }
        public string ReceiptNo { get; set; }
        public string Date { get; set; }
        public string CarNo { get; set; }
        public string TelNo { get; set; }
        public string Total { get; set; }
    }

    class jscontent
    {
        public string Service { get; set; }
        public string Amount { get; set; }
    }

    class sccontent
    {
        public string scode { get; set; }
        public string sname { get; set; }
    }

    class ssite
    {
        public string site { get; set; }
        public string sitename { get; set; }
    }

    class sproduct
    {
        public string pcode { get; set; }
        public string pname { get; set; }
    }
}
