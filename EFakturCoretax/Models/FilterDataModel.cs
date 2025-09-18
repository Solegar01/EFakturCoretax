using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EFakturCoretax.Models
{
    public class FilterDataModel
    {
        public string DocEntry { get; set; }
        public string DocNo { get; set; }
        public string CardCode { get; set; }
        public string CardName { get; set; }
        public string ObjType { get; set; }
        public string ObjName { get; set; }
        public string PostDate { get; set; }
        public string BranchCode { get; set; }
        public string BranchName { get; set; }
        public string OutletCode { get; set; }
        public string OutletName { get; set; }
        public bool Selected { get; set; } = true;
    }
}
