using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;

namespace BasicSettingsMVC.Models
{
    public partial class Usr
    {
        public int ID { get; set; }
        public string WechatID { get; set; }
        public string Code { get; set; }
        public string Name { get; set; }
        public string Desc { get; set; }
        public string Tel { get; set; }
        public string Tel1 { get; set; }
        public string Mobile { get; set; }
        public string Mobile1 { get; set; }
        public string Addr { get; set; }
        public string Addr1 { get; set; }
        public int DepartmentID { get; set; }
        public int? VendorID { get; set; }
        public int? RoleID { get; set; }
        public bool Disable { get; set; }

        public string QuoteCommit { get; set; }
        public string QuoteDetailRead { get; set; }
        public string QuoteAudit { get; set; }
        public string QuoteAudit2 { get; set; }
        public string PurchaceCommit { get; set; }
        public string PurchaceAudit { get; set; }
        public string PurchaceAudit2 { get; set; }
        public string PurchaceAudit3 { get; set; }
        public string ChargeBackCommit { get; set; }
        public string ChargeBackAudit { get; set; }
        public string DepotAdmin { get; set; }
        public string ReportExport { get; set; }

    }
}
