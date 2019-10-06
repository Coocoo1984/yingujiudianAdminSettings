using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;

namespace BasicSettingsMVC.Models
{
    public partial class Usr
    {
        public long ID { get; set; }
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
        [NotMapped]
        public int DepartmentID { get; set; }
        [NotMapped]
        public int? VendorID { get; set; }
        public int? RoleID { get; set; }
        public bool Disable { get; set; }

        [NotMapped]
        public string QuoteCommit { get; set; }
        [NotMapped]
        public string QuoteDetailRead { get; set; }
        [NotMapped]
        public string QuoteAudit { get; set; }
        [NotMapped]
        public string QuoteAudit2 { get; set; }
        [NotMapped]
        public string PurchaceCommit { get; set; }
        [NotMapped]
        public string PurchaceAudit { get; set; }
        [NotMapped]
        public string PurchaceAudit2 { get; set; }
        [NotMapped]
        public string PurchaceAudit3 { get; set; }
        [NotMapped]
        public string ChargeBackCommit { get; set; }
        [NotMapped]
        public string ChargeBackAudit { get; set; }
        [NotMapped]
        public string DepotAdmin { get; set; }
        [NotMapped]
        public string ReportExport { get; set; }

        public virtual Role Role { get; set; }

        [NotMapped]
        private string _roleName;
        [NotMapped]
        public string RoleName
        {
            get
            {
                return _roleName ?? Role?.Name;
            }
            set
            {
                _roleName = value;
            }
        }

    }
}
