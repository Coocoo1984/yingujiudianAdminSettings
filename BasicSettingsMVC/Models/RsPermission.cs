using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;

namespace BasicSettingsMVC.Models
{
    public partial class RsPermission
    {
        public long Id { get; set; }
        public string UsrWechatId { get; set; }
        public long PermissionId { get; set; }

        public virtual Permission Permission { get; set; }
        //public virtual Usr Usr { get; set; }

        [NotMapped]
        private string _permissionName;
        [NotMapped]
        public string PermissionName
        {
            get
            {
                return _permissionName ?? Permission?.Name;
            }
            set
            {
                _permissionName = value;
            }
        }
    }
}
