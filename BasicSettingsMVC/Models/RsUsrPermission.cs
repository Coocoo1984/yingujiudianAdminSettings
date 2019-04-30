using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;

namespace BasicSettingsMVC.Models
{
    public partial class RsUsrPermission
    {
        public long Id { get; set; }
        public string UsrWechatId { get; set; }
        public long PermissionId { get; set; }

        public virtual Usr Usr { get; set; }
        public virtual ICollection<Permission> Permission { get; set; }

    }
}
