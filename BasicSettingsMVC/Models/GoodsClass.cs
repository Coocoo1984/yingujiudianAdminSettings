using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;

namespace BasicSettingsMVC.Models
{
    public partial class GoodsClass
    {
        public long Id { get; set; }
        public string Code { get; set; }
        public string Name { get; set; }
        public string Specification { get; set; }
        public string Desc { get; set; }
        public string Remark { get; set; }
        public long? BizTypeId { get; set; }
        public bool Disable { get; set; }

        [NotMapped]
        public string DisableForShow
        {
            get
            {
                return Disable ? "否" : "是";
            }
            set
            {
                DisableForShow = value;
                Disable = value == "否" ? true : false;
            }
        }
        [NotMapped]
        public string BizTypeName
        {
            get
            {
                return BizType?.Name;
            }
        }
        public virtual BizType BizType { get; set; }
        public virtual ICollection<Goods> Goods { get; set; }
    }
}
