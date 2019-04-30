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
        public virtual BizType BizType { get; set; }
        public virtual ICollection<Goods> Goods { get; set; }

        [NotMapped]
        private string _disableForShow;
        [NotMapped]
        public string DisableForShow
        {
            get
            {
                if(_disableForShow == null)
                    _disableForShow = Disable ? "否" : "是";
                return _disableForShow;
            }
            set
            {
                _disableForShow = value;
                Disable = value == "否" ? true : false;
            }
        }
        [NotMapped]
        private string _bizTypeName;
        [NotMapped]
        public string BizTypeName
        {
            get
            {
                return _bizTypeName?? BizType?.Name;
            }
            set
            {
                _bizTypeName = value;
            }
        }

    }
}
