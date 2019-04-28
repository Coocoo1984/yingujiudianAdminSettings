using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;

namespace BasicSettingsMVC.Models
{
    public partial class BizType
    {
        public long Id { get; set; }
        public string Code { get; set; }
        public string Name { get; set; }
        public string Desc { get; set; }
        public bool Disable { get; set; }

        [NotMapped]
        private string _disableForShow;
        [NotMapped]
        public string  DisableForShow {
            get {
                _disableForShow = Disable ? "否" : "是";
                return _disableForShow;
            }
            set
            {
                _disableForShow = value;
                Disable = (_disableForShow == "否") ? true : false;
            }
        }
        public virtual ICollection<GoodsClass> GoodsClasses { get; set; }
    }
}
