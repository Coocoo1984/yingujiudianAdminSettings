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
        public virtual string  DisableForShow {
            get {
                return Disable?"否":"是" ;
            }
            set
            {
                DisableForShow =  value;
                Disable = value == "否" ? true : false;
            }
        }
        public virtual ICollection<GoodsClass> GoodsClasses { get; set; }
    }
}
