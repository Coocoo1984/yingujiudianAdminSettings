using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;

namespace BasicSettingsMVC.Models
{
    public partial class Goods
    {
        public long Id { get; set; }
        public string Code { get; set; }
        public string Name { get; set; }
        public string Desc { get; set; }
        public long? GoodsUnitId { get; set; }
        public long? GoodsClassId { get; set; }
        public string Specification { get; set; }
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
        public string GoodsClassName
        {
            get
            {
                return GoodsClass?.Name;
            }
        }
        [NotMapped]
        public string GoodsUnitName
        {
            get
            {
                return GoodsUnit?.Name;
            }
        }
        [NotMapped]
        public string BizTypeName
        {
            get
            {
                return GoodsClass.BizType.Name;
            }
        }

        public virtual GoodsClass GoodsClass { get; set; }
        public virtual GoodsUnit GoodsUnit { get; set; }
    }
}
