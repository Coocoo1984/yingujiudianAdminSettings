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
        private string _disableForShow;
        [NotMapped]
        public string DisableForShow
        {
            get
            {
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
        private string _goodsClassName;
        [NotMapped]
        public string GoodsClassName
        {
            get
            {
                return _goodsClassName ?? GoodsClass?.Name;
            }
            set
            {
                _goodsClassName = value;
            }
        }
        [NotMapped]
        private string _goodsUnitName;
        [NotMapped]
        public string GoodsUnitName
        {
            get
            {
                return _goodsUnitName ?? GoodsUnit?.Name;
            }
            set
            {
                _goodsUnitName = value;
            }
        }
        [NotMapped]
        private string _bizTypeName;
        [NotMapped]
        public string BizTypeName
        {
            get
            {
                return _bizTypeName?? GoodsClass?.BizType?.Name;
            }
            set
            {
                _bizTypeName = value;
            }
        }

        public virtual GoodsClass GoodsClass { get; set; }
        public virtual GoodsUnit GoodsUnit { get; set; }
    }
}
