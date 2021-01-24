namespace topmeperp.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class PLAN_CERT_ORDER_ITEM
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public long CERT_ORD_ITEM_ID { get; set; }

        [StringLength(30)]
        public string CERT_ORD_ID { get; set; }

        [StringLength(50)]
        public string PLAN_ITEM_ID { get; set; }

        [StringLength(6)]
        public string TYPE_CODE { get; set; }

        [StringLength(6)]
        public string SUB_TYPE_CODE { get; set; }

        [StringLength(50)]
        public string ITEM_ID { get; set; }

        [StringLength(1000)]
        public string ITEM_DESC { get; set; }

        [StringLength(10)]
        public string ITEM_UNIT { get; set; }

        public decimal? ORDER_QTY { get; set; }

        public decimal? ORDER_PRICE { get; set; }

        public decimal? ACCUMULTATE_QTY { get; set; }

        [StringLength(1000)]
        public string ITEM_REMARK { get; set; }

        [StringLength(30)]
        public string MODIFY_ID { get; set; }

        public DateTime? MODIFY_DATE { get; set; }

        public decimal? WAGE_PRICE { get; set; }
    }
}
