namespace topmeperp.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class PLAN_CERT_ORDER
    {
        [StringLength(30)]
        public string CERT_ORD_ID { get; set; }

        [Key]
        [StringLength(50)]
        public string INQUIRY_FORM_ID { get; set; }

        [StringLength(200)]
        public string SUPPLIER_ID { get; set; }

        [StringLength(30)]
        public string PROJECT_ID { get; set; }

        [StringLength(30)]
        public string CREATE_ID { get; set; }

        public DateTime? CREATE_DATE { get; set; }

        [StringLength(30)]
        public string MODIFY_ID { get; set; }

        public DateTime? MODIFY_DATE { get; set; }

        [StringLength(20)]
        public string STATUS { get; set; }
    }
}
