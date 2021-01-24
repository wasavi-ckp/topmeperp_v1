using System;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity;
using System.Linq;

namespace topmeperp.Models
{
    public partial class Plan_Cert_OrderContext : DbContext
    {
        public Plan_Cert_OrderContext()
            : base("name=Plan_Cert_Order")
        {
        }

        public virtual DbSet<PLAN_CERT_ORDER_ITEM> PLAN_CERT_ORDER_ITEM { get; set; }
        public virtual DbSet<PLAN_CERT_ORDER> PLAN_CERT_ORDER { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Entity<PLAN_CERT_ORDER_ITEM>()
                .Property(e => e.CERT_ORD_ID)
                .IsUnicode(false);

            modelBuilder.Entity<PLAN_CERT_ORDER_ITEM>()
                .Property(e => e.PLAN_ITEM_ID)
                .IsUnicode(false);

            modelBuilder.Entity<PLAN_CERT_ORDER_ITEM>()
                .Property(e => e.TYPE_CODE)
                .IsUnicode(false);

            modelBuilder.Entity<PLAN_CERT_ORDER_ITEM>()
                .Property(e => e.SUB_TYPE_CODE)
                .IsUnicode(false);

            modelBuilder.Entity<PLAN_CERT_ORDER_ITEM>()
                .Property(e => e.ITEM_ID)
                .IsUnicode(false);

            modelBuilder.Entity<PLAN_CERT_ORDER_ITEM>()
                .Property(e => e.ITEM_UNIT)
                .IsUnicode(false);

            modelBuilder.Entity<PLAN_CERT_ORDER_ITEM>()
                .Property(e => e.ORDER_QTY)
                .HasPrecision(18, 3);

            modelBuilder.Entity<PLAN_CERT_ORDER_ITEM>()
                .Property(e => e.ORDER_PRICE)
                .HasPrecision(18, 3);

            modelBuilder.Entity<PLAN_CERT_ORDER_ITEM>()
                .Property(e => e.ACCUMULTATE_QTY)
                .HasPrecision(18, 3);

            modelBuilder.Entity<PLAN_CERT_ORDER_ITEM>()
                .Property(e => e.MODIFY_ID)
                .IsUnicode(false);

            modelBuilder.Entity<PLAN_CERT_ORDER_ITEM>()
                .Property(e => e.WAGE_PRICE)
                .HasPrecision(18, 3);

            modelBuilder.Entity<PLAN_CERT_ORDER>()
                .Property(e => e.CERT_ORD_ID)
                .IsUnicode(false);

            modelBuilder.Entity<PLAN_CERT_ORDER>()
                .Property(e => e.INQUIRY_FORM_ID)
                .IsUnicode(false);

            modelBuilder.Entity<PLAN_CERT_ORDER>()
                .Property(e => e.SUPPLIER_ID)
                .IsUnicode(false);

            modelBuilder.Entity<PLAN_CERT_ORDER>()
                .Property(e => e.PROJECT_ID)
                .IsUnicode(false);

            modelBuilder.Entity<PLAN_CERT_ORDER>()
                .Property(e => e.CREATE_ID)
                .IsUnicode(false);

            modelBuilder.Entity<PLAN_CERT_ORDER>()
                .Property(e => e.MODIFY_ID)
                .IsUnicode(false);

            modelBuilder.Entity<PLAN_CERT_ORDER>()
                .Property(e => e.STATUS)
                .IsUnicode(false);
        }
    }
}
