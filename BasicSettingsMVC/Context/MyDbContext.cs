using System;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata;

namespace BasicSettingsMVC.Models
{
    public partial class MyDbContext : DbContext
    {
        public MyDbContext()
        {
        }

        public MyDbContext(DbContextOptions<MyDbContext> options)
            : base(options)
        {
        }

        public virtual DbSet<BizType> BizType { get; set; }
        public virtual DbSet<Goods> Goods { get; set; }
        public virtual DbSet<GoodsClass> GoodsClass { get; set; }
        public virtual DbSet<GoodsUnit> GoodsUnit { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            if (!optionsBuilder.IsConfigured)
            {
                optionsBuilder.UseSqlite("Data Source=C:\\Users\\Juan\\Desktop\\test.db");
            }
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.HasAnnotation("ProductVersion", "2.2.4-servicing-10062");

            modelBuilder.Entity<BizType>(entity =>
            {
                entity.ToTable("biz_type");

                entity.Property(e => e.Id)
                    .HasColumnName("id")
                    .ValueGeneratedNever();

                entity.Property(e => e.Code)
                    .HasColumnName("code")
                    .HasColumnType("text(4)")
                    .HasDefaultValueSql("''");

                entity.Property(e => e.Desc)
                    .HasColumnName("desc")
                    .HasColumnType("text(48)")
                    .HasDefaultValueSql("''");

                entity.Property(e => e.Disable)
                    .HasColumnName("disable")
                    .HasColumnType("INT(1)")
                    .HasDefaultValueSql("0");

                entity.Property(e => e.Name)
                    .HasColumnName("name")
                    .HasColumnType("text(24)")
                    .HasDefaultValueSql("''");
                //
                entity.HasMany(e => e.GoodsClasses).WithOne(e => e.BizType).HasForeignKey(e => e.BizTypeId);
            });

            modelBuilder.Entity<Goods>(entity =>
            {
                entity.ToTable("goods");

                entity.Property(e => e.Id)
                    .HasColumnName("id")
                    .ValueGeneratedNever();

                entity.Property(e => e.Code)
                    .HasColumnName("code")
                    .HasColumnType("text(24)")
                    .HasDefaultValueSql("''");

                entity.Property(e => e.Desc)
                    .HasColumnName("desc")
                    .HasColumnType("text(48)")
                    .HasDefaultValueSql("''");

                entity.Property(e => e.Disable)
                    .HasColumnName("disable")
                    .HasColumnType("INT(1)")
                    .HasDefaultValueSql("0");

                entity.Property(e => e.GoodsClassId).HasColumnName("goods_class_id");

                entity.Property(e => e.GoodsUnitId).HasColumnName("goods_unit_id");

                entity.Property(e => e.Name)
                    .IsRequired()
                    .HasColumnName("name")
                    .HasColumnType("text(24)")
                    .HasDefaultValueSql("''");

                entity.Property(e => e.Specification)
                    .HasColumnName("specification")
                    .HasColumnType("text(48)")
                    .HasDefaultValueSql("''");
            });

            modelBuilder.Entity<GoodsClass>(entity =>
            {
                entity.ToTable("goods_class");

                entity.Property(e => e.Id)
                    .HasColumnName("id")
                    .ValueGeneratedNever();

                entity.Property(e => e.BizTypeId).HasColumnName("biz_type_id");

                entity.Property(e => e.Code)
                    .HasColumnName("code")
                    .HasColumnType("text(8)")
                    .HasDefaultValueSql("''");

                entity.Property(e => e.Desc)
                    .HasColumnName("desc")
                    .HasColumnType("text(48)")
                    .HasDefaultValueSql("''");

                entity.Property(e => e.Disable)
                    .HasColumnName("disable")
                    .HasColumnType("INT(1)")
                    .HasDefaultValueSql("0");

                entity.Property(e => e.Name)
                    .IsRequired()
                    .HasColumnName("name")
                    .HasColumnType("text(24)")
                    .HasDefaultValueSql("''");

                entity.Property(e => e.Remark)
                    .HasColumnName("remark")
                    .HasColumnType("text(96)")
                    .HasDefaultValueSql("''");

                entity.Property(e => e.Specification)
                    .HasColumnName("specification")
                    .HasColumnType("text(48)")
                    .HasDefaultValueSql("''");

                //
                entity.HasMany(e => e.Goods).WithOne(e => e.GoodsClass).HasForeignKey(e => e.GoodsClassId);
            });

            modelBuilder.Entity<GoodsUnit>(entity =>
            {
                entity.ToTable("goods_unit");

                entity.Property(e => e.Id)
                    .HasColumnName("id")
                    .ValueGeneratedNever();

                entity.Property(e => e.Code)
                    .HasColumnName("code")
                    .HasColumnType("text(8)")
                    .HasDefaultValueSql("''");

                entity.Property(e => e.Desc)
                    .HasColumnName("desc")
                    .HasColumnType("text(48)")
                    .HasDefaultValueSql("''");

                entity.Property(e => e.Name)
                    .IsRequired()
                    .HasColumnName("name")
                    .HasColumnType("text(24)")
                    .HasDefaultValueSql("''");

                entity.Property(e => e.Remark)
                    .HasColumnName("remark")
                    .HasColumnType("text(48)")
                    .HasDefaultValueSql("''");

                entity.HasMany(e => e.Goods).WithOne(e => e.GoodsUnit).HasForeignKey(e => e.GoodsUnitId);
            });
        }
    }
}
