using System;
using System.Linq;
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
        public virtual DbSet<Permission> Permission { get; set; }
        public virtual DbSet<RsPermission> RsPermission { get; set; }
        public virtual DbSet<Usr> Usr { get; set; }
        public virtual DbSet<Role> Role { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.EnableSensitiveDataLogging();
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
                    .ValueGeneratedOnAdd();

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

                entity.HasMany(e => e.GoodsClasses).WithOne(e => e.BizType).HasForeignKey(e => e.BizTypeId);
            });

            modelBuilder.Entity<Goods>(entity =>
            {
                entity.ToTable("goods").HasKey(k => k.Id);

                entity.Property(e => e.Id)
                    .HasColumnName("id")
                    .ValueGeneratedOnAdd();

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
                entity.ToTable("goods_class").HasKey(k => k.Id);

                entity.Property(e => e.Id)
                    .HasColumnName("id")
                    .ValueGeneratedOnAdd();

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

                entity.HasMany(e => e.Goods).WithOne(e => e.GoodsClass).HasForeignKey(e => e.GoodsClassId);
            });

            modelBuilder.Entity<GoodsUnit>(entity =>
            {
                entity.ToTable("goods_unit").HasKey(k => k.Id);

                entity.Property(e => e.Id)
                    .HasColumnName("id")
                    .ValueGeneratedOnAdd();

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

            modelBuilder.Entity<Permission>(entity =>
            {
                entity.ToTable("permission").HasKey(k => k.Id);

                entity.Property(e => e.Id)
                    .HasColumnName("id")
                    .ValueGeneratedOnAdd();

                entity.Property(e => e.Code)
                    .HasColumnName("code")
                    .HasColumnType("text(24)")
                    .HasDefaultValueSql("''");

                entity.Property(e => e.Desc)
                    .HasColumnName("desc")
                    .HasColumnType("text(48)")
                    .HasDefaultValueSql("''");

                entity.Property(e => e.Name)
                    .IsRequired()
                    .HasColumnName("name")
                    .HasColumnType("text(48)")
                    .HasDefaultValueSql("''");

                entity.HasMany(e => e.RsPermission).WithOne(e => e.Permission).HasForeignKey(e => e.PermissionId);
            });

            modelBuilder.Entity<RsPermission>(entity =>
            {
                entity.ToTable("rs_permission").HasKey(k => k.Id);

                entity.Property(e => e.Id)
                    .HasColumnName("id")
                    .ValueGeneratedOnAdd();

                entity.Property(e => e.UsrWechatId)
                    .HasColumnName("usr_wechat_id")
                    .HasColumnType("text(64)")
                    .HasDefaultValueSql("''");

                entity.Property(e => e.Disable)
                    .HasColumnName("disable")
                    .HasColumnType("INT(1)")
                    .HasDefaultValueSql("0");

                entity.Property(e => e.PermissionId).HasColumnName("permission_id");
            });

            modelBuilder.Entity<Role>(entity =>
            {
                entity.ToTable("role").HasKey(k => k.Id);

                entity.Property(e => e.Id)
                    .HasColumnName("id")
                    .ValueGeneratedOnAdd();

                entity.Property(e => e.Code)
                    .HasColumnName("code")
                    .HasColumnType("text(24)")
                    .HasDefaultValueSql("''");

                entity.Property(e => e.Name)
                    .HasColumnName("name")
                    .HasColumnType("text(48)")
                    .HasDefaultValueSql("''");

                entity.HasMany(e => e.Usr).WithOne(e => e.Role).HasForeignKey(e => e.RoleID);
            });

            modelBuilder.Entity<Usr>(entity =>
            {
                entity.ToTable("usr").HasKey(k => k.ID);

                entity.Property(e => e.ID)
                    .HasColumnName("id")
                    .ValueGeneratedOnAdd();

                entity.Property(e => e.WechatID)
                    .HasColumnName("wechat_id")
                    .HasColumnType("text(64)")
                    .HasDefaultValueSql("''");

                entity.Property(e => e.Code)
                    .HasColumnName("code")
                    .HasColumnType("text(64)")
                    .HasDefaultValueSql("''");

                entity.Property(e => e.Name)
                    .IsRequired()
                    .HasColumnName("name")
                    .HasColumnType("text(64)")
                    .HasDefaultValueSql("''");

                entity.Property(e => e.Desc)
                    .HasColumnName("desc")
                    .HasColumnType("text(48)")
                    .HasDefaultValueSql("''");

                entity.Property(e => e.Tel)
                    .HasColumnName("tel")
                    .HasColumnType("text(16)")
                    .HasDefaultValueSql("''");

                entity.Property(e => e.Tel1)
                    .HasColumnName("tel1")
                    .HasColumnType("text(16)")
                    .HasDefaultValueSql("''");

                entity.Property(e => e.Mobile)
                    .HasColumnName("mobile")
                    .HasColumnType("text(11)")
                    .HasDefaultValueSql("''");

                entity.Property(e => e.Mobile1)
                    .HasColumnName("mobile1")
                    .HasColumnType("text(11)")
                    .HasDefaultValueSql("''");
                
                entity.Property(e => e.RoleID).HasColumnName("role_id");
            });
        }
    }
}
