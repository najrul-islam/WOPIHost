using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Generic;
using Microsoft.EntityFrameworkCore;
using WopiHost.Data.Models;
using Microsoft.Extensions.Configuration;

namespace WopiHost.Data
{
    public class WopiHostContext : DbContext
    {
        private readonly IConfiguration Configuration;


        public DbSet<User> User { get; set; }
        public DbSet<Document> Document { get; set; }
        public WopiHostContext(DbContextOptions<WopiHostContext> options,
            IConfiguration configuration)
            : base(options)
        {
            Configuration = configuration;
        }
        //public string DbPath { get; private set; }

        /*public WopiHostContext()
        {
            var folder = Environment.SpecialFolder.LocalApplicationData;
            var path = Environment.GetFolderPath(folder);
            DbPath = $"{path}{System.IO.Path.DirectorySeparatorChar}wopihost.db";
        }*/

        // The following configures EF to create a Sqlite database file in the
        // special "local" folder for your platform.
        protected override void OnConfiguring(DbContextOptionsBuilder options)
        {
            if (!options.IsConfigured)
            {
                //UseSqlite
                /*string wopisqlite = $"DataSource={Configuration.GetSection("WebRootPath")}\\LiteDb\\wopisqlite.db";
                options.UseSqlite(wopisqlite);*/

                //UseSqlServer
                string sqlServer = $"{Configuration.GetSection("WopiHostContext").Value}";
                options.UseSqlServer(sqlServer);
            }
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<User>(entity =>
            {
                entity.Property(e => e.UserName).HasColumnType("NVARCHAR").HasMaxLength(250);
                entity.Property(e => e.UserFullName).HasColumnType("NVARCHAR").HasMaxLength(250);
                entity.Property(e => e.Email).HasColumnType("NVARCHAR").HasMaxLength(250);
                entity.Property(e => e.Password).HasColumnType("NVARCHAR").HasMaxLength(250);
            });

            modelBuilder.Entity<Document>(entity =>
            {
                entity.Property(e => e.DocumentName).HasColumnType("NVARCHAR").HasMaxLength(250);
                //entity.Property(e => e.FileName).HasColumnType("NVARCHAR").HasMaxLength(250);
                entity.Property(e => e.Blob).HasColumnType("NVARCHAR").HasMaxLength(250);
            });
        }
    }
}
