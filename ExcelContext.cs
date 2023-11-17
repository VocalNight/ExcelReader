using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;

namespace ExcelReader
{
    public class ExcelContext : DbContext
    {
        public DbSet<DynamicTable> DynamicTables { get; set; }

        protected override void OnConfiguring( DbContextOptionsBuilder optionsBuilder )
        {
            SqlConnectionStringBuilder builder = new();

            builder.DataSource = "(localdb)\\mssqllocaldb";
            builder.InitialCatalog = "ExcelReader";
            builder.IntegratedSecurity = true;
            builder.TrustServerCertificate = true;
            builder.MultipleActiveResultSets = true;
            builder.ConnectTimeout = 3;

            string? connection = builder.ConnectionString;

            optionsBuilder.UseSqlServer(connection);
        }
    }
}
