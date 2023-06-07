using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Data.Entity.ModelConfiguration;

namespace WServiceIISM3.DataContext
{
    // Класс конфигурации
    public class MyDbConfig : DbConfiguration
    {
        public MyDbConfig()
        {
            SqlConnectionFactory defaultFactory =
                new SqlConnectionFactory("Server=TSQL02-ID;User=s_p;Password=password;TrustServerCertificate=True;");
                //new SqlConnectionFactory("Server=TSQL02-ID;User=s_p;Password=password;TrustServerCertificate=True;");

            this.SetDefaultConnectionFactory(defaultFactory);
            /*this.SetProviderServices("System.Data.SqlClient",
                System.Data.Entity.SqlServer.SqlProviderServices.Instance);*/

        }
    }
    [DbConfigurationType(typeof(MyDbConfig))]
    //[DbConfigurationType("WServiceIISM3.MyDbConfig, MyAssembly")]
    internal class DataModelContext : DbContext
    {
        //public DataModelContext() : base("name=DBConn")
        public DataModelContext() : base("RUserBaseTest")
        {
            //Database.SetInitializer(new CreateDatabaseIfNotExists<DataContext>());
            Database.SetInitializer(new DataContextInitializer());
        }
        public DbSet<Server> Servers { get; set; }
        public DbSet<NEmail> NEmails { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Configurations.Add(new ServerConfigurations());
            modelBuilder.Configurations.Add(new NEmailConfigurations());
        }

        private class ServerConfigurations : EntityTypeConfiguration<Server>
        {
            public ServerConfigurations()
            {
                this.HasKey(e => e.Name);
                this.Property(e => e.Name).HasMaxLength(100).IsRequired();
                this.Property(e => e.Ipaddr);
                this.Property(e => e.Field1).HasMaxLength(100);
                this.Property(e => e.Field2).HasMaxLength(100);
                this.Property(e => e.Field3).HasMaxLength(100);
            }
        }
        private class NEmailConfigurations : EntityTypeConfiguration<NEmail>
        {
            public NEmailConfigurations()
            {
                this.HasKey(e => e.Id);
                this.Property(e => e.EmailName).HasMaxLength(150).IsRequired();
            }
        }

        public class DataContextInitializer : CreateDatabaseIfNotExists<DataModelContext>
        {
            protected override void Seed(DataModelContext context)
            {
                /*IList<User> users = new List<User>();
                users.Add(new User() { Id = 1, Name = "Tom", Age = 23 });
                users.Add(new User() { Id = 2, Name = "Alisa", Age = 32 });

                foreach (var user in users)
                    context.Users.Add(user);*/
                base.Seed(context);
                //context.SaveChanges();
            }
        }
    }
}
