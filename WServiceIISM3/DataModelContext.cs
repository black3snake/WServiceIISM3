using System;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Data.Entity.ModelConfiguration;
using System.Net.NetworkInformation;

namespace WServiceIISM3
{
    // Класс конфигурации
    public class MyDbConfig : DbConfiguration
    {
        public MyDbConfig()
        {
            if (!string.IsNullOrEmpty(ServiceIISM.configiniD.PasswordUserDB) & !ServiceIISM.configiniD.LocalDB)
            {
                try
                {
                    Ping png = new Ping();
                    PingReply pingReply = png.Send(ServiceIISM.configiniD.NameMSSQL, 2000);
                    if (pingReply.Status == IPStatus.Success)
                    {
                        // Получим пароль из config.ini
                        ServiceIISM.SetPlainDBpass(RijndaelAlgorithm.Decrypt(
                            ServiceIISM.configiniD.PasswordUserDB,
                            ServiceIISM.passPhrase,
                            ServiceIISM.saltValue,
                            ServiceIISM.hashAlgorithm,
                            ServiceIISM.passwordIterations,
                            ServiceIISM.initVector,
                            ServiceIISM.keySize
                            ));

                        SqlConnectionFactory defaultFactory = new SqlConnectionFactory($@"Data Source={ServiceIISM.configiniD.NameMSSQL};User ID={ServiceIISM.configiniD.LoginDB};Password={ServiceIISM.GetPlainDBpass()};Integrated security=false;TrustServerCertificate=True;");
                        this.SetDefaultConnectionFactory(defaultFactory);
                    }
                    else
                    {
                        ServiceIISM.logger.Info($"Сервер SQL:{ServiceIISM.configiniD.NameMSSQL} не пингуется");
                        System.Environment.Exit(0);
                    }
                }
                catch (Exception ex)
                {
                    ServiceIISM.logger.Info(ex.Message);
                }
            }
            else if (ServiceIISM.configiniD.LocalDB)
            {
                ServiceIISM.logger.Info($"Поле LocalDB:{ServiceIISM.configiniD.LocalDB} пробуем использовать локальную БД");

                SqlConnectionFactory defaultFactory = new SqlConnectionFactory($@"Server=(localdb)\mssqllocaldb;Database={ServiceIISM.configiniD.NameDataBase};Trusted_Connection=True;");
                //optionsBuilder.UseSqlServer(@"Server=(localdb)\mssqllocaldb;Database=IISManadgerBase;Trusted_Connection=True;");

            }
            else
            {
                ServiceIISM.logger.Info($"Проверте наличие файла Config.ini и содержимое {ServiceIISM.configiniD.PasswordUserDB} поля с паролем к БД");
            }

            //SqlConnectionFactory defaultFactory =
             //new SqlConnectionFactory($@"Data Source={ServiceIISM.configiniD.NameMSSQL};User ID={ServiceIISM.configiniD.LoginDB};Password={ServiceIISM.configiniD.PasswordUserDB};Integrated security=false;");
                //new SqlConnectionFactory(@"Data Source=TSD-SQL02-ID;Initial Catalog=RUserBaseTest;TrustServerCertificate=True;Integrated Security=True;");
                //new SqlConnectionFactory(@"Server=TSD-SQL02-ID;User=sh;Password=y;TrustServerCertificate=True;");
                //new SqlConnectionFactory("Server=TSD-SQL02-ID;User=a_p;Password=4;TrustServerCertificate=True;");

            //this.SetDefaultConnectionFactory(defaultFactory);
            /*this.SetProviderServices("System.Data.SqlClient",
                System.Data.Entity.SqlServer.SqlProviderServices.Instance);*/

        }
    }
    [DbConfigurationType(typeof(MyDbConfig))]
    //[DbConfigurationType("WServiceIISM3.MyDbConfig, MyAssembly")]
    internal class DataModelContext : DbContext
    {
        //public DataModelContext() : base("name=DBConn")
        public DataModelContext() : base($"{ServiceIISM.configiniD.NameDataBase}")
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
