using Config.Net;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Win32;
using NLog;
using NLog.Config;
using NLog.Targets;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.ServiceProcess;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace WServiceIISM3
{
    public partial class ServiceIISM : ServiceBase
    {
        private bool _canseled;
        internal static string appPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
        internal static string appDirectory = System.IO.Path.GetDirectoryName(appPath);
        //internal readonly static string path = appDirectory + @"\testLog.txt";

        internal static Logger logger = LogManager.GetCurrentClassLogger();
        //static string CurrentWorkPath = AppDomain.CurrentDomain.BaseDirectory;
        static string LogDir = appDirectory + @"\logs";
        internal static IMySettings configiniD = new ConfigurationBuilder<IMySettings>()
            .UseIniFile(appDirectory + @"\config.ini", true)
            .Build();

        #region Access to pass
        // пароль для DB
        static string plainDBpass;
        internal static string GetPlainDBpass() => plainDBpass;
        internal static void SetPlainDBpass(string value) => plainDBpass = value;
        // пароль для EWS
        static string plainMailpass;
        internal static string GetPlainMailpass() => plainMailpass;
        internal static void SetPlainMailpass(string value) => plainMailpass = value;
        #endregion

        static string ServerName { get; set; }
        static string IpNew { get; set; }
        static readonly string ewsUserName = @"restartIIS@go.ru";
        static readonly string shortEwsUserName = @"restartIIS";
        static readonly string mainsite = @"\s(go.ru)";
        static readonly string ewssite = @"\s(ews.go.ru)";
        static readonly string autodiscoversite = @"\s(autodiscover.go.ru)";
        static List<string> listH = new List<string>();
        static List<string> listNE2 = new List<string>();
        
        //отключим создание делегата, ,будем работать с событием
        //public delegate void MethodSendMail(Data dat, ExchangeService service);

        struct DataFieldBD
        {
            public string Field1 { get; set; }
            public string Field2 { get; set; }
            public string Field3 { get; set; }

            public DataFieldBD(string field1, string field2, string field3)
            {
                this.Field1 = field1;
                this.Field2 = field2;
                this.Field3 = field3;
            }
        }

        // AES параметры
        #region AES начальные данные
        public static string passPhrase = "TolpaBilaORnebila";        //Может быть любой строкой
        public static string saltValue = "KrikunNaPloshadiMolchal";        // Может быть любой строкой
        public static string hashAlgorithm = "SHA256";             // может быть "MD5"
        public static int passwordIterations = 2;                //Может быть любым числом
        public static string initVector = "!1A3g2D5s9K996g7"; // Должно быть 16 байт
        public static int keySize = 256;                // Может быть 192 или 128
        #endregion

        // Определение возможных действий на IIS
        #region Определение возможных действий на IIS
        enum iisPool { START, STOP, RECYCLE }
        static Dictionary<iisPool, string> iisDpool = new Dictionary<iisPool, string> {
            { iisPool.STOP, "stop" },
            { iisPool.START, "start" },
            { iisPool.RECYCLE, "recycle" }
        };
        enum iisSite { START, STOP, RESTART }
        static Dictionary<iisSite, string> iisDsite = new Dictionary<iisSite, string> {
            { iisSite.STOP, "stop" },
            { iisSite.START, "start" },
            { iisSite.RESTART, "restart" }
        };
        public enum iisVar { SITE, POOL, NULL }
        static Dictionary<iisVar, string> iisDvar;

        #endregion Определение возможных действий на IIS

        public ServiceIISM()
        {
            InitializeComponent();
            // Если в config.ini пустое поле ServerName то получем его из Env
            if (string.IsNullOrEmpty(configiniD.ServerName))
            {
                ServerName = Environment.MachineName;
            }

            // Получим IP адреса на сервере (формат IPv4)
            IpNew = IPhostGet();

        }

        protected override void OnStart(string[] args)
        {
            _canseled = false;
            System.Threading.Tasks.Task task = System.Threading.Tasks.Task.Run(() => { Processing(); });
        }


        protected override void OnStop()
        {
            _canseled = true;
            logger.Info("Пришла Команда на Выход!!!");
            EventLog.WriteEntry("Службе дан приказ на выход", EventLogEntryType.Information);
        }

        void Processing()
        {
            #region NLog Initializator

            var config = new NLog.Config.LoggingConfiguration();
            LogManager.Configuration = new LoggingConfiguration();
            const string LayoutFile = @"[${date:format=yyyy-MM-dd HH\:mm\:ss}] >> ${message} ${exception: format=ToString}";

            var logfile = new FileTarget();


            if (!Directory.Exists(LogDir))
                Directory.CreateDirectory(LogDir);

            // Rules for mapping loggers to targets
            config.AddRule(LogLevel.Trace, LogLevel.Fatal, logfile);

            DateTime dtlog = DateTime.Now;

            logfile.CreateDirs = true;
            logfile.FileName = $"{LogDir}{Path.DirectorySeparatorChar}{dtlog:yyyy-MM-dd}.log";
            logfile.AutoFlush = true;
            logfile.LineEnding = LineEndingMode.CRLF;
            logfile.Layout = LayoutFile;
            logfile.FileNameKind = FilePathKind.Absolute;
            logfile.ConcurrentWrites = false;
            logfile.KeepFileOpen = false;
            // по умолчанию стоит utf-8
            logfile.Encoding = Encoding.Default;

            // Apply config
            NLog.LogManager.Configuration = config;

            #endregion NLog Initializator

            logger.Info($"Запущена служба: WSrviceIIS3");
            EventLog.WriteEntry("Служба стартанула", EventLogEntryType.Information);

            // генерация ms от 10сек до 60сек
            int msR = RND();

            // Получим пароль из config.ini
            if (!string.IsNullOrEmpty(configiniD.Password))
            {
                SetPlainMailpass(RijndaelAlgorithm.Decrypt(
                    configiniD.Password,
                    passPhrase,
                    saltValue,
                    hashAlgorithm,
                    passwordIterations,
                    initVector,
                    keySize
                    ));
            }
            else
            {
                logger.Info("В Config.ini пустое поле пароля :(, так не должно быть: Exit");
                System.Environment.Exit(0);
            }

            // Выведем в лог есть ли записи которые нам необходимы
            List<string> listHosts = HostList();
            if(listHosts.Count > 0)
            {
                logger.Info("Необходимые нам записи в файле Hosts");
                listHosts.ForEach(delegate(string str) { logger.Info(str); });
            } else
            {
                logger.Info("Наших записей адресов в Hosts нет..");
            }

            try
            {
                while (!_canseled)
                {
                    #region Создание нового файла для лога при смене даты дня
                    DateTime Dlog = DateTime.Now;
                    bool FileNotDate = false;
                    var Logfiles = Directory.GetFiles(LogDir, "*.log");
                    foreach (var fi in Logfiles)
                    {
                        FileInfo file = new FileInfo(fi);
                        string tmp_str = file.Name.Substring(0, file.Name.Length - 4);
                        string tmp_data_string = $"{Dlog:yyyy-MM-dd}";
                        //logger.Info($"{tmp_str} - {tmp_data_string}");

                        if (tmp_str == tmp_data_string)
                        {
                            //logger.Info("TRUE");
                            FileNotDate = false;
                            break;
                        }
                        else
                        {
                            //logger.Info("false");
                            FileNotDate = true;
                        }

                    }
                    if (FileNotDate)
                    {
                        using (FileStream fs = File.Create($"{LogDir}{Path.DirectorySeparatorChar}{Dlog:yyyy-MM-dd}.log"))
                        {
                            byte[] info = new UTF8Encoding(true).GetBytes($"New Log File Create..{Dlog:yyyy-MM-dd HH-mm}" + Environment.NewLine);
                            // Add some information to the file.
                            fs.Write(info, 0, info.Length);
                        }
                        logfile.FileName = $"{LogDir}{Path.DirectorySeparatorChar}{Dlog:yyyy-MM-dd}.log";
                    }
                    #endregion Создание нового файла для лога при смене даты дня

                    #region TEST section
                    /*logger.Info("Служу, не тужу в системе сижу.");

                    using (DataModelContext DC = new DataModelContext())
                    {
                        //DC.Database.Log = x => File.AppendAllText(path, x, Encoding.Default);
                        DC.Database.Log = x => logger.Info(x);
                        logger.Info("");
                        logger.Info(DC.Database.Connection.ConnectionString);

                        var Zapros = DC.Servers.ToList();
                        if (Zapros.Count == 0)
                        {
                            logger.Info($"..{Zapros.ToArray()}");
                            Server server = new Server { Name = ServerName, Ipaddr = IpNew, Field1 = configiniD.Field1, Field2 = configiniD.Field2, Field3 = configiniD.Field3 };
                            DC.Servers.Add(server);
                            DC.SaveChanges();

                        }
                        else
                        {
                            foreach (var s in Zapros)
                            {
                                logger.Info($"{s.Name}.{s.Ipaddr} : {s.Field1},\r\n{s.Field2},\r\n{s.Field3}\r\n");
                            }

                        }

                    }*/
                    #endregion TEST

                    // получим список записей из файла hosts
                    listH = HostList();
                    // Заполним структуру
                    DataFieldBD dStruck = new DataFieldBD(configiniD.Field1, configiniD.Field2, configiniD.Field3);

                    // запишем наши данные в базу и файл Hosts если там нет нужных записей
                    try
                    {
                         DBWrite(listH, IpNew, dStruck);
                         listNE2 = NEmailsM();
                    }
                    catch (Exception ex)
                    {
                        logger.Info("Working with DB + hosts: ");
                        logger.Error(ex.Message);
                    }

                    // Тут будет продолжена работа с севисом EWS
                    // создать сервис
                    
                    ExchangeService service = CreateService(shortEwsUserName, GetPlainMailpass(), ewsUserName);

                    //logger.Info($"*** Testing email for: {ewsUserName} ***");

                    EntireInbox(service);

                    // Обработка логов
                    FLogs();








                    Thread.Sleep(60000 + msR);
                }


            }
            catch (Exception ex)
            {
                logger.Info("Global Error:(.. to exit");
                logger.Info($"{ex.Message}");
                System.Environment.Exit(0);
            }
        }


        // метод случайных числел для изменения времени запуска
        public int RND()
        {
            Random rnd = new Random();
            int value = rnd.Next(10000, 30000);
            return value;
        }

        // Прочитаем Таблицу разрешенных Email адресов
        List<string> NEmailsM()
        {
            List<string> listNE2 = new List<string>();
            using (DataModelContext context = new DataModelContext())
            {
                var NEZapros = context.NEmails.ToList();
                if (NEZapros?.Count > 0)
                {
                    foreach (var item in NEZapros)
                    {
                        listNE2.Add(item.EmailName);
                    }
                }
            }
            return listNE2;
        } // Прочитаем Таблицу разрешенных Email адресов



        // Чтение файла hosts из system32
        List<string> HostList()
        {
            List<string> list = new List<string>();
            //HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters
            RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Services\Tcpip\Parameters", false);
            if (string.IsNullOrEmpty(key?.GetValue("DataBasePath")?.ToString()))
            {
                logger.Info($"Такого ключа {key} в реестре нет");
                return null;
            }

            string line = string.Empty;
            using (StreamReader stream = new StreamReader(key?.GetValue("DataBasePath")?.ToString() + "\\hosts", Encoding.Default))
            {
                try
                {
                    while ((line = stream.ReadLine()) != null)
                    {
                        if ((Regex.IsMatch(line, mainsite, RegexOptions.IgnoreCase)) |
                            (Regex.IsMatch(line, ewssite, RegexOptions.IgnoreCase)) |
                            (Regex.IsMatch(line, autodiscoversite, RegexOptions.IgnoreCase)))
                        {
                            char[] separators = new char[] { ' ', '\t' };
                            string[] line_tmp = line.Trim().Split(separators, StringSplitOptions.RemoveEmptyEntries);

                            if (line_tmp.Count() == 2)
                            {
                                list.Add(line_tmp[1] + "|" + line_tmp[0]);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    logger.Info("Ошибка в методе HostList():");
                    logger.Info(ex.Message);
                }
            }
            
            // Уберем лишний вывод из логов
            /*foreach (var item in list)
            {
                logger.Info($"{item}");
            }*/
            return list;
        }

        // Запись в БД метод
        void DBWrite(List<string> list, string ipNew, DataFieldBD dStruck)
        {
            // запишем наши данные в базу
            if (list?.Count == 3)
            {
                using (DataModelContext context = new DataModelContext())
                {
                    var NServer = context.Servers.FirstOrDefault(u => u.Name == ServerName);

                    if (NServer == null)
                    {
                        Server ser_ = new Server { Name = ServerName, Ipaddr = ipNew, Field1 = configiniD.Field1, Field2 = configiniD.Field2, Field3 = configiniD.Field3 };

                        context.Servers.Add(ser_);
                        context.SaveChanges();
                    }
                    else
                    {
                        if (string.IsNullOrEmpty(NServer.Field1) | !NServer.Field1.Equals(list[0]))
                        {
                            NServer.Field1 = configiniD.Field1;
                        }
                        if (string.IsNullOrEmpty(NServer.Field2) | !NServer.Field2.Equals(list[1]))
                        {
                            NServer.Field2 = configiniD.Field2;
                        }
                        if (string.IsNullOrEmpty(configiniD.Field3) | !NServer.Field3.Equals(list[2]))
                        {
                            NServer.Field3 = configiniD.Field3;
                        }
                        if (string.IsNullOrEmpty(NServer.Ipaddr) | !NServer.Ipaddr.Equals(ipNew))
                        {
                            NServer.Ipaddr = ipNew;
                        }
                        context.SaveChanges();
                    }
                    HostLogicWrite(list, dStruck);
                }
            }
            else
            {
                using (DataModelContext contextExistBD = new DataModelContext())
                {
                    var NServ = contextExistBD.Servers.FirstOrDefault(u => u.Name == ServerName);

                    if (NServ == null)
                    {
                        Server ser_ = new Server() { Name = ServerName, Ipaddr = ipNew, Field1 = configiniD.Field1, Field2 = configiniD.Field2, Field3 = configiniD.Field3 };
                        contextExistBD.Servers.Add(ser_);
                        contextExistBD.SaveChanges();
                    }
                    else if (NServ != null)
                    {
                        if (string.IsNullOrEmpty(configiniD.Field1) | !configiniD.Field1.Equals(NServ.Field1))
                        {
                            NServ.Field1 = configiniD.Field1;
                        }
                        if (string.IsNullOrEmpty(configiniD.Field2) | !configiniD.Field2.Equals(NServ.Field2))
                        {
                            NServ.Field2 = configiniD.Field2;
                        }
                        if (string.IsNullOrEmpty(configiniD.Field3) | !configiniD.Field3.Equals(NServ.Field3))
                        {
                            NServ.Field3 = configiniD.Field3;
                        }
                        if (string.IsNullOrEmpty(NServ.Ipaddr) | !NServ.Ipaddr.Equals(ipNew))
                        {
                            NServ.Ipaddr = ipNew;
                        }
                        contextExistBD.SaveChanges();
                    }
                    HostLogicWrite(list, dStruck);
                }
            }
        } // Запись в БД метод

        // Запись Логики данных в hosts из Config.INI
        void HostLogicWrite(List<string> list, DataFieldBD dStruck)
        {
            if (list?.Count > 0)
            {
                if ((list.Contains(dStruck.Field1)) & (list.Contains(dStruck.Field2)) & (list.Contains(dStruck.Field3)))
                {
                    // Уберем лишний вывод в лог
                    //logger.Info("Записи в host совпали с записями в config.ini");
                }
                else
                {
                    logger.Info("Запишем в host, записи из config.ini");
                    if (!HWrite(list, dStruck))
                    {
                        logger.Info($"Стирание полей из файла Hosts: не удалась, проверте доступы к файлу hosts");
                        System.Environment.Exit(0);
                    }
                    if (!HWrite(dStruck))
                    {
                        logger.Info($"Запись полей {dStruck.Field1},{Environment.NewLine}{dStruck.Field2},{Environment.NewLine}{dStruck.Field3}: не удалась, проверте доступы к файлу hosts");
                        System.Environment.Exit(0);
                    }
                }
            }
            else
            {
                logger.Info("В файле hosts нет записей, давай их туда запишем");
                if (!HWrite(dStruck))
                {
                    logger.Info($"Запись полей в hosts:{Environment.NewLine}{dStruck.Field1},{Environment.NewLine}{dStruck.Field2},{Environment.NewLine}{dStruck.Field3}: не удалась, exit");
                    System.Environment.Exit(0);
                }
            }
        } // Запись Логики данных в hosts из Config.INI

        // Запись непосредственно в файл HOSTS (Значения разные в файле Hosts будем их исправлять)
        bool HWrite(List<string> list, DataFieldBD dStruck)
        {
            bool resultB = false;
            RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Services\Tcpip\Parameters", false);
            if (string.IsNullOrEmpty(key?.GetValue("DataBasePath")?.ToString()))
            {
                logger.Info($"Такого ключа {key} в реестре нет");
                return false;
            }

            // Значения разные в файле Hosts будем их справлять
            string strF = string.Empty;
            using (StreamReader stream = new StreamReader(key?.GetValue("DataBasePath")?.ToString() + "\\hosts", Encoding.Default))
            {
                try
                {
                    strF = stream.ReadToEnd();
                    resultB = true;
                }
                catch (Exception ex)
                {
                    logger.Info(ex.Message);
                }

            }
            char[] separators = new char[] { '|' };
            List<string> list_tmp = new List<string>();
            foreach (var str in list)
            {
                string[] line_tmp = str.Trim().Split(separators, StringSplitOptions.RemoveEmptyEntries);
                list_tmp.Add($@"^.*?\s{line_tmp[0]}");
            }

            foreach (var item in list_tmp)
            {
                strF = Regex.Replace(strF, item, "", RegexOptions.Multiline);
            }

            using (StreamWriter streamW = new StreamWriter(key?.GetValue("DataBasePath")?.ToString() + "\\hosts", false))
            {
                try
                {
                    streamW.Write(strF);
                    resultB = true;
                }
                catch (Exception ex)
                {
                    logger.Info(ex.Message);
                    resultB = false;
                }
            }

            string line_tmp2 = string.Empty;
            string line = string.Empty;
            using (StreamReader stream = new StreamReader(key?.GetValue("DataBasePath")?.ToString() + "\\hosts", Encoding.Default))
            {
                try
                {
                    while ((line = stream.ReadLine()) != null)
                    {
                        if (!string.IsNullOrEmpty(line))
                        {
                            line_tmp2 += line + Environment.NewLine;
                        }
                    }
                }
                catch (Exception ex)
                {
                    logger.Info(ex.Message);
                    resultB = false;
                }
            }

            using (StreamWriter streamW = new StreamWriter(key?.GetValue("DataBasePath")?.ToString() + "\\hosts", false))
            {
                try
                {
                    streamW.Write(line_tmp2);
                    resultB = true;
                }
                catch (Exception ex)
                {
                    logger.Info(ex.Message);
                    resultB = false;
                }
            }
            return resultB;
        }

        // Тупо допишим записи в файл Hosts
        bool HWrite(DataFieldBD dStruck)
        {
            bool resultB = false;
            RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Services\Tcpip\Parameters", false);
            if (string.IsNullOrEmpty(key?.GetValue("DataBasePath")?.ToString()))
            {
                logger.Info($"Такого ключа {key} в реестре нет");
                return false;
            }

            if (!string.IsNullOrEmpty(dStruck.Field1) & !string.IsNullOrEmpty(dStruck.Field2) & !string.IsNullOrEmpty(dStruck.Field3))
            {
                try
                {
                    string strFEnd = string.Empty;
                    bool PerevodCar = false;

                    using (StreamReader streamReader = new StreamReader(key?.GetValue("DataBasePath")?.ToString() + "\\hosts", Encoding.Default))
                    {
                        strFEnd = streamReader.ReadToEnd();
                        if (!strFEnd.EndsWith(Environment.NewLine))
                        {
                            PerevodCar = true;
                        }
                    }


                    using (StreamWriter streamW = new StreamWriter(key?.GetValue("DataBasePath")?.ToString() + "\\hosts", true))
                    {
                        char[] separators = new char[] { '|' };
                        string[] line_field_tmp1 = dStruck.Field1.Trim().Split(separators, StringSplitOptions.RemoveEmptyEntries);
                        string[] line_field_tmp2 = dStruck.Field2.Trim().Split(separators, StringSplitOptions.RemoveEmptyEntries);
                        string[] line_field_tmp3 = dStruck.Field3.Trim().Split(separators, StringSplitOptions.RemoveEmptyEntries);

                        if (PerevodCar)
                        {
                            streamW.WriteLine();
                        }
                        streamW.WriteLine($"{line_field_tmp1[1]} {line_field_tmp1[0]}");
                        streamW.WriteLine($"{line_field_tmp2[1]} {line_field_tmp2[0]}");
                        streamW.WriteLine($"{line_field_tmp3[1]} {line_field_tmp3[0]}");
                    }
                    resultB = true;
                }
                catch (Exception ex)
                {
                    logger.Info("Ошибка записи в файл Host, Метод HWrite(DataFieldBD dStruck)");
                    logger.Info(ex.Message);

                }
            }
            else
            {
                logger.Info("Заполните поля field[1,2,3] в файле config.ini");
            }
            return resultB;
        } // Тупо допишим записи в файл Hosts

        // Определим количество разрешенных лог файлов и возьмем его в config.ini
        void FLogs()
        {
            string[] filePaths = Directory.GetFiles(LogDir + @"\", "*.log");
            string[] filePaths2 = null;
            if (filePaths.Count() > configiniD.QuantityLogs)
            {
                //DateTime dtDate;
                CultureInfo provider = CultureInfo.CreateSpecificCulture("en-US");
                DateTimeStyles styles = DateTimeStyles.None;
                int count = 0;
                foreach (var logFile in filePaths)
                {
                    if (DateTime.TryParse(logFile.Remove(logFile.Length - 4).Remove(0, 5), provider, styles, out DateTime dtDate))
                    {
                        if (DateTime.Today.AddDays(-configiniD.QuantityLogs) > dtDate)
                        {
                            if (filePaths2?.Count() == configiniD.QuantityLogs + 1) { break; }
                            try
                            {
                                File.Delete(logFile);
                                count++;
                            }
                            catch (Exception ex)
                            {
                                logger.Info(ex.Message);
                            }
                        }
                    }
                    filePaths2 = Directory.GetFiles(LogDir + @"\", "*.log");
                }
                logger.Info($"Всего было лог файлов {filePaths.Length} из них удалено старых {count}");

                logger.Info($"Дата устаревания {DateTime.Today.AddDays(-configiniD.QuantityLogs)} и разрешенное кол-во логов {configiniD.QuantityLogs}");
            }
        }





    }

    // интерфейс для работы с полями в config.ini
    public interface IMySettings
    {
        // пароль к почтовому ящику
        string Password { get; }
        string ServerName { get; set; }
        string DomainPostName { get; }
        // количество мин когда письмо устаревает и должно быть удалено
        int Min { get; }
        // количество лог файлов в директории
        int QuantityLogs { get; }
        // Сервер MS SQL
        string NameMSSQL { get; }
        // Имя Базы Данных
        string NameDataBase { get; }
        // Имя пользователя для работы с Базой Данных
        string LoginDB { get; }
        // Пароль пользователя к БД
        string PasswordUserDB { get; }
        string Address { get; }
        string Field1 { get; }
        string Field2 { get; }
        string Field3 { get; }
        bool LocalDB { get; }

    }
}
