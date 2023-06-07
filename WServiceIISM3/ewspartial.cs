using Microsoft.Exchange.WebServices.Autodiscover;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Web.Administration;
using System;
using System.Collections.Generic;
using System.Linq;
//using System.Security.Policy;
using System.Text.RegularExpressions;

namespace WServiceIISM3
{
    partial class ServiceIISM
    {
        // Методы для работы с EWS
        ExchangeService CreateService(string userName, string password, string autodiscoverAddress)
        {
            // подключение к Exchange сервису
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            const string autodiscoverUrl =
              @"https://ews.irkutskenergo.ru/EWS/Exchange.asmx";
            try
            {
                if (string.IsNullOrEmpty(userName) | string.IsNullOrEmpty(password))
                {
                    service.UseDefaultCredentials = true;
                }
                else
                {
                    service.Credentials = new WebCredentials(userName, password);
                }
                service.AutodiscoverUrl(autodiscoverAddress, RedirectionUrlValidationCallback);
            }

            catch (AutodiscoverRemoteException ex)
            {
                logger.Info("Exception is thrown: " + ex.Error.Message);
                service.Url = new Uri(autodiscoverUrl);
            }
            catch (AutodiscoverLocalException ex)
            {
                logger.Info("Exception is thrown: " + ex.Message);
                service.Url = new Uri(autodiscoverUrl);
            }
            //logger.Info(service.Url);
            return service;
        }

        /*static bool RedirectionUrlValidationCallback(String redirectionUrl)
        {
            // допускаем все переадресации
            return true;
        }*/
        bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;
            Uri redirectionUri = new Uri(redirectionUrl);
            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }

        // Главня логика обработки приходящих писем на почтовый ящик
        void EntireInbox(ExchangeService service)
        {
            // две записи равноценны (в первом случае создается объект делегата)
            //MethodSendMail methodS = new MethodSendMail(SendM);
            //MethodSendMail methodS1 = SendM;

            // Создадим объект где создано событие, подпишемся на него и допишем обработчик отправляющий в logger
            SendMail2 sendMail2 = new SendMail2(service);
            sendMail2.Notify += message => logger.Info(message);


            //Regex regexR = new Regex(@"restart", RegexOptions.Compiled | RegexOptions.IgnoreCase);
            var view = new ItemView(50);
            try
            {
                FindItemsResults<Item> findResults;
                do
                {
                    findResults = service.FindItems(WellKnownFolderName.Inbox, view);
                    var current = 0;
                    if (findResults.Items.Count() == 0)
                    {
                        // Лишнего нам в логах не надо
                        //logger.Info($"Писем в почтовом ящике {ewsUserName}: 0");
                    }
                    else
                    {
                        logger.Info($"Писем в почтовом ящике {ewsUserName}: {findResults?.TotalCount}");
                    }

                    // Отсортируем резултаты по возрастанию
                    var findResultsSorted = findResults.Items.OrderBy(a => a.DateTimeReceived);


                    foreach (var item in findResultsSorted)
                    {
                        string tmpstr0 = "", tmpstr1 = "", tmpstr2 = "";
                        var message = item as EmailMessage;
                        if (message == null)
                            continue;

                        // загрузка свойств письма
                        message.Load(new PropertySet(
                          BasePropertySet.FirstClassProperties));
                        // информационная строка с описанием письма
                        /*Console.WriteLine($"\r\nEmail:{++current}\r\nSubject:{message.Subject}\r\nFrom:{message.From}\r\n" +
                            $"To:{String.Join(String.Format("{0}t", Environment.NewLine), message.ToRecipients.Select(address => address.ToString()))}\r\n" +
                            $"Send:{message.DateTimeSent.ToShortDateString()} {message.DateTimeSent.ToShortTimeString()}\r\n" +
                            $"Get:{message.DateTimeReceived.ToShortDateString()} {message.DateTimeReceived.ToShortTimeString()}");*/
                        logger.Info($"Email:{++current}; Subject:{message.Subject}\r\nFrom:{message.From}\r\n" +
                            $"To:{String.Join(String.Format("{0}t", Environment.NewLine), message.ToRecipients.Select(address => address.ToString()))}\r\n" +
                            $"Send:{message.DateTimeSent.ToShortDateString()} {message.DateTimeSent.ToShortTimeString()} - " +
                            $"Get:{message.DateTimeReceived.ToShortDateString()} {message.DateTimeReceived.ToShortTimeString()}");


                        DateTime dtNow = DateTime.Now;

                        // переменная совпадения почтового ящика с разрешенным для управления
                        // Var 1
                        bool permittedEmail = false;
                        foreach (string li in listNE2)
                        {
                            if (Regex.IsMatch(message.From.ToString(), li, RegexOptions.IgnoreCase))
                            {
                                permittedEmail = true;
                                break;
                            }
                        }

                        // Variant 2 use LINQ (Создание своих переменых это преимущество перед методами расширений)
                        /*var permittedEmail2 = from p in listNE2
                                               let test = Regex.IsMatch(message.From.ToString(), p, RegexOptions.IgnoreCase)
                                               where test == true
                                               select p;
                        foreach(string p in permittedEmail2)
                            Console.WriteLine(p);*/

                        // Необходимо определить Site или Pool у нас в Subject и правильное действие есть в Subject
                        if (((Regex.IsMatch(message.Subject, "restart", RegexOptions.IgnoreCase)) |
                                (Regex.IsMatch(message.Subject, "recycle", RegexOptions.IgnoreCase)) |
                                (Regex.IsMatch(message.Subject, "start", RegexOptions.IgnoreCase)) |
                                (Regex.IsMatch(message.Subject, "stop", RegexOptions.IgnoreCase))) &
                                (message.DateTimeSent.AddMinutes(configiniD.Min) >= dtNow) & permittedEmail)
                        {

                            //разделим Subject на части и обработаем их
                            var PartMess = message.Subject.Split(':');

                            if (PartMess.Count() > 2 & (PartMess[0].Trim().ToUpper() == iisVar.POOL.ToString() | PartMess[0].Trim().ToUpper() == iisVar.SITE.ToString()))
                            {
                                tmpstr0 = PartMess[0].Trim();
                                tmpstr1 = PartMess[1].Trim();
                                tmpstr2 = PartMess[2].Trim();
                                if (tmpstr0.ToUpper() == iisVar.POOL.ToString())
                                {
                                    iisDvar = iisDetect(tmpstr1, iisVar.POOL.ToString());
                                }
                                else if (tmpstr0.ToUpper() == iisVar.SITE.ToString())
                                {
                                    iisDvar = iisDetect(tmpstr1, iisVar.SITE.ToString());
                                }

                            }
                            else
                            {
                                tmpstr1 = PartMess[0].Trim();
                                tmpstr2 = PartMess[1].Trim();
                                iisDvar = iisDetect(tmpstr1);
                            }
                        }
                        else
                        {
                            //Console.WriteLine($"В письме не найдено совпадение:(.. {message.Subject}");
                            // удаление письма
                            logger.Info($"Удалим письмо где в Subject нет триггира или по времени письмо устарело {message.Subject}\r\n от From:{message.From}\r\n" +
                                $"Send:{message.DateTimeSent.ToShortDateString()} {message.DateTimeSent.ToShortTimeString()}");
                            item.Delete(DeleteMode.HardDelete, true);
                            continue;
                        }

                        // Работа с POOL
                        if (iisDvar.ContainsKey(iisVar.POOL))
                        {
                            Data dat = new Data();
                            dat.SiteName = iisDvar[iisVar.POOL];
                            dat.To = message.From;
                            dat.Result = false;
                            if (string.IsNullOrEmpty(ServerName))
                                dat.From = configiniD.ServerName + configiniD.DomainPostName;
                            else
                                dat.From = ServerName + configiniD.DomainPostName;

                            if (iisDpool[iisPool.RECYCLE] == tmpstr2.ToLower())
                            {
                                logger.Info($"{iisDvar[iisVar.POOL]}:{iisDpool[iisPool.RECYCLE]} - {tmpstr2.ToLower()}");
                                if (iisPoolRSS(iisDvar, iisPool.RECYCLE.ToString().ToLower()))
                                {
                                    logger.Info($"Пул {iisDvar[iisVar.POOL]} успешно перегружен, удалим письмо");
                                    item.Delete(DeleteMode.HardDelete, true);

                                    dat.Doing = iisPool.RECYCLE.ToString();
                                    dat.PoolorSite = iisVar.POOL.ToString();
                                    dat.Result = true;


                                    // Вызов метода через создание объекта
                                    sendMail2.SendM(dat, ServerName);

                                    // 3 варианта вызова метода (2 первых через делегат) MethDel + MethDel2 одинаковы - разная запись
                                    // Отключим..
                                    //methodS(dat, service);     

                                    /*Action<Data, ExchangeService> MethDel = (x, y) => { SendM(x, y); };
                                    MethDel(dat, service);

                                    void MethDel2(Data x, ExchangeService y) { SendM(x, y); };
                                    MethDel2(dat, service);

                                    MethodSendMail method = (x, y) => SendM(x, y);
                                    method(dat, service);

                                    void method2(Data x, ExchangeService y) => SendM(x, y);
                                    method2(dat, service);*/
                                    //SendM(dat, service);

                                    System.Threading.Thread.Sleep(1000);
                                }
                                else
                                {
                                    logger.Info($"Пул {iisDvar[iisVar.POOL]} перегрузка не удалась, удалим письмо");
                                    item.Delete(DeleteMode.HardDelete, true);

                                    dat.Doing = iisPool.RECYCLE.ToString();
                                    dat.PoolorSite = iisVar.POOL.ToString();
                                    dat.Result = false;

                                    sendMail2.SendM(dat, ServerName);
                                    //methodS(dat, service);
                                    //SendM(dat, service);
                                    System.Threading.Thread.Sleep(1000);
                                }
                                continue;
                            }
                            else if (iisDpool[iisPool.START] == tmpstr2.ToLower())
                            {
                                logger.Info($"{iisDvar[iisVar.POOL]}:{iisDpool[iisPool.START]} - {tmpstr2.ToLower()}");
                                if (iisPoolRSS(iisDvar, iisPool.START.ToString().ToLower()))
                                {
                                    logger.Info($"Пул {iisDvar[iisVar.POOL]} успешно запущен, удалим письмо");
                                    item.Delete(DeleteMode.HardDelete, true);

                                    dat.Doing = iisPool.START.ToString();
                                    dat.PoolorSite = iisVar.POOL.ToString();
                                    dat.Result = true;

                                    sendMail2.SendM(dat, ServerName);
                                    //methodS(dat, service);
                                    //SendM(dat, service);
                                    System.Threading.Thread.Sleep(1000);
                                }
                                else
                                {
                                    logger.Info($"Пул {iisDvar[iisVar.POOL]} запуск не произведен, удалим письмо");
                                    item.Delete(DeleteMode.HardDelete, true);

                                    dat.Doing = iisPool.START.ToString();
                                    dat.PoolorSite = iisVar.POOL.ToString();
                                    dat.Result = false;

                                    sendMail2.SendM(dat, ServerName);
                                    //methodS(dat, service);
                                    //SendM(dat, service);
                                    System.Threading.Thread.Sleep(1000);
                                }
                                continue;
                            }
                            else if (iisDpool[iisPool.STOP] == tmpstr2.ToLower())
                            {
                                logger.Info($"{iisDvar[iisVar.POOL]}:{iisDpool[iisPool.STOP]} - {tmpstr2.ToLower()}");
                                if (iisPoolRSS(iisDvar, iisPool.STOP.ToString().ToLower()))
                                {
                                    logger.Info($"Пул {iisDvar[iisVar.POOL]} успешно выключен, удалим письмо");
                                    item.Delete(DeleteMode.HardDelete, true);

                                    dat.Doing = iisPool.STOP.ToString();
                                    dat.PoolorSite = iisVar.POOL.ToString();
                                    dat.Result = true;

                                    sendMail2.SendM(dat, ServerName);
                                    //methodS(dat, service);
                                    //SendM(dat, service);
                                    System.Threading.Thread.Sleep(1000);
                                }
                                else
                                {
                                    logger.Info($"Пул {iisDvar[iisVar.POOL]} выключить не удалось, удалим письмо");
                                    item.Delete(DeleteMode.HardDelete, true);

                                    dat.Doing = iisPool.STOP.ToString();
                                    dat.PoolorSite = iisVar.POOL.ToString();
                                    dat.Result = false;

                                    sendMail2.SendM(dat, ServerName);
                                    //methodS(dat, service);
                                    //SendM(dat, service);
                                    System.Threading.Thread.Sleep(1000);
                                }
                                continue;
                            }
                            else
                            {
                                //Ошибка с действием -> удаление письма
                                logger.Info($"Удалим письмо которое в Subject нет триггира или он неправилен {message.Subject}\r\n от From:{message.From}\r\n" +
                                    $"Send:{message.DateTimeSent.ToShortDateString()} {message.DateTimeSent.ToShortTimeString()}");
                                item.Delete(DeleteMode.HardDelete, true);
                                dat.Doing = tmpstr2;
                                dat.PoolorSite = iisVar.POOL.ToString();
                                dat.Result = false;

                                sendMail2.SendM(dat, ServerName);
                                //methodS(dat, service);
                                //SendM(dat, service);
                                System.Threading.Thread.Sleep(1000);
                                continue;
                            }

                        }

                        // Работа с Site
                        if (iisDvar.ContainsKey(iisVar.SITE))
                        {
                            Data dat = new Data();
                            dat.SiteName = iisDvar[iisVar.SITE];
                            dat.To = message.From;
                            dat.Result = false;
                            if (string.IsNullOrEmpty(ServerName))
                                dat.From = configiniD.ServerName + configiniD.DomainPostName;
                            else
                                dat.From = ServerName + configiniD.DomainPostName;


                            if (iisDsite[iisSite.RESTART] == tmpstr2.ToLower())
                            {
                                logger.Info($"{iisDvar[iisVar.SITE]}:{iisDsite[iisSite.RESTART]} - {tmpstr2.ToLower()}");
                                if (iisSiteRSS(iisDvar, iisSite.RESTART.ToString().ToLower()))
                                {
                                    logger.Info($"Сайт {iisDvar[iisVar.SITE]} успешно перегружен, удалим письмо");
                                    item.Delete(DeleteMode.HardDelete, true);

                                    dat.Doing = iisSite.RESTART.ToString();
                                    dat.PoolorSite = iisVar.SITE.ToString();
                                    dat.Result = true;

                                    sendMail2.SendM(dat, ServerName);
                                    //methodS(dat, service);
                                    //SendM(dat, service);
                                    System.Threading.Thread.Sleep(1000);
                                }
                                else
                                {
                                    logger.Info($"Сайт {iisDvar[iisVar.SITE]} перегрузка не удалась, удалим письмо");
                                    item.Delete(DeleteMode.HardDelete, true);

                                    dat.Doing = iisSite.RESTART.ToString();
                                    dat.PoolorSite = iisVar.SITE.ToString();
                                    dat.Result = false;

                                    sendMail2.SendM(dat, ServerName);
                                    //methodS(dat, service);
                                    //SendM(dat, service);
                                    System.Threading.Thread.Sleep(1000);
                                }
                                continue;
                            }
                            else if (iisDsite[iisSite.START] == tmpstr2.ToLower())
                            {
                                logger.Info($"{iisDvar[iisVar.SITE]}:{iisDsite[iisSite.START]} - {tmpstr2.ToLower()}");
                                if (iisSiteRSS(iisDvar, iisSite.START.ToString().ToLower()))
                                {
                                    logger.Info($"Сайт {iisDvar[iisVar.SITE]} успешно запущен, удалим письмо");
                                    item.Delete(DeleteMode.HardDelete, true);

                                    dat.Doing = iisSite.START.ToString();
                                    dat.PoolorSite = iisVar.SITE.ToString();
                                    dat.Result = true;

                                    sendMail2.SendM(dat, ServerName);
                                    //methodS(dat, service);
                                    //SendM(dat, service);
                                    System.Threading.Thread.Sleep(1000);
                                }
                                else
                                {
                                    logger.Info($"Сайт {iisDvar[iisVar.SITE]} запуск не произведен, удалим письмо");
                                    item.Delete(DeleteMode.HardDelete, true);

                                    dat.Doing = iisSite.START.ToString();
                                    dat.PoolorSite = iisVar.SITE.ToString();
                                    dat.Result = false;

                                    sendMail2.SendM(dat, ServerName);
                                    //methodS(dat, service);
                                    //SendM(dat, service);
                                    System.Threading.Thread.Sleep(1000);
                                }
                                continue;
                            }
                            else if (iisDsite[iisSite.STOP] == tmpstr2.ToLower())
                            {

                                logger.Info($"{iisDvar[iisVar.SITE]}:{iisDsite[iisSite.STOP]} - {tmpstr2.ToLower()}");
                                if (iisSiteRSS(iisDvar, iisSite.STOP.ToString().ToLower()))
                                {
                                    logger.Info($"Сайт {iisDvar[iisVar.SITE]} успешно выключен, удалим письмо");
                                    item.Delete(DeleteMode.HardDelete, true);

                                    dat.Doing = iisSite.STOP.ToString();
                                    dat.PoolorSite = iisVar.SITE.ToString();
                                    dat.Result = true;

                                    sendMail2.SendM(dat, ServerName);
                                    //methodS(dat, service);
                                    //SendM(dat, service);
                                    System.Threading.Thread.Sleep(1000);
                                }
                                else
                                {
                                    logger.Info($"Сайт {iisDvar[iisVar.SITE]} выключить не удалось, удалим письмо");
                                    item.Delete(DeleteMode.HardDelete, true);

                                    dat.Doing = iisSite.STOP.ToString();
                                    dat.PoolorSite = iisVar.SITE.ToString();
                                    dat.Result = false;

                                    sendMail2.SendM(dat, ServerName);
                                    //methodS(dat, service);
                                    //SendM(dat, service);
                                    System.Threading.Thread.Sleep(1000);
                                }
                                continue;
                            }
                            else
                            {
                                //Ошибка с дейсвием удаление письма
                                logger.Info($"Удалим письмо которое в Subject нет триггира или он неправилен {message.Subject}\r\n от From:{message.From}\r\n" +
                                    $"Send:{message.DateTimeSent.ToShortDateString()} {message.DateTimeSent.ToShortTimeString()}");
                                item.Delete(DeleteMode.HardDelete, true);
                                dat.Doing = tmpstr2;
                                dat.PoolorSite = iisVar.SITE.ToString();
                                dat.Result = false;

                                sendMail2.SendM(dat, ServerName);
                                //methodS(dat, service);
                                //SendM(dat, service);
                                System.Threading.Thread.Sleep(1000);
                                continue;
                            }


                        }

                        // Работа с NULL (Очистка писем где ненайден триггер)
                        if (iisDvar.ContainsKey(iisVar.NULL))
                        {
                            if (iisDvar[iisVar.NULL] == null)
                            {
                                // удаление письма
                                logger.Info($"Удалим письмо которое в Subject нет триггира {message.Subject}\r\n от From:{message.From}\r\n" +
                                    $"Send:{message.DateTimeSent.ToShortDateString()} {message.DateTimeSent.ToShortTimeString()}");
                                item.Delete(DeleteMode.HardDelete, true);
                            }
                            continue;
                        }


                    }
                    view.Offset += 50;
                } while (findResults.MoreAvailable);
            }
            catch (ServiceResponseException ex)
            {
                logger.Info("Exception is thrown (ExchangeService): " + ex.Message);
            }
        } // Главня логика обработки приходящих писем на почтовый ящик

        // Необходимо рапознать что перед нами Пул или Сайт №1
        Dictionary<iisVar, string> iisDetect(string str)
        {
            Dictionary<iisVar, string> iisDvar_m = new Dictionary<iisVar, string>();

            using (ServerManager manager = new ServerManager())
            {
                ApplicationPool appPool = manager.ApplicationPools.FirstOrDefault(ap => ap.Name == str);
                Site site = manager.Sites.FirstOrDefault(ap => ap.Name == str);
                if (!string.IsNullOrEmpty(appPool?.Name))
                {
                    iisDvar_m.Add(iisVar.POOL, appPool.Name);
                    logger.Info($"Имя пула:{appPool.Name}");
                }
                else if (!string.IsNullOrEmpty(site?.Name))
                {
                    iisDvar_m.Add(iisVar.SITE, site.Name);
                    logger.Info($"Имя сайта:{site.Name}");
                }
                else
                {
                    iisDvar_m.Add(iisVar.NULL, null);
                }
            }
            return iisDvar_m;
        }

        // Необходимо рапознать что перед нами Пул или Сайт №2
        Dictionary<iisVar, string> iisDetect(string str, string iPS)
        {
            Dictionary<iisVar, string> iisDvar_m = new Dictionary<iisVar, string>();
            ApplicationPool appPool = null;
            Site site = null;

            using (ServerManager manager = new ServerManager())
            {

                if (iPS == "POOL")
                {
                    appPool = manager.ApplicationPools.FirstOrDefault(ap => ap.Name == str);

                }
                else if (iPS == "SITE")
                {
                    site = manager.Sites.FirstOrDefault(ap => ap.Name == str);
                }

                if (!string.IsNullOrEmpty(appPool?.Name))
                {
                    iisDvar_m.Add(iisVar.POOL, appPool.Name);
                    logger.Info($"Имя пула:{appPool.Name}");

                }
                else if (!string.IsNullOrEmpty(site?.Name))
                {
                    iisDvar_m.Add(iisVar.SITE, site.Name);
                    logger.Info($"Имя сайта:{site.Name}");

                }
                else
                {
                    if(iPS == "POOL" | iPS == "SITE")
                        iisDvar_m.Add(iisVar.NULL, iPS);
                    else
                        iisDvar_m.Add(iisVar.NULL, null);
                }
            }
            return iisDvar_m;
        }

        // POOL: recycle, остановка, запуск
        bool iisPoolRSS(Dictionary<iisVar, string> iiP, string Doing)
        {
            bool iisPoolRSSbool = false;
            try
            {
                using (ServerManager manager = new ServerManager())
                {
                    ApplicationPool appPool = manager.ApplicationPools.FirstOrDefault(ap => ap.Name == iiP[iisVar.POOL]);
                    if (appPool != null)
                    {
                        //Get the current state of the app pool
                        bool appPoolRunning = appPool.State == ObjectState.Started || appPool.State == ObjectState.Starting;
                        bool appPoolStopped = appPool.State == ObjectState.Stopped || appPool.State == ObjectState.Stopping;

                        //The app pool is running -> go to Recycle
                        if (appPoolRunning & Doing == "recycle")
                        {
                            appPool.Recycle();
                            System.Threading.Thread.Sleep(3000);
                            iisPoolRSSbool = true;

                        }
                        else if (appPoolRunning & Doing == "stop")
                        {
                            appPool.Stop();
                            System.Threading.Thread.Sleep(3000);
                            appPoolStopped = true;
                            iisPoolRSSbool = true;
                        }
                        else if (appPoolStopped & Doing == "start")
                        {
                            appPool.Start();
                            System.Threading.Thread.Sleep(3000);
                            appPoolRunning = true;
                            iisPoolRSSbool = true;
                        }
                    }
                    else
                    {
                        throw new Exception(string.Format($"An Application Pool does not exist with the name {iiP[iisVar.POOL]}"));
                    }
                }

            }
            catch (Exception ex)
            {
                throw new Exception(string.Format($"Unable to restart the application pools for {iiP[iisVar.POOL]}"), ex.InnerException);

            }
            return iisPoolRSSbool;
        }

        // SITE: перегрузка, остановка, запуск
        bool iisSiteRSS(Dictionary<iisVar, string> iiP, string Doing)
        {
            bool iisSiteRSSbool = false;
            try
            {
                using (ServerManager manager = new ServerManager())
                {
                    Site site = manager.Sites.FirstOrDefault(ap => ap.Name == iiP[iisVar.SITE]);

                    if (site != null)
                    {
                        //Get the current state of the app pool
                        bool appSiteRunning = site.State == ObjectState.Started || site.State == ObjectState.Starting;
                        bool appSiteStopped = site.State == ObjectState.Stopped || site.State == ObjectState.Stopping;

                        //The app pool is running -> go to Recycle
                        if (appSiteRunning & Doing == "restart")
                        {
                            site.Stop();
                            System.Threading.Thread.Sleep(5000);
                            site.Start();
                            System.Threading.Thread.Sleep(5000);
                            if (appSiteRunning = site.State == ObjectState.Started || site.State == ObjectState.Starting)
                            {
                                iisSiteRSSbool = true;
                            }

                        }
                        else if (appSiteRunning & Doing == "stop")
                        {
                            site.Stop();
                            System.Threading.Thread.Sleep(5000);
                            if (appSiteStopped = site.State == ObjectState.Stopped || site.State == ObjectState.Stopping)
                            {
                                iisSiteRSSbool = true;
                            }
                        }
                        else if (appSiteStopped & Doing == "start")
                        {
                            site.Start();
                            System.Threading.Thread.Sleep(5000);
                            if (appSiteRunning = site.State == ObjectState.Started || site.State == ObjectState.Starting)
                            {
                                iisSiteRSSbool = true;
                            }
                        }
                    }
                    else
                    {
                        throw new Exception(string.Format($"An Application Pool does not exist with the name {iiP[iisVar.SITE]}"));
                    }
                }

            }
            catch (Exception ex)
            {
                throw new Exception(string.Format($"Unable to restart the application pools for {iiP[iisVar.SITE]}"), ex.InnerException);

            }
            return iisSiteRSSbool;
        }




    }
}
