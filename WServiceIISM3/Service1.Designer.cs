using NLog;
using System;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;

namespace WServiceIISM3
{
    partial class ServiceIISM
    {
        /// <summary> 
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary> 
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            // 
            // ServiceIISM
            // 
            this.ServiceName = "ServiceIISM";

        }

        #endregion

        // Определим какие IPv4 адреса сущестуют на сервере (v6 опустим)
        private string IPhostGet()
        {
            IPHostEntry hostIP2 = Dns.GetHostEntry(ServerName);
            IpNew = string.Empty;
            if (hostIP2 != null)
            {
                foreach (var ip in hostIP2.AddressList)
                {
                    // только в формате ipv4
                    if (Regex.IsMatch(ip.ToString(), @"\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}", RegexOptions.IgnoreCase))
                        IpNew += $"{ip} & ";
                }
                IpNew = IpNew.Trim(' ', '&');
                //Console.WriteLine($"{ipNew}");
                Array.ForEach(IpNew.Split('&').Select(s => s.Trim()).ToArray(), logger.Info);
            }
            else
            {
                logger.Info("НЕ могу найти ServerName, пропиши его тогда руками в Config.ini -> Exit");
                System.Environment.Exit(0);
            }
            return IpNew;
        }


    }
}
