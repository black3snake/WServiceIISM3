using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WServiceIISM3
{
    public partial class ServiceIISM
    {
        void SendM(Data data, ExchangeService service)
        {
            DateTime dt = DateTime.Now;
            string htmlH = @"<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01//EN' 'http://www.w3.org/TR/html4/strict.dtd'>
                            <html><head>
                            <meta http-equiv='Content-Type' content='text/html; charset=utf-8'>
                            <title> Ответ программы управления IIS </title>
                            <style>
                            .layer1 { font: normal 12pt/10pt serif;} 
                            .cap { font: bold italic 12pt serif; }
                             </style></head>";

            string htmlT1 = @"<table width='700' cellpadding='5' cellspacing='1' border='0'>
                            <tr>
	                        <td align='center' class='cap'>форма отчета программы управления IIS<td>
                            </tr>
                            </table>";

            string htmlT2 = $@"<table width='700' cellpadding='5' cellspacing='1' border='1'>
	                        <tr bgcolor='#81B764'>
                            <td colspan= '2' class='layer1' align='left'>Pool or Site on Server</td>
		                    <td align='center'>{dt:dd-MM-yyyy}</td>
                            </tr>
                            <tr>
		                    <td align='left'>{ServerName}:{data.PoolorSite}:{data.PoolName}{data.SiteName}:</td>
		                    <td align='center'>{data.Doing}</td>
		                    
                            <td align='center'>Результат:{data.Result}</td>
	                        </tr>
                            </tables>";
            string htmlF = @"</html>";

            EmailMessage message = new EmailMessage(service)
            {
                // Set properties on the email message.
                Subject = $"Отчет программы управления IIS {dt:dd.MM.yyyy}",
                Body = htmlH + htmlT1 + htmlT2 + htmlF
            };
            message.ToRecipients.Add(data.To.Address);
            // Send the email message and save a copy.
            // This method call results in a CreateItem call to EWS.
            //message.SendAndSaveCopy();
            message.Send();
            logger.Info("An email with the subject '" + message.Subject + "' has been sent to '" + message.ToRecipients[0] + "' and saved in the SendItems folder.");


        }


    }
}
