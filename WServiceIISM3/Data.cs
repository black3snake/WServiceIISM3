using Microsoft.Exchange.WebServices.Data;

namespace WServiceIISM3
{
    public class Data
    {
        public string PoolName { get; set; }
        public string SiteName { get; set; }
        public string Doing { get; set; }
        public string PoolorSite { get; set; }
        public bool Result { get; set; }
        public string From { get; set; }
        public EmailAddress To { get; set; }
        public EmailAddress Cc { get; set; }
    }

}
