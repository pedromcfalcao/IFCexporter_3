using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GetJarvisInfoTest
{
    public class JarvisItem
    {
        [JsonProperty("rlin")]
        public String Rlin { get; set; }
        [JsonProperty("name")]
        public String Name { get; set; }
        [JsonProperty("ref")]
        public String Reference { get; set; }
        [JsonProperty("qty_ord")]
        public String Qty_ord { get; set; }
        [JsonProperty("qty_man")]
        public String Qty_man { get; set; }
        [JsonProperty("qty_del")]
        public String Qty_del { get; set; }

    }
    public class JarvisOrder
    {
        [JsonProperty("order_no")]
        public String OrderNo { get; set; }
        [JsonProperty("items")]
        public List<JarvisItem> ItemList { get; set; }


    }
    public class JarvisJob
    {
        [JsonProperty("job_no")]
        public String JobNo { get; set; }
        [JsonProperty("orders")]
        public List<JarvisOrder> OrderList { get; set; }

    }

    public class JarvisInfo
    {
        [JsonProperty("status_code")]
        public String StatusCode { get; set; }
        [JsonProperty("status_msg")]
        public String StatusMsg { get; set; }
        [JsonProperty("items_count")]
        public int ItemsCount { get; set; }
        [JsonProperty("items_info")]
        public List<JarvisJob> JobList { get; set; }

    }

    public class GetJarvisInfo
    {
        private const String JARVIS_SERVER_URL  = "http://192.168.1.137/";
        private const String JARVIS_GET_INFO_BY_DATE = "/intranet/w_pedidos/ajax_get_items_info_by_date.php";
        private const String JARVIS_GET_INFO_BY_ORDERNO = "/intranet/w_pedidos/ajax_get_items_info_by_orderno.php";

        public JarvisInfo GetItemsInfoByDate(DateTime initialDate, DateTime endDate, String jobNo)
        {
            string responseFromServer;



            WebClient client = new WebClient();

            // Add a user agent header in case the 
            // requested URI contains a query.

            client.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
            client.Credentials = new System.Net.NetworkCredential("smlibrary", "Htdaab?SML33");

            Stream data = client.OpenRead(JARVIS_SERVER_URL + JARVIS_GET_INFO_BY_DATE + "?startdate=" + initialDate.ToString("dd-MM-yyyy") + "&enddate=" + endDate.ToString("dd-MM-yyyy") + "&jobno=" + jobNo);
            StreamReader reader = new StreamReader(data);
            responseFromServer = reader.ReadToEnd();
            data.Close();
            reader.Close();



            var jarvisInfo = JsonConvert.DeserializeObject<JarvisInfo>(responseFromServer);
            return jarvisInfo;
        }


        public static bool TryConnectJarvis()
        {
            //http://192.168.1.58/ Jarvis server
            //http://192.168.1.137/ Jarvis server test


            try
            {
                WebClient client = new WebClient();
                client.Credentials = new System.Net.NetworkCredential("smlibrary", "Htdaab?SML33");
                Stream data = client.OpenRead(JARVIS_SERVER_URL);
                StreamReader reader = new StreamReader(data);
                string responseFromServer = reader.ReadToEnd();
                data.Close();
                reader.Close();
               return  true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                MessageBox.Show("Error: " + ex.Message);
                return false;

            }

           // return;
        }




            public static JarvisInfo GetItemsInfoByOrder(String orderNo, String itemReference)
        {
            string responseFromServer;

            if (TryConnectJarvis() == true)
            {


                WebClient client = new WebClient();

                // Add a user agent header in case the 
                // requested URI contains a query.

                client.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
                client.Credentials = new System.Net.NetworkCredential("smlibrary", "Htdaab?SML33");

                Stream data = client.OpenRead(JARVIS_SERVER_URL + JARVIS_GET_INFO_BY_ORDERNO + "?orderno=" + orderNo + "&reference=" + itemReference);
                StreamReader reader = new StreamReader(data);
                responseFromServer = reader.ReadToEnd();
                data.Close();
                reader.Close();

                var jarvisInfo = JsonConvert.DeserializeObject<JarvisInfo>(responseFromServer);
                return jarvisInfo;
            }
            else
            {
                return null;
            }
        }
    }
}
