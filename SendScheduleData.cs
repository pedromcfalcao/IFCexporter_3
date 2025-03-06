using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace GetJarvisInfoTest
{
    class SendScheduleData
    {
        public const String JARVIS_SERVER_URL = "http://192.168.1.137/";
        public const String JARVIS_SCHEDULE = "/intranet/job_schedules/edit.php";
        
        public String SendData(String jobNo, String defaultFactory, String scheduleType, String scheduleData, ref String htmlResponse)
        {
            String _scheduleId = "";

            // Create a request using a URL that can receive a post.
            WebRequest request = WebRequest.Create(JARVIS_SERVER_URL + JARVIS_SCHEDULE);
            // Set the Method property of the request to POST.
            request.Method = "POST";
            // Create POST data and convert it to a byte array.
            string postData = "callMode=WS&btnAccion=Save&n_obra=" + jobNo + "&fabrica=" + defaultFactory + "&schedule_type=" + scheduleType + "&pasted_data=" + scheduleData;
            byte[] byteArray = Encoding.UTF8.GetBytes(postData);
            // Set the ContentType property of the WebRequest.
            request.ContentType = "application/x-www-form-urlencoded";
            // Set the ContentLength property of the WebRequest.
            request.ContentLength = byteArray.Length;

            request.Credentials = new System.Net.NetworkCredential("smlibrary", "Htdaab?SML33");

            // Get the request stream.
            Stream dataStream = request.GetRequestStream();
            // Write the data to the request stream.
            dataStream.Write(byteArray, 0, byteArray.Length);
            // Close the Stream object.
            dataStream.Close();
            // Get the response.
            WebResponse response = request.GetResponse();
            // Display the status.
            Console.WriteLine(((HttpWebResponse)response).StatusDescription);
            // Get the stream containing content returned by the server.
            dataStream = response.GetResponseStream();
            // Open the stream using a StreamReader for easy access.
            StreamReader reader = new StreamReader(dataStream);
            // Read the content.
            string responseFromServer = reader.ReadToEnd();
            // Display the content.
            Console.WriteLine(responseFromServer);
            // Clean up the streams.
            reader.Close();
            dataStream.Close();
            response.Close();

            htmlResponse = responseFromServer;

            //Parse the Server Response
            var htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(responseFromServer);

            var htmlNodes = htmlDoc.DocumentNode.SelectNodes("//div[@id='ws_callmode_return_id']");

            if (htmlNodes!=null && htmlNodes.Count == 1)
            {
                _scheduleId = htmlNodes.ElementAt(0).InnerHtml;
            }

            return _scheduleId;
        }
    }
}
