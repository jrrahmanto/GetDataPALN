using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Drawing;
using System.Net.Http;
using System.IO;
using Newtonsoft.Json;
using RestSharp;
using System.Net;
using System.Text;
using IronPdf;
using System.Text.RegularExpressions;
using IronOcr;
using OpenPop.Pop3;
using OpenPop.Mime;
using System.Security.Policy;

namespace WorkerEmail
{
    public class Worker : BackgroundService
    {
        //tutorial google drive
        //https://www.youtube.com/watch?v=pHOweM1Gl6c
        //create project
        //open api drive
        private readonly ILogger<Worker> _logger;
        public Worker(ILogger<Worker> logger)
        {
            _logger = logger;
        }
        public override Task StartAsync(CancellationToken cancellationToken)
        {
            return base.StartAsync(cancellationToken);
        }
        public override Task StopAsync(CancellationToken cancellationToken)
        {
            _logger.LogInformation("Service stopped");
            return base.StopAsync(cancellationToken);
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    var Ocr = new IronTesseract();
                    IronOcr.Installation.LicenseKey = "IRONOCR.PTKLIRINGBERJANGKAINDONESIA.IRO211213.9250.23127.312112-1EA71407D8-HWNZXEAJ3YDY3-N2BU5BL3WRB6-YYXFD7XQVLTB-YLR4RKFSW22L-F7OLHD7TWYX3-MUPLUQ-LVPA4EVNX2WIEA-PROFESSIONAL.SUB-DNMTTQ.RENEW.SUPPORT.13.DEC.2022";
                    var dr = new DataSet1TableAdapters.iDfsFasPalnTableAdapter();
                    string path = AppDomain.CurrentDomain.BaseDirectory + "\\report\\{0}";
                    Pop3Client client = new Pop3Client();
                    client.Connect("outlook.office365.com", 995, true);
                    client.Authenticate("automatic_ptkbi@outlook.com", "Jakarta2021");
                    var messageCount = client.GetMessageCount();
                    for (int j = 0; j < (messageCount); j++)
                    {
                        Message getMessage = client.GetMessage(j + 1);
                        var headers = getMessage.Headers;
                        Console.WriteLine(headers.Subject);
                        if (headers.Subject.ToString().ToLower().Contains("equity paln vaf"))
                        {
                            sendTelegram("-1001671146559", "Proccess insert PALN VALBURY EQUITY from email start " + DateTime.Now.ToString("HH:mm:ss"));

                            foreach (var attachment in getMessage.FindAllAttachments())
                            {
                                var caption = attachment.ContentType.Name;
                                //string ext = Path.GetExtension(attachment.ContentType.Name);

                                string path_file1 = string.Format(path, caption);

                                if (System.IO.File.Exists(path_file1))
                                {
                                    System.IO.File.Delete(path_file1);
                                }

                                FileStream Stream = new FileStream(path_file1, FileMode.Create);
                                BinaryWriter BinaryStream = new BinaryWriter(Stream);
                                BinaryStream.Write(attachment.Body);
                                BinaryStream.Close();
                                string hasil = "0";
                                using (var input = new OcrInput())
                                {
                                    input.AddImage(path_file1);
                                    var Result = Ocr.Read(input);
                                    var x = Result.Words;//43
                                    hasil = x[75].ToString().Replace("$", "");
                                    hasil = hasil.Replace(",", "");
                                    File.Delete(path_file1);
                                }
                                var dt_val = dr.GetDataByDate(DateTime.Now.Date, Convert.ToDecimal(14));
                                if (dt_val.Count == 0)
                                {
                                    dr.Insert(14, DateTime.Now.Date, Convert.ToDecimal(hasil));
                                    sendTelegram("-1001671146559", "Success insert PALN VALBURY EQUITY : $ " + hasil + "\nTimestamp " + DateTime.Now.ToString("HH:mm:ss"));
                                }
                                else
                                {
                                    sendTelegram("-1001671146559", "PALN VALBURY already , EQUITY : $ " + dt_val[0].PALN + "\nTimestamp " + DateTime.Now.ToString("HH:mm:ss"));
                                }
                            }
                        }
                        else if (headers.Subject.ToString().ToLower().Contains("daily statement pt. straits"))
                        {
                            sendTelegram("-1001671146559", "Proccess insert PALN STRAIT EQUITY from email start " + DateTime.Now.ToString("HH:mm:ss"));

                            Decimal total = 0;
                            foreach (var attachment in getMessage.FindAllAttachments())
                            {
                                var caption = attachment.ContentType.Name;
                                string ext = Path.GetExtension(attachment.ContentType.Name);

                                if (ext == ".pdf")
                                {
                                    string path_file1 = string.Format(path, caption);

                                    if (System.IO.File.Exists(path_file1))
                                    {
                                        System.IO.File.Delete(path_file1);
                                    }

                                    FileStream Stream = new FileStream(path_file1, FileMode.Create);
                                    BinaryWriter BinaryStream = new BinaryWriter(Stream);
                                    BinaryStream.Write(attachment.Body);
                                    BinaryStream.Close();
                                    using (var input = new OcrInput())
                                    {
                                        input.AddPdf(path_file1);
                                        var Result = Ocr.Read(input);
                                        var splitresult = Result.Text.Split(new string[] { "\r\nTotal Equity " }, StringSplitOptions.None);
                                        var hasil = splitresult[1].Split(' ')[0];
                                        System.IO.File.WriteAllText(AppDomain.CurrentDomain.BaseDirectory + "si1.txt", hasil);
                                        sendTelegram("-1001671146559", "Proses get total equity SFI : " + hasil + "\n" + DateTime.Now.ToString("hh:mm:ss"));
                                        File.Delete(path_file1);
                                        total = total + Convert.ToDecimal(hasil.Replace(",", "").Replace(".", ","));
                                    }
                                }
                            }
                            //insert db
                            var dt_val = dr.GetDataByDate(DateTime.Now.Date, Convert.ToDecimal(115));
                            if (dt_val.Count == 0)
                            {
                                dr.Insert(115, DateTime.Now.Date, total);
                                sendTelegram("-1001671146559", "Success insert PALN STRAIT EQUITY : $ " + total + "\nTimestamp " + DateTime.Now.ToString("HH:mm:ss"));
                            }
                            else
                            {
                                sendTelegram("-1001671146559", "PALN STRAIT already , EQUITY : $ " + dt_val[0].PALN + "\nTimestamp " + DateTime.Now.ToString("HH:mm:ss"));
                            }
                        }
                    }
                }
                catch (Exception x)
                {
                    sendTelegram("-1001671146559", "Proccess insert PALN VALBURY EQUITY from email failed: " + x.Message + " " + DateTime.Now.ToString("HH:mm:ss"));

                }
                await Task.Delay(60000, stoppingToken);
            }
        }

        public static Decimal readOCR(string filepath)
        {
            try
            {
                var Ocr = new IronTesseract();

                IronOcr.Installation.LicenseKey = "IRONOCR.PTKLIRINGBERJANGKAINDONESIA.IRO211213.9250.23127.312112-1EA71407D8-HWNZXEAJ3YDY3-N2BU5BL3WRB6-YYXFD7XQVLTB-YLR4RKFSW22L-F7OLHD7TWYX3-MUPLUQ-LVPA4EVNX2WIEA-PROFESSIONAL.SUB-DNMTTQ.RENEW.SUPPORT.13.DEC.2022";

                using (var input = new OcrInput())
                {
                    input.AddImage(filepath);
                    var Result = Ocr.Read(input);
                    var x = Result.Words;//43
                    var hasil = x[75].ToString().Replace("$", "");
                    hasil = hasil.Replace(",", "");
                    return Convert.ToDecimal(hasil);
                }
            }
            catch (Exception ex)
            {

                throw;
            }
        }
        private static void sendTelegram(string chatId, string body)
        {
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.DefaultConnectionLimit = 9999;

            var client = new RestClient("https://api.telegram.org/bot2144239635:AAFjcfn_GdHP4OkzzZomaZt4XbwpHDGyR-U/sendMessage?chat_id=" + chatId + "&text=" + body);
            RestRequest requestWa = new RestRequest("https://api.telegram.org/bot2144239635:AAFjcfn_GdHP4OkzzZomaZt4XbwpHDGyR-U/sendMessage?chat_id=" + chatId + "&text=" + body, Method.Get);
            requestWa.Timeout = -1;
            var responseWa = client.ExecutePostAsync(requestWa);
            Console.WriteLine(responseWa.Result.Content);
        }
        private static string monitoringServices(string servicename, string servicedescription, string servicelocation, string appstatus)
        {
            string jsonString = "{" +
                                "\"service_name\" : \"" + servicename + "\"," +
                                "\"service_description\": \"" + servicedescription + "\"," +
                                "\"service_location\":\"" + servicelocation + "\"," +
                                "\"app_status\":\"" + appstatus + "\"," +
                                "}";
            var client = new RestClient("http://10.10.10.99:84/api/ServiceStatus");

            RestRequest requestWa = new RestRequest("http://10.10.10.99:84/api/ServiceStatus", Method.Post);
            requestWa.Timeout = -1;
            requestWa.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            requestWa.AddParameter("data", jsonString);
            var responseWa = client.ExecutePostAsync(requestWa);
            return (responseWa.Result.Content);
        }

    }
}
