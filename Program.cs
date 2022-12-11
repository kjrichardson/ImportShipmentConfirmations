using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Net;
using System.Text;
using System.IO;
using ZXing;
using ZXing.QrCode;
using ImageMagick;
using System.Drawing;
using System.Linq;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;


namespace ImportShipmentConfirmations
{
    class Program
    {
        private static Configuration AppConfig = new Configuration();
        private static CookieContainer Container = new CookieContainer();
        private static List<HttpStatusCode> ValidCodes;
        static void Main(string[] args)
        {
            //reads the configuration from the System. Uses user secrets on the system for the API settings.
            //for more information on how to use user secrets, visit https://learn.microsoft.com/en-us/aspnet/core/security/app-secrets?view=aspnetcore-7.0&tabs=windows
            var config = new ConfigurationBuilder()
                .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                .AddJsonFile("appsettings.json")
                .AddUserSecrets<Program>()
                .Build();
            config.Bind(AppConfig);
            //valid request codes that come from Acumatica. 
            ValidCodes = new List<HttpStatusCode>();
            ValidCodes.Add(HttpStatusCode.OK);
            ValidCodes.Add(HttpStatusCode.Created);
            ValidCodes.Add(HttpStatusCode.NoContent);


            try
            {
                Login();
                Import();

            }
            catch (Exception exc)
            {
                Console.WriteLine("Error - " + exc.Message);
                //optionally, log errors to file.
            }
            finally
            {
                Logout();
            }
        }

        /// <summary>
        /// Import the shipment confirmations from a folder to Acumatica
        /// </summary>
        public static void Import()
        {
            //need to set where ghostscript is installed.
            //tested with version 9.56.1 available at https://github.com/ArtifexSoftware/ghostpdl-downloads/releases/tag/gs9561
            //license AGPL or Commercial listed at https://www.ghostscript.com/licensing/
            MagickNET.SetGhostscriptDirectory(AppConfig.GhostScriptFolder);

            //get each file in the input folder
            foreach (var nextFile in Directory.GetFiles(AppConfig.Shipment.InputFolder))
            {
                FileInfo file = new FileInfo(nextFile);
                try
                {
                    string ShipmentNbr = "";
                    //if it is a pdf, try to load the barcode and get the shipment number from there.
                    if (file.Extension == ".pdf")
                    {
                        //load imageMagic configuration. Tested with ImageMagick-7.1.0-29.
                        //Download via https://imagemagick.org/script/download.php
                        //License listed at https://imagemagick.org/script/license.php

                        var settings = new MagickReadSettings();
                        // Settings the density to 300 dpi will create an image with a better quality
                        settings.Density = new Density(300, 300);


                        // load image into memory
                        byte[] imageMem = null;
                        using (MemoryStream memStream = new MemoryStream())
                        {
                            using (var images = new MagickImageCollection())
                            {
                                images.Read(nextFile, settings);
                                using (var horizontal = images.AppendVertically())
                                {
                                    horizontal.Write(memStream);
                                    imageMem = horizontal.ToByteArray();
                                }
                            }
                        }

                        //create bitmap from the image
                        Bitmap bitmap = new MagickImage(imageMem, settings).ToBitmap();
                        //initialize the barcode reader
                        BarcodeReader reader = new BarcodeReader();
                        //this flag makes it "try harder" but is much more accurate.
                        reader.Options.TryHarder = true;
                        //decodes all of the barcode son the page.
                        var results = reader.DecodeMultiple(bitmap);
                        //i am looking for the second barcode in the list. first is the customer number on our shipment confirmations.
                        if (results.Count() > 1)
                        {
                            ShipmentNbr = results[1].Text;
                        }
                    }
                    else
                    {
                        //attempt to read the shipment number from the file name of the scanned image.
                        ShipmentNbr = file.Name.Split('.')[0];
                        ShipmentNbr = ShipmentNbr.Split(' ')[0];
                        ShipmentNbr = ShipmentNbr.Split('-')[0];
                    }

                    var FileContents = File.ReadAllBytes(nextFile);
                    string NewFileName = "PackingSlip-" + ShipmentNbr + "-" + file.Name;
                    //PUT the file to the web service.
                    SubmitRequest(AppConfig.Shipment.Url + ShipmentNbr + "/files/" + NewFileName, "PUT", "", FileContents);

                    //move the processed file to an output folder.
                    string OutputFile = AppConfig.Shipment.OutputFolder + '\\' + file.Name;
                    if (File.Exists(OutputFile)) //if exists, add a datestamp to it.
                        OutputFile = OutputFile.Substring(0, OutputFile.Length - 4) + "-" + DateTime.Now.ToString("MMddyyyyhhmmss") + file.Extension;
                    File.Move(file.FullName, OutputFile);
                }

                catch (Exception exc)
                {
                    //log errors to a text file in a folder, and move the file that had a problem with it to that folder.
                    string OutputFile = AppConfig.Shipment.ProblemFolder + '\\' + file.Name;
                    if (File.Exists(OutputFile))
                        OutputFile = OutputFile.Substring(0, OutputFile.Length - 4) + "-" + DateTime.Now.ToString("MMddyyyyhhmmss") + file.Extension;
                    File.Move(file.FullName, OutputFile);
                    File.WriteAllText(OutputFile + "-error.txt", (exc.ToString() + "\n\n\n" + exc.StackTrace + "\n\n\n" + exc.InnerException?.ToString()));
                }

            }

        }

        /// <summary>
        /// Login to the Acumatica web service
        /// </summary>
        public static void Login()
        {
            JObject auth = new JObject();
            auth.Add("name", AppConfig.ApiUser);
            auth.Add("password", AppConfig.ApiPassword);

            using (var Response = SubmitRequest("auth/login/", "POST", JsonConvert.SerializeObject(auth)))
            {


                if (ValidCodes.Contains(Response.StatusCode) == false)
                {
                    string ResponseString = "";
                    using (StreamReader ResponseStreamReader = new StreamReader(Response.GetResponseStream()))
                    {
                        ResponseString = ResponseStreamReader.ReadToEnd();
                        ResponseStreamReader.Close();
                    }
                    throw new Exception("HTTP Status Error while Authenticating: " + Response.StatusCode + Environment.NewLine + "Response: " + ResponseString);
                }
            }

        }

        /// <summary>
        /// Logut from the Acumatica web service
        /// </summary>
        public static void Logout()
        {
            JObject auth = new JObject();
            auth.Add("name", AppConfig.ApiUser);
            auth.Add("password", AppConfig.ApiPassword);

            using (var Response = SubmitRequest("auth/logout/", "POST", JsonConvert.SerializeObject(auth)))
            {
                if (ValidCodes.Contains(Response.StatusCode) == false)
                {
                    string ResponseString = "";
                    using (StreamReader ResponseStreamReader = new StreamReader(Response.GetResponseStream()))
                    {
                        ResponseString = ResponseStreamReader.ReadToEnd();
                        ResponseStreamReader.Close();
                    }
                    throw new Exception("HTTP Status Error while Authenticating: " + Response.StatusCode + Environment.NewLine + "Response: " + ResponseString);
                }
            }
        }


        /// <summary>
        /// This is a helper method to submit a web request to acumatica.
        /// </summary>
        /// <param name="url">The path, after the root URL</param>
        /// <param name="Method">HTTP Method, PUT/GET/etc</param>
        /// <param name="RequestJson">Optional JSON to post as a body</param>
        /// <returns></returns>
        private static HttpWebResponse SubmitRequest(string url, string Method, string RequestJson = "", byte[] fileInfo = null)
        {
            var RequestURL = AppConfig.BaseURL + url;
            // Create an HTTP web request using the URL:
            HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(new Uri(RequestURL));
            //use the cookie container to save the session for the login/logout/etc
            request.CookieContainer = Container;
            request.ContentType = "application/json";
            request.Accept = "application/json";
            request.Method = Method;


            //submit JSON text as body
            if (String.IsNullOrWhiteSpace(RequestJson) == false)
            {
                using (Stream PostStream = request.GetRequestStream())
                {
                    byte[] byteArray = Encoding.UTF8.GetBytes(RequestJson);
                    PostStream.Write(byteArray, 0, byteArray.Length);
                    PostStream.Close();
                }
            }
            //submit binary data as body
            else if (fileInfo != null)
            {
                using (Stream PostStream = request.GetRequestStream())
                {
                    PostStream.Write(fileInfo, 0, fileInfo.Length);
                    PostStream.Close();
                }
            }

            var Response = (HttpWebResponse)request.GetResponse();
            return Response;
        }



        /*
           Licensed under the ImageMagick License (the "License"); you may not use
           this file except in compliance with the License.  You may obtain a copy
           of the License at

             https://imagemagick.org/script/license.php

           Unless required by applicable law or agreed to in writing, software
           distributed under the License is distributed on an "AS IS" BASIS, WITHOUT
           WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.  See the
           License for the specific language governing permissions and limitations
           under the License.  

           */
    }
}
