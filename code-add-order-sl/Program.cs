using ClosedXML.Excel;
using code_add_order_sl.Model;
using code_add_order_sl.Model.Response;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;

namespace code_add_order_sl
{
    internal class Program
    {
        private static string baseUrl = "http://192.168.0.219:50001/b1s/v1"; // Replace with your SAP Service Layer base URL


        static void Main(string[] args)
        {
            var sessionId = Login();

            var excelFilePath = @"C:\Temp\test.xlsx";

            if (!string.IsNullOrEmpty(sessionId))
            {
                // We first group the documents immediately after reading by Business Partner and DueDate
                var groupedDocuments = GroupDocuments(ReadExcelData(excelFilePath));

                if (groupedDocuments != null
                    && groupedDocuments.Count > 0)
                {
                    // For each grouped document - add a SAP Sales Order
                    foreach (var groupedDocument in groupedDocuments)
                    {
                        AddOrder(sessionId, groupedDocument.Value);
                    }
                }
                Logout(sessionId);
            }
            else
            {
                Console.WriteLine("Login failed.");
            }

            Console.ReadLine();
        }

        private static string Login()
        {
            // Request details
            string url = $"{baseUrl}/Login";
            LoginRequest loginRequest = new LoginRequest()
            {
                UserName = "manager",
                Password = "corexray",
                CompanyDB = "SBODemoUS"
            };

            // Serialize request body to JSON
            string jsonRequestBody = JsonConvert.SerializeObject(loginRequest);

            // Make the request
            var httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
            httpWebRequest.ContentType = "application/json";
            httpWebRequest.Method = "POST";
            httpWebRequest.KeepAlive = true;
            httpWebRequest.ServerCertificateValidationCallback += (sender, certificate, chain, sslPolicyErrors) => true;
            httpWebRequest.ServicePoint.Expect100Continue = false;

            using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
            {
                streamWriter.Write(jsonRequestBody);
            }

            try
            {
                // Call Service Layer
                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();

                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    var result = streamReader.ReadToEnd();

                    // Deserialize success response
                    var responseInstance = JsonConvert.DeserializeObject<LoginResponse>(result);

                    Console.WriteLine("Logged in successfully.");

                    return responseInstance.SessionId;
                }
            }
            catch (Exception ex)
            {
                // Unauthorized, etc.
                Console.WriteLine("Unexpected: " + ex.Message);
            }

            return null;
        }

        private static void AddOrder(string sessionId, List<Document> groupedDocuments)
        {
            string postingUrl = $"{baseUrl}/Orders";

            // Lines First
            var lines = new JArray();
            foreach (var document in groupedDocuments)
            {
                lines.Add(new JObject()
                        {
                            { "LineNum", document.LineNum },
                            { "ItemCode", document.ItemCode },
                            { "Quantity", document.Quantity }
                            // Add more line fields as required
                        });
            }
            
            // Order Header + Lines from above
            JObject payload = new JObject
            {
                { "CardCode", groupedDocuments[0].CardCode },
                { "DocDueDate", groupedDocuments[0].DocDueDate.ToString("yyyy-MM-dd") },
                // Add more header fields as required
                { "DocumentLines", lines }
            };

            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(postingUrl);
                request.Method = "POST";
                request.ContentType = "application/json";
                request.Accept = "application/json";
                request.Headers.Add("Cookie", $"B1SESSION={sessionId}");

                using (var streamWriter = new StreamWriter(request.GetRequestStream()))
                {
                    streamWriter.Write(payload.ToString());
                    streamWriter.Flush();
                    streamWriter.Close();
                }

                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                {
                    var result = "";
                    using (var streamReader = new StreamReader(response.GetResponseStream()))
                    {
                        result = streamReader.ReadToEnd();
                    }
                    
                    if (response.StatusCode == HttpStatusCode.Created)
                    {
                        JObject jsonObject = JObject.Parse(result);

                        // Extract the meaningful values
                        int docEntry = (int)jsonObject["DocEntry"];
                        int docNum = (int)jsonObject["DocNum"];

                        Console.WriteLine(string.Format("Sales Order added successfully with DocEntry = {0} and DocNum = {1}", docEntry, docNum));
                    }
                    else
                    {
                        Console.WriteLine("Failed to add Sales Order. Status code: " + response.StatusCode);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while adding the Sales Order: " + ex.Message);
            }
        }

        private static List<Document> ReadExcelData(string filePath)
        {
            var documents = new List<Document>();

            try
            {
                using (var workbook = new XLWorkbook(filePath))
                {
                    // First worksheet in the document
                    var worksheet = workbook.Worksheets.Worksheet(1);

                    // Skip the first row and read the rest
                    foreach (var row in worksheet.RowsUsed().Skip(1))
                    {
                        var document = new Document
                        {
                            CardCode = row.Cell(1).GetValue<string>(),
                            DocDueDate = DateTime.ParseExact(row.Cell(2).GetValue<string>(), "yyyyMMdd", CultureInfo.InvariantCulture),
                            LineNum = row.Cell(3).GetValue<int>(),
                            ItemCode = row.Cell(4).GetValue<string>(),
                            Quantity = row.Cell(5).GetValue<decimal>()
                        };

                        documents.Add(document);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading Excel file: {ex.Message}");
            }

            return documents;
        }

        private static Dictionary<string, List<Document>> GroupDocuments(List<Document> documents)
        {
            var groupedDocuments = new Dictionary<string, List<Document>>();

            if (documents != null
                && documents.Count > 0)
            {
                foreach (var document in documents)
                {
                    // Construct the key as CardCode + "|" + DocDueDate (DateTime converted to string)
                    var key = document.CardCode + "|" + document.DocDueDate.ToString("yyyy-MM-dd");

                    if (!groupedDocuments.ContainsKey(key))
                    {
                        // If the key does not exist, add a new entry with a new list containing the current document
                        groupedDocuments[key] = new List<Document> { document };
                    }
                    else
                    {
                        // If the key exists, add the current document to the existing list
                        groupedDocuments[key].Add(document);
                    }
                }
            }

            return groupedDocuments;
        }

        private static void Logout(string sessionId)
        {
            string logoutUrl = $"{baseUrl}/Logout";

            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(logoutUrl);
                request.Method = "POST";
                request.Accept = "application/json";
                request.Headers.Add("Cookie", $"B1SESSION={sessionId}");

                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                {
                    if (response.StatusCode == HttpStatusCode.NoContent
                        || response.StatusCode == HttpStatusCode.OK)
                    {
                        Console.WriteLine("Logged out successfully.");
                    }
                    else
                    {
                        Console.WriteLine("Logout failed. Status code: " + response.StatusCode);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during logout: " + ex.Message);
            }
        }
    }
}
