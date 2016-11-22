using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace JSON2CSV
{
    public static class Program
    {
        private static readonly string WorkPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        private static readonly StringBuilder Sb = new StringBuilder();

        private static void Main()
        {
            Restart:
            Console.WriteLine("Enter filename:");
            var jsonFile = WorkPath + "\\" + Console.ReadLine() + ".json";

            #region Definition

            //var jsonDef = new object[]
            //{
            //    new
            //    {
            //        DRID = 0,
            //        Rank = 0,
            //        Publications = 0,
            //        RecentDate = "",
            //        NPI = "",
            //        REVIEWER_ID = "",
            //        Specialty = "",
            //        First_Name = "",
            //        Last_Name = "",
            //        Address = "",
            //        City = "",
            //        State = "",
            //        Zipcode = "",
            //        Phone = "",
            //        Fax = "",
            //        Email_Address = "",
            //        County = "",
            //        Company_Name = "",
            //        Latitude = "",
            //        Longitude = "",
            //        Timezone = "",
            //        Website = "",
            //        Gender = "",
            //        Credentials = "",
            //        Taxonomy_Code = "",
            //        Taxonomy_Classification = "",
            //        Taxonomy_Specialization = "",
            //        License_Number = "",
            //        License_State = "",
            //        Medical_School = "",
            //        Residency_Training = "",
            //        Graduation_Year = "",
            //        Patients = 0,
            //        Claims = 0,
            //        Prescriptions = 0,
            //        Country = ""
            //    }
            //}; 

            #endregion

            var startTime = DateTime.Now;

            try
            {
                using (var file = File.OpenText(jsonFile))

                using (var reader = new JsonTextReader(file))
                {
                    var json = JToken.ReadFrom(reader);

                    var result = JsonConvert.DeserializeAnonymousType(json.ToString(), new object[] {new {}});

                    Console.WriteLine("Input file contains \t" + result.Length + " lines");

                    var x = 0;
                    foreach (JObject o in result)
                    {
                        foreach (var property in o)
                        {
                            Sb.Append(property.Key + ",");
                        }
                        Sb.Replace(Sb.ToString(), Sb.ToString().Substring(0, Sb.Length - 1));
                        Sb.Append(Environment.NewLine);
                        x++;
                        if (x > 0) break;
                    }

                    foreach (JObject o in result)
                    {
                        Sb.Append("\"");

                        foreach (var property in o)
                        {
                            Sb.Append(property.Value + "\",\"");
                        }
                        Sb.Replace(Sb.ToString(), Sb.ToString().Substring(0, Sb.Length - 2));
                        Sb.Append(Environment.NewLine);
                    }

                    CreateCsv(jsonFile, Sb, startTime);
                }
            }
            catch (Exception ex)
            {
                if (ex.GetType().ToString().Contains("FileNotFound"))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("File not found. Re-enter the name.");
                    Console.ResetColor();
                    goto Restart;
                }

            }
        }

        private static void CreateCsv(string jsonFile, StringBuilder sb, DateTime startTime)
        {
            try
            {
                var csvFile = jsonFile.Replace(".json", ".csv");

                if (File.Exists(csvFile)) File.Delete(csvFile);

                File.WriteAllText(csvFile, sb.ToString());

                var lineCount = File.ReadLines(csvFile).Count();

                Console.WriteLine("Output file contains \t" + (lineCount - 1) + " lines.");
                Console.WriteLine("Processed in \t\t" + $"{DateTime.Now - startTime}");
                Console.WriteLine("Press any key to open output file.");

                Console.ReadKey(true);

                Process.Start(csvFile);
            }
            catch (Exception ex)
            {
                if (ex.ToString().Contains("The process cannot access the file"))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("File is open. Please close the file and press any key.");
                    Console.ResetColor();
                    Console.ReadKey(true);
                    CreateCsv(jsonFile, Sb, startTime);
                }
            }
        }
    }
}
