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
        private static DateTime _startTime;
        private static string _jsonFile;

        private static void Main()
        {
            Restart:
            Console.WriteLine("Enter filename:");

            _jsonFile = WorkPath + "\\" + Console.ReadLine() + ".json";
            _startTime = DateTime.Now;

            try
            {
                using (var file = File.OpenText(_jsonFile))
                {
                    using (var reader = new JsonTextReader(file))
                    {
                        var json = JToken.ReadFrom(reader);

                        var result = JsonConvert.DeserializeAnonymousType(json.ToString(), new object[] {new {}});

                        Console.WriteLine("Input file contains \t" + result.Length + " lines");

                        var x = 0;
                        foreach (JObject o in result)
                        {
                            foreach (var property in o)

                                if (property.Key != o.Properties().Last().Name)
                                    Sb.Append(property.Key + ",");
                                else
                                    Sb.Append(property.Key + Environment.NewLine);
                            x++;
                            if (x > 0) break;
                        }

                        Console.WriteLine("File header inserted.");

                        var y = 0;
                        var sw = new Stopwatch();
                        sw.Start();

                        foreach (JObject o in result)
                        {
                            y++;
                            Sb.Append("\"");

                            foreach (var property in o)
                                if (property.Key != o.Properties().Last().Name)
                                    Sb.Append(property.Value + "\",\"");
                                else
                                    Sb.Append(property.Value + "\"" + Environment.NewLine);

                            if (y%1000 == 0)
                                Console.WriteLine("Processed \t" + y + " lines \t\t " +
                                                  $"{(double) sw.ElapsedMilliseconds/1000:#,0.00}");
                        }
                        sw.Stop();
                    }
                }
                CreateCsv();
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

        private static void CreateCsv()
        {
            try
            {
                var csvFile = _jsonFile.Replace(".json", ".csv");

                if (File.Exists(csvFile)) File.Delete(csvFile);

                File.WriteAllText(csvFile, Sb.ToString());

                var lineCount = File.ReadLines(csvFile).Count();

                Console.WriteLine("Output file contains \t" + (lineCount - 1) + " lines.");
                Console.WriteLine("Processed in \t\t" + $"{DateTime.Now - _startTime}");
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
                    CreateCsv();
                }
            }
        }
    }
}