using System;
using System.IO;
using System.Text.RegularExpressions;
using OfficeOpenXml;


namespace ScoresParser
{
    internal class Program
    {
        public static int IndexOfAny2(string value, params string[] targets)
        {
            var index = -1;
            if (targets == null || targets.Length == 0) return index;

            foreach (var target in targets)
            {
                var targetIndex = value.IndexOf(target, 0, index > -1 ? index + target.Length - 1 : value.Length);

                if (targetIndex >= 0 && (index == -1 || targetIndex < index))
                {
                    index = targetIndex;
                    if (index == 0)
                    {
                        break;
                    }
                }
            }

            return index;
        }
        
        public static double[] ParseFile(string[] search, string filePath, double[] result, string reg)
        {
            StreamReader file = new StreamReader(filePath);
            string str;
            Regex regex = new Regex(reg);
            int i = 0;
            while ((str = file.ReadLine()) != null)
            {
                int index = IndexOfAny2(str, search);
                if (index > -1)
                {
                    str = str.Replace(".", ",");
                    MatchCollection matches = regex.Matches(str);
                    int k = 1;
                    foreach (Match match in matches)
                    {
                        if (search[0] == "Avr:")
                        {
                            if (k == 3 || k == 6)
                            {
                                result[i] = Double.Parse(match.Value);
                                i++;
                            }
                            k++;
                        }
                        else
                        {
                            result[i] = Double.Parse(match.Value);
                            i++;
                            k++;
                        }
                    }
                }
            }
            file.Close();
            return result;
        }
        
        
        public static void Main(string[] args)
        {
            string excelPath;
            Console.Write("Введите путь к таблице результатов(с файлом): ");
            excelPath = Console.ReadLine();
            FileInfo fileExcel = new FileInfo(excelPath);
            ExcelPackage pck = new ExcelPackage(fileExcel);
            var ws = pck.Workbook.Worksheets[0];
            
            
            int[] startFill =
            {
                85,     //Cinebench R15
                81,     //Cinebench R23
                51,     //LatencyMon
                91,     //7Zip
                75,     //Geekbench CPU
                77,     //OpenCL
                78,     //Vulkan
                95,     //GFXBench
                211,    //CSGO
                238,    //Osu
                186,    //Heaven Basic
                191,    //Heaver Maximum
                197,    //Heaven Custom
                42,     //Time Spy
                37,     //Fire Strike
                46,     //Night Raid
            };
            string resultsPath, patternsPath;
            Console.Write("Введите путь к файлам результатов(в конце обязательно \\): ");
            resultsPath = Console.ReadLine();
            Console.Write("Введите путь к файлам паттернов(\\)(оставить пустым если файлы в папки patterns...): ");
            if ((patternsPath = Console.ReadLine()) == "")
                patternsPath = resultsPath + "patterns\\";
            Console.Write("Enter letter of column: ");
            int column = Char.ToUpper(Convert.ToChar(Console.Read())) - 64;
            string[] resultsFiles =
            {
                "cb15.txt",
                "cb23.txt",
                "latmon.txt",
                "7zip.txt",
                "gbcpu.txt",
                "OpenCL.txt",
                "Vulkan.txt",
                "gfxbench.txt",
                "csgo_benchmark.txt",
                "osu_benchmark.txt",
                "basic.html",
                "maximum.html",
                "custom.html",
                "timespy.xml",
                "firestrike.xml",
                "nightraid.xml"
            };
            string[] patternsFiles =
            {
                "cb15_pattern.txt",
                "cb23_pattern.txt",
                "latmon_pattern.txt",
                "7zip_pattern.txt",
                "gbcpu_pattern.txt",
                "OpenCL_pattern.txt",
                "Vulkan_pattern.txt",
                "gfxbench_pattern.txt",
                "csgo_pattern.txt",
                "osu_pattern.txt",
                "heaven_pattern.txt",
                "heaven_pattern.txt",
                "heaven_pattern.txt",
                "timespy_pattern.txt",
                "firestrike_pattern.txt",
                "nightraid_pattern.txt"
            };
            string[] regexPattern =
            {
                "(\\d*),(\\d*)",                            //CB15
                " (\\d+)",                                  //CB23
                "(\\d*),(\\d*)",                            //LatencyMon
                " (\\d+)",                                  //7Zip
                " (\\d+)",                                  //GeekBech CPU
                " (\\d+)",                                  //OpenCL
                " (\\d+)",                                  //Vulkan
                "\\d(?>(?:(?>\\d+)\\w)?)\\d(?>[^ ]+)",      //GFXBench
                "(?>\\w+)\\,\\w(?= )",                      //CSGO
                "(?>\\w+)\\,\\w(?= )",                      //OSU!
                "(\\d+).(\\d+)",                            //Heaven
                "(\\d+).(\\d+)",                            //Heaven
                "(\\d+).(\\d+)",                            //Heaven
                "\\d(?>\\d+)",                              //3DMark
                "\\d(?>\\d+)",                              //3DMark
                "\\d(?>\\d+)"                               //3DMark
            };
            int[] countOfScores =
            {
                4,      //CB15
                2,      //CB23
                4,      //LatencyMon
                2,      //7Zip
                2,      //GeekBech CPU
                1,      //OpenCL
                1,      //Vulkan
                88,     //GFXBench
                20,     //CSGO(5*4)
                5,      //OSU!
                4,      //Heaven
                4,      //Heaven
                4,      //Heaven
                3,      //Time Spy
                4,      //Fire Strike
                3       //Night Raid
            };
            int countOfTests = countOfScores.Length;
            for (int i = 0; i < countOfTests; i++)
            {
                double[] result = new double[countOfScores[i]];
                string[] pattern = new string[countOfScores[i]];
                StreamReader file = new StreamReader(patternsPath + patternsFiles[i]);
                for (int j = 0; j < countOfScores[i]; j++)
                    pattern[j] = file.ReadLine();
                result = ParseFile(pattern, resultsPath + resultsFiles[i], result, regexPattern[i]);
                foreach (var res in result)
                {
                    Console.WriteLine(res);
                }
                Console.WriteLine();
                for (int j = 0; j < countOfScores[i]; j++)
                    ws.Cells[j + startFill[i], column].Value = result[j];
            }
            pck.Save();
        }
    }
}