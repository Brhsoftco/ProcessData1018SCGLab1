using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace ProcessData1018SCGLab1
{
    public class Program
    {
        public static string ResultsFolder { get; set; } = @"";
        public static string ResultsFile => $@"{ResultsFolder}\ResultCalculations.xlsx";
        public static DataTable[] ResultTrials { get; set; }

        public static void ProcessResults(DataTable[] trials)
        {
            try
            {
                if (trials?.Length > 0)
                {
                    var resultsWorkbook = new XLWorkbook();
                    var trialCount = 0;
                    foreach (var trial in trials)
                    {
                        if (trial?.Rows?.Count > 0)
                        {
                            var sets = new[] {
                                new DataTable(@"Distance1"),
                                new DataTable(@"Distance2")
                            };
                            var averageVelocities = new double[2];
                            var standardDeviations = new double[2];
                            var reasonable = new double[2];
                            for (var set = 0; set < 2; set++)
                            {
                                //table setup
                                sets[set].Columns.AddRange(new DataColumn[]
                                {
                                    new DataColumn(@"MeasureID", typeof(int)),
                                    new DataColumn(@"Time1", typeof(double)),
                                    new DataColumn(@"Time2", typeof(double)),
                                    new DataColumn(@"TimeD", typeof(double)),
                                    new DataColumn(@"Velocity", typeof(double))
                                });

                                var counter = 0;
                                var stateCell = 1 + (set * 2);
                                var distanceCell = 2 + (set * 2);
                                var deltaTList = new List<double>();
                                var velocities = new List<double>();
                                for (var i = 2; i < trial.Rows.Count; i += 2)
                                {
                                    if (i > 0)
                                    {
                                        var time1 = 0d;
                                        var time2 = 0d;
                                        var deltaT = 0d;
                                        var velocity = 0d;
                                        var stateCellValid = false;
                                        DataRow rowPrev = trial.Rows[i - 2];
                                        DataRow row = trial.Rows[i];
                                        if (int.TryParse(row.ItemArray[stateCell].ToString(), out _))
                                        {
                                            if (int.TryParse(rowPrev.ItemArray[stateCell].ToString(), out _))
                                            {
                                                stateCellValid = true;
                                            }
                                        }
                                        if (double.TryParse(rowPrev.ItemArray[0].ToString(), out var t1))
                                        {
                                            if (double.TryParse(row.ItemArray[0].ToString(), out var t2))
                                            {
                                                if (stateCellValid)
                                                {
                                                    var dT = t2 - t1;
                                                    time1 = t1;
                                                    time2 = t2;
                                                    deltaT = dT;
                                                    deltaTList.Add(dT);
                                                }
                                            }
                                        }
                                        if (double.TryParse(rowPrev.ItemArray[distanceCell].ToString(), out var d1))
                                        {
                                            if (double.TryParse(row.ItemArray[distanceCell].ToString(), out var d2))
                                            {
                                                if (int.TryParse(row.ItemArray[stateCell].ToString(), out _))
                                                {
                                                    if (stateCellValid)
                                                    {
                                                        var v = (d2 - d1) / deltaT;
                                                        velocity = v;
                                                        velocities.Add(v);
                                                    }
                                                }
                                            }
                                        }
                                        if (int.TryParse(row.ItemArray[stateCell].ToString(), out _))
                                        {
                                            if (stateCellValid)
                                            {
                                                sets[set].Rows.Add(counter + 1, time1, time2, deltaT, velocity);
                                            }
                                        }
                                        if (stateCellValid)
                                            counter++;
                                    }
                                }
                                averageVelocities[set] = velocities.Sum() / velocities.Count;
                                standardDeviations[set] = velocities.StandardDeviation();
                                reasonable[set] = Math.Round(100 - (standardDeviations[set] / averageVelocities[set]), 2);
                            }
                            //add sets
                            resultsWorkbook.AddWorksheet($"Trial{trialCount + 1}", trialCount);
                            resultsWorkbook.Worksheet(trialCount).Cell(1, 1).Value = $"Trial {trialCount + 1}";
                            resultsWorkbook.Worksheet(trialCount).Cell(1, 1).Style.Font.Bold = true;
                            resultsWorkbook.Worksheet(trialCount).Cell(1, 1).Style.Font.FontSize = 20d;
                            resultsWorkbook.Worksheet(trialCount).Cell(3, 1).Value = "Set 1";
                            resultsWorkbook.Worksheet(trialCount).Cell(3, 1).Style.Font.Bold = true;
                            resultsWorkbook.Worksheet(trialCount).Cell(3, 1).Style.Font.FontSize = 14d;
                            resultsWorkbook.Worksheet(trialCount).Cell(4, 1).InsertTable(sets[0]);
                            resultsWorkbook.Worksheet(trialCount).Cell(sets[0].Rows.Count + 5, 1).Value = @"Uncertainty";
                            resultsWorkbook.Worksheet(trialCount).Cell(sets[0].Rows.Count + 5, 1).Style.Font.Bold = true;
                            resultsWorkbook.Worksheet(trialCount).Cell(sets[0].Rows.Count + 5, 2).Value = standardDeviations[0];
                            resultsWorkbook.Worksheet(trialCount).Cell(sets[0].Rows.Count + 5, 2).Style.Font.Bold = true;
                            resultsWorkbook.Worksheet(trialCount).Cell(sets[0].Rows.Count + 6, 1).Value = @"Avg. Vel.";
                            resultsWorkbook.Worksheet(trialCount).Cell(sets[0].Rows.Count + 6, 1).Style.Font.Bold = true;
                            resultsWorkbook.Worksheet(trialCount).Cell(sets[0].Rows.Count + 6, 2).Value = averageVelocities[0];
                            resultsWorkbook.Worksheet(trialCount).Cell(sets[0].Rows.Count + 6, 2).Style.Font.Bold = true;
                            resultsWorkbook.Worksheet(trialCount).Cell(sets[0].Rows.Count + 7, 1).Value = @"Reasonable";
                            resultsWorkbook.Worksheet(trialCount).Cell(sets[0].Rows.Count + 7, 1).Style.Font.Bold = true;
                            resultsWorkbook.Worksheet(trialCount).Cell(sets[0].Rows.Count + 7, 2).Value = $"{reasonable[1]}%";
                            resultsWorkbook.Worksheet(trialCount).Cell(sets[0].Rows.Count + 7, 2).DataType = XLDataType.Number;
                            resultsWorkbook.Worksheet(trialCount).Cell(sets[0].Rows.Count + 7, 2).Style.NumberFormat.NumberFormatId = 10;
                            resultsWorkbook.Worksheet(trialCount).Cell(sets[0].Rows.Count + 7, 2).Style.Font.Bold = true;
                            resultsWorkbook.Worksheet(trialCount).Cell(sets[0].Rows.Count + 9, 1).Value = @"Set 2";
                            resultsWorkbook.Worksheet(trialCount).Cell(sets[0].Rows.Count + 9, 1).Style.Font.Bold = true;
                            resultsWorkbook.Worksheet(trialCount).Cell(sets[0].Rows.Count + 9, 1).Style.Font.FontSize = 14d;
                            resultsWorkbook.Worksheet(trialCount).Cell(sets[0].Rows.Count + 10, 1).InsertTable(sets[1]);
                            resultsWorkbook.Worksheet(trialCount).Cell(sets[0].Rows.Count + sets[1].Rows.Count + 11, 1).Value = @"Uncertainty";
                            resultsWorkbook.Worksheet(trialCount).Cell(sets[0].Rows.Count + sets[1].Rows.Count + 11, 1).Style.Font.Bold = true;
                            resultsWorkbook.Worksheet(trialCount).Cell(sets[0].Rows.Count + sets[1].Rows.Count + 11, 2).Value = standardDeviations[1];
                            resultsWorkbook.Worksheet(trialCount).Cell(sets[0].Rows.Count + sets[1].Rows.Count + 11, 2).Style.Font.Bold = true;
                            resultsWorkbook.Worksheet(trialCount).Cell(sets[0].Rows.Count + sets[1].Rows.Count + 12, 1).Value = @"Avg. Vel.";
                            resultsWorkbook.Worksheet(trialCount).Cell(sets[0].Rows.Count + sets[1].Rows.Count + 12, 1).Style.Font.Bold = true;
                            resultsWorkbook.Worksheet(trialCount).Cell(sets[0].Rows.Count + sets[1].Rows.Count + 12, 2).Value = averageVelocities[1];
                            resultsWorkbook.Worksheet(trialCount).Cell(sets[0].Rows.Count + sets[1].Rows.Count + 12, 2).Style.Font.Bold = true;
                            resultsWorkbook.Worksheet(trialCount).Cell(sets[0].Rows.Count + sets[1].Rows.Count + 13, 1).Value = @"Reasonable";
                            resultsWorkbook.Worksheet(trialCount).Cell(sets[0].Rows.Count + sets[1].Rows.Count + 13, 1).Style.Font.Bold = true;
                            resultsWorkbook.Worksheet(trialCount).Cell(sets[0].Rows.Count + sets[1].Rows.Count + 13, 2).Value = $"{reasonable[1]}%";
                            resultsWorkbook.Worksheet(trialCount).Cell(sets[0].Rows.Count + sets[1].Rows.Count + 13, 2).DataType = XLDataType.Number;
                            resultsWorkbook.Worksheet(trialCount).Cell(sets[0].Rows.Count + sets[1].Rows.Count + 13, 2).Style.NumberFormat.NumberFormatId = 10;
                            resultsWorkbook.Worksheet(trialCount).Cell(sets[0].Rows.Count + sets[1].Rows.Count + 13, 2).Style.Font.Bold = true;
                            trialCount++;
                        }
                    }
                    resultsWorkbook.SaveAs(ResultsFile);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        public static void PauseProgram(bool newLine = true)
        {
            if (newLine)
            {
                Console.WriteLine();
            }
            Console.WriteLine(@"Press any key to continue...");
            Console.ReadKey();
        }

        public static void Main(string[] args)
        {
            if (args.Length > 0)
            {
                if (!string.IsNullOrWhiteSpace(args[0]))
                {
                    if (Directory.Exists(args[0]))
                    {
                        ResultsFolder = args[0];

                        var results = new List<DataTable>();
                        foreach (var s in Directory.GetFiles(ResultsFolder))
                        {
                            if (Path.GetExtension(s) == @".csv")
                            {
                                var d = DataHelpers.GetDataTableFromCsv(s);
                                if (d?.Rows?.Count > 0)
                                {
                                    results.Add(d);
                                    Console.WriteLine($"Processed {Path.GetFileName(s)}");
                                }
                            }
                        }
                        ResultTrials = results.ToArray();
                        if (results.Count > 0)
                        {
                            Console.WriteLine($"Found and loaded {results.Count} trials!");
                            Console.WriteLine();
                            Console.WriteLine(@"Processing data...");
                            ProcessResults(ResultTrials);
                            Console.WriteLine();
                            Console.WriteLine($"Exported loaded data: {ResultsFile}");
                        }
                        else
                        {
                            Console.WriteLine("No results were found in the provided folder");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Invalid results folder");
                    }
                }
                else
                {
                    Console.WriteLine("Invalid results folder");
                }
            }
            else
            {
                Console.WriteLine("No results folder specified");
            }
            PauseProgram();
        }
    }
}