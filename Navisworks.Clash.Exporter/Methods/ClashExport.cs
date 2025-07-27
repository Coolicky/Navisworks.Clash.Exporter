using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Clash;
using ClosedXML.Excel;
using Navisworks.Clash.Exporter.ClashGrouping;
using Navisworks.Clash.Exporter.Data;
using Navisworks.Clash.Exporter.Extensions;
using Navisworks.Clash.Exporter.Models;
using Serilog;

namespace Navisworks.Clash.Exporter.Methods
{
    public class ClashExport
    {
        private static FileStream _stream;
        private static readonly ILogger Logger = Log.ForContext<ClashExport>();

        public static int ExportClashes(CommandLineOptions options)
        {
            try
            {
                var file = $"{Path.GetFileNameWithoutExtension(options.NavisworksFile)}.xlsx";
                var fileName = Path.Combine(options.ExportFolder, file);

                var historicalTable = NewHistoricalTable();
                if (options.SavePrevious && File.Exists(fileName))
                {
                    Logger.Information("Saving copy of a previous export....");
                    var directory = Directory.CreateDirectory(Path.Combine(options.ExportFolder, "Previous"));
                    var previousFileName = Path.Combine(directory.FullName,
                        file.Replace(".xlsx", $"_{DateTime.Now:yyyyMMdd}.xlsx"));
                    var newFile = Path.Combine(directory.FullName, previousFileName);
                    if (File.Exists(newFile)) File.Delete(newFile);
                    File.Copy(fileName, newFile);

                    historicalTable = GetHistoricalTable(fileName);
                }

                Logger.Information("Locking file {Name} for saving", fileName);
                _stream = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.Read);

                Logger.Information("Reading Clashes Manager from Navisworks File");
                var clashManager = Application.MainDocument.GetClash();
                if (!options.SkipRefresh)
                    RunClashes(clashManager);

                var clashTests = clashManager.TestsData;
                GroupClashes(clashTests, options);

                var dataTables = CreateDataTables(clashTests);
                AddHistoricalTable(dataTables, historicalTable);

                Logger.Information("Converting to Excel Workbook...");
                var excel = dataTables.ToExcel();
                _stream.Close();
                Logger.Information("File Name {Name}", fileName);
                Logger.Information("Saving the Excel File...");
                excel.SaveAs(fileName);

                Logger.Information("Successfully Completed");
            }
            catch (Exception e)
            {
                Logger.Error("Error while exporting clashes: {Error} \n {Stack}", e.Message, e.StackTrace);
                return 0;
            }
            finally
            {
                _stream?.Close();
            }

            return 1;
        }

        private static DataTable NewHistoricalTable()
        {
            var dt = new DataTable("History");

            dt.Columns.Add("TestName", typeof(string));
            dt.Columns.Add("New", typeof(int));
            dt.Columns.Add("Active", typeof(int));
            dt.Columns.Add("Reviewed", typeof(int));
            dt.Columns.Add("Approved", typeof(int));
            dt.Columns.Add("Resolved", typeof(int));
            dt.Columns.Add("Total", typeof(int));
            dt.Columns.Add("Date", typeof(DateTime));
            return dt;
        }

        private static DataTable GetHistoricalTable(string fileName)
        {
            try
            {
                var existingWorkbook = new XLWorkbook(fileName);
                var existingHistory = existingWorkbook.TryGetWorksheet("History", out var historical);
                if (existingHistory)
                {
                    if (CheckHistoricalTable(historical))
                        return historical.ToDataTable();
                }

                return NewHistoricalTable();
            }
            catch
            {
                return NewHistoricalTable();
            }
        }

        private static bool CheckHistoricalTable(IXLWorksheet historical)
        {
            var columns = historical.Columns().ToList();
            if (columns.Count != 8)
                return false;
            if (columns[0].Cell(1).Value.ToString() != "TestName")
                return false;
            if (columns[1].Cell(1).Value.ToString() != "New")
                return false;
            if (columns[2].Cell(1).Value.ToString() != "Active")
                return false;
            if (columns[3].Cell(1).Value.ToString() != "Reviewed")
                return false;
            if (columns[4].Cell(1).Value.ToString() != "Approved")
                return false;
            if (columns[5].Cell(1).Value.ToString() != "Resolved")
                return false;
            if (columns[6].Cell(1).Value.ToString() != "Total")
                return false;
            if (columns[7].Cell(1).Value.ToString() != "Date")
                return false;

            return true;
        }

        private static void AddHistoricalTable(List<DataTable> dataTables, DataTable historicalTable)
        {
            var summary = dataTables.First(r => r.TableName == "Test Summary");
            var dt = summary.Copy();
            dt.Columns.Add("Date", typeof(DateTime));
            foreach (DataRow row in dt.Rows)
            {
                row["Date"] = DateTime.Now;
            }

            foreach (DataRow row in dt.Rows)
            {
                historicalTable.Rows.Add(row.ItemArray);
            }

            dataTables.Add(historicalTable);
        }

        /// <summary>
        /// This will take 85% of the progress!
        /// </summary>
        public static List<DataTable> CreateDataTables(DocumentClashTests clashTests, Progress progress = null)
        {
            double currentProgress = 0;
            var summaries = new List<SummaryDto>();
            var tests = new List<ClashTestDto>();
            var groups = new List<ClashGroupDto>();
            var clashes = new List<ClashResultDto>();
            var elements = new List<ElementDto>();
            var comments = new List<CommentDto>();

            progress?.BeginSubOperation(0.9, "Reading Clash Data");
            Logger.Information("Reading Separate Clash Data");
            for (var index = 0; index < clashTests.Tests.Count; index++)
            {
                var savedItem = clashTests.Tests[index];
                if (!(savedItem is ClashTest clashTest)) continue;

                var left = clashTests.Tests.Count - index;
                var fraction = 1d / left;
                progress?.BeginSubOperation(fraction, "Extracting Clash Test Data...");
                
                Logger.Information("Processing Clash Test Results for: {Name}", clashTest.DisplayName);

                summaries.Add(new SummaryDto(clashTest));

                var dto = new ClashTestDto(clashTest);
                tests.Add(dto);
                comments.AddRange(dto.Comments);

                var clashResults = clashTest.Children
                    .Where(r => r.GetType() == typeof(ClashResult))
                    .Cast<ClashResult>()
                    .ToList();
                var clashGroups = clashTest.Children
                    .Where(r => r.GetType() == typeof(ClashResultGroup))
                    .Cast<ClashResultGroup>()
                    .ToList();

                var operationCount = clashResults.Count + clashGroups.Count;
                var progressFraction = 1d / clashTests.Tests.Count / operationCount;

                Logger.Information("No. Clash Groups Found {Number}", clashGroups.Count);

                foreach (var clashGroup in clashGroups)
                {
                    currentProgress += progressFraction;
                    progress?.Update(currentProgress.Clamp(0, 1));

                    var groupDto = new ClashGroupDto(clashGroup, savedItem.Guid);
                    groups.Add(groupDto);
                    comments.AddRange(groupDto.Comments);
                    clashResults.AddRange(clashGroup.Children
                        .Where(r => r.GetType() == typeof(ClashResult))
                        .Cast<ClashResult>()
                        .ToList());
                }

                Logger.Information("No. Clash Results Found {Number}", clashResults.Count);

                foreach (var clashResult in clashResults)
                {
                    currentProgress += progressFraction;
                    progress?.Update(currentProgress.Clamp(0, 1));

                    var resultDto = new ClashResultDto(clashResult, savedItem.Guid);
                    clashes.Add(resultDto);
                    comments.AddRange(resultDto.Comments);
                    elements.AddRange(resultDto.Items);
                }
                progress?.EndSubOperation(true);
            }
            progress?.EndSubOperation(true);

            Logger.Information("Converting Data...");
            var dataTables = new List<DataTable>
            {
                summaries.ToDataTable(),
                tests.ToDataTable(),
                groups.ToDataTable(),
                clashes.ToDataTable(),
                elements.ToElementDataTable(),
                comments.ToDataTable()
            };
            progress?.Update(0.85);
            return dataTables;
        }

        private static void GroupClashes(DocumentClashTests clashTests, CommandLineOptions options)
        {
            if (options.GroupBy == "None" || string.IsNullOrEmpty(options.GroupBy))
                return;

            var customTests = clashTests.Tests
                .Where(r => r.GetType() == typeof(ClashTest))
                .Cast<ClashTest>()
                .Select(r => new CustomClashTest(r))
                .ToList();

            var firstEnum = (GroupingMode)Enum.Parse(typeof(GroupingMode), options.GroupBy);
            GroupingMode secondEnum;
            if (options.ThenBy == "None" || string.IsNullOrEmpty(options.ThenBy))
                secondEnum = GroupingMode.None;
            else
                secondEnum = (GroupingMode)Enum.Parse(typeof(GroupingMode), options.ThenBy);

            foreach (var clashTest in customTests)
            {
                GroupingFunctions.GroupClashes(clashTest.ClashTest, firstEnum, secondEnum, options.KeepGroups);
            }
        }

        private static void RunClashes(DocumentClash clash)
        {
            foreach (var test in clash.TestsData.Tests)
            {
                if (test.GetType() != typeof(ClashTest)) continue;
                var clashTest = (ClashTest)test;
                Logger.Information("Running Clash: {Name}", clashTest.DisplayName);
                clash.TestsData.TestsRunTest(clashTest);
            }
        }
    }
}