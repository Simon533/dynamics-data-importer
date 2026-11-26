using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.PowerPlatform.Dataverse.Client;

namespace ProfessionImporter
{
    internal class Program
    {
        // Entity Logical Name
        private const string EntityLogicalName = "ncba_navinsprofession";

        // Optimal settings for Dataverse CreateMultiple
        private const int BatchSize = 250;           // Reduced from 500 to prevent payload timeouts
        private const int MaxDegreeOfParallelism = 1; // Set to 1 for sequential batches to avoid API throttling

        static async Task Main(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            string dataverseUrl = Environment.GetEnvironmentVariable("DATAVERSE_URL");
            if (string.IsNullOrWhiteSpace(dataverseUrl))
            {
                Console.WriteLine("[ERROR] Set the DATAVERSE_URL environment variable.");
                return;
            }

            string excelFilePath = Environment.GetEnvironmentVariable("PROFESSION_EXCEL_FILE");

            Console.Title = "Dynamics 365 NAV/INS Profession Importer";
            Console.WriteLine("==================================================");
            Console.WriteLine(" Connecting to Dataverse...");
            Console.WriteLine("==================================================");

            string connectionString =
                "AuthType=OAuth;" +
                $"Url={dataverseUrl};" +
                "AppId=51f81489-12ee-4a9e-aaae-a2591f45987d;" +
                "RedirectUri=http://localhost;" +
                "LoginPrompt=Auto;";

            using var serviceClient = new ServiceClient(connectionString);

            if (!serviceClient.IsReady)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"[FAIL] Connection failed: {serviceClient.LastError}");
                Console.ResetColor();
                return;
            }

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("[SUCCESS] Connected to Dataverse successfully!");
            Console.ResetColor();

            // Step 1: Fetch existing profession codes from Dataverse
            Console.WriteLine($"\nFetching existing '{EntityLogicalName}' records by 'ncba_code'...");
            var existingCodes = GetExistingProfessionCodes(serviceClient);
            Console.WriteLine($"Found {existingCodes.Count:N0} existing profession code(s) in Dataverse.");

            // Step 2: Read Excel file & filter out duplicates
            if (string.IsNullOrWhiteSpace(excelFilePath) || !File.Exists(excelFilePath))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[ERROR] Set PROFESSION_EXCEL_FILE to the path of the Excel file.");
                Console.ResetColor();
                return;
            }

            Console.WriteLine($"\nReading Excel file from: {excelFilePath}...");
            var professionRecords = ReadAndMapProfessionsFromExcel(excelFilePath, existingCodes);
            Console.WriteLine($"\nReady to import {professionRecords.Count:N0} NEW profession records.");

            if (professionRecords.Count == 0)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("[COMPLETED] No new records to import (all records were duplicates or empty).");
                Console.ResetColor();
                return;
            }

            // Step 3: Execute Batch Upload
            Console.WriteLine($"\nStarting batch upload into '{EntityLogicalName}' entity...");
            var stopwatch = Stopwatch.StartNew();

            await ImportProfessionsInBatchesAsync(serviceClient, professionRecords);

            stopwatch.Stop();
            Console.WriteLine("\n==================================================");
            Console.WriteLine($" Import Finished!");
            Console.WriteLine($" Total Elapsed Time: {stopwatch.Elapsed.TotalMinutes:F2} minutes");
            Console.WriteLine("==================================================");
        }

        private static HashSet<string> GetExistingProfessionCodes(ServiceClient client)
        {
            var existingCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            var query = new QueryExpression(EntityLogicalName)
            {
                ColumnSet = new ColumnSet("ncba_code"),
                PageInfo = new PagingInfo { Count = 5000, PageNumber = 1 }
            };

            while (true)
            {
                EntityCollection results = client.RetrieveMultiple(query);

                foreach (var entity in results.Entities)
                {
                    string codeVal = entity.GetAttributeValue<string>("ncba_code")?.Trim();
                    if (!string.IsNullOrWhiteSpace(codeVal))
                    {
                        existingCodes.Add(codeVal);
                    }
                }

                if (results.MoreRecords)
                {
                    query.PageInfo.PageNumber++;
                    query.PageInfo.PagingCookie = results.PagingCookie;
                }
                else break;
            }

            return existingCodes;
        }

        private static List<Entity> ReadAndMapProfessionsFromExcel(string filePath, HashSet<string> existingCodes)
        {
            var professions = new List<Entity>();
            var seenInFile = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = ExcelReaderFactory.CreateReader(stream);

            var dataset = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
            });

            DataTable table = dataset.Tables[0];

            // Debug Log: Print Detected Headers
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("[DEBUG] Detected Excel Columns:");
            foreach (DataColumn col in table.Columns)
            {
                Console.WriteLine($"  - Header: '{col.ColumnName}'");
            }
            Console.ResetColor();

            int duplicateCount = 0;
            int emptyCount = 0;

            foreach (DataRow row in table.Rows)
            {
                // Read Excel columns (flexible search)
                string excelCode = GetColumnValue(row, table, "Code", "ncba_code")?.Trim();
                string excelDescription = GetColumnValue(row, table, "Description", "ncba_name")?.Trim();

                if (string.IsNullOrWhiteSpace(excelCode) && string.IsNullOrWhiteSpace(excelDescription))
                {
                    emptyCount++;
                    continue;
                }

                // Reversed Schema Mapping:
                // Excel 'Description' -> Dataverse 'ncba_code' (Used as primary deduplication key)
                // Excel 'Code'        -> Dataverse 'ncba_name'
                string mappedCrmCode = excelDescription;
                string mappedCrmName = excelCode;

                // Fallback: If Description is empty, use Code as deduplication key
                if (string.IsNullOrWhiteSpace(mappedCrmCode))
                {
                    mappedCrmCode = mappedCrmName;
                }

                // Deduplication Check
                if (!string.IsNullOrWhiteSpace(mappedCrmCode))
                {
                    if (existingCodes.Contains(mappedCrmCode) || seenInFile.Contains(mappedCrmCode))
                    {
                        duplicateCount++;
                        continue; // Skip existing or repeated record
                    }

                    seenInFile.Add(mappedCrmCode);
                }

                var profession = new Entity(EntityLogicalName);

                if (!string.IsNullOrWhiteSpace(mappedCrmName))
                {
                    profession["ncba_name"] = mappedCrmName;
                }

                if (!string.IsNullOrWhiteSpace(mappedCrmCode))
                {
                    profession["ncba_code"] = mappedCrmCode;
                }

                professions.Add(profession);
            }

            Console.WriteLine($"[SUMMARY] Excel Total Rows: {table.Rows.Count}");
            if (emptyCount > 0)
            {
                Console.WriteLine($"[SUMMARY] Skipped Empty Rows: {emptyCount}");
            }
            if (duplicateCount > 0)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"[SUMMARY] Skipped Duplicate Records: {duplicateCount}");
                Console.ResetColor();
            }

            return professions;
        }

        private static async Task ImportProfessionsInBatchesAsync(ServiceClient client, List<Entity> records)
        {
            var batches = PartitionList(records, BatchSize);
            var options = new ParallelOptions { MaxDegreeOfParallelism = MaxDegreeOfParallelism };

            int completedCount = 0;

            await Parallel.ForEachAsync(batches, options, async (batch, token) =>
            {
                var createMultipleReq = new CreateMultipleRequest
                {
                    Targets = new EntityCollection(batch) { EntityName = EntityLogicalName }
                };

                createMultipleReq.Parameters["BypassCustomPluginExecution"] = true;

                try
                {
                    await client.ExecuteAsync(createMultipleReq, token);
                    int currentTotal = System.Threading.Interlocked.Add(ref completedCount, batch.Count);
                    Console.WriteLine($"Progress: {currentTotal:N0} / {records.Count:N0} uploaded...");
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"\n[ERROR IN BATCH] {ex.Message}");
                    Console.ResetColor();

                    // Fallback to individual creates on batch failure to isolate troublesome records
                    foreach (var record in batch)
                    {
                        try
                        {
                            await client.CreateAsync(record, token);
                        }
                        catch (Exception singleEx)
                        {
                            Console.ForegroundColor = ConsoleColor.Yellow;
                            Console.WriteLine($"[RECORD REJECTED] Code: '{record.GetAttributeValue<string>("ncba_code")}': {singleEx.Message}");
                            Console.ResetColor();
                        }
                    }
                }
            });
        }

        private static string GetColumnValue(DataRow row, DataTable table, params string[] possibleColumnNames)
        {
            foreach (var colName in possibleColumnNames)
            {
                foreach (DataColumn column in table.Columns)
                {
                    string cleanTarget = colName.Replace(" ", "").Trim();
                    string cleanColumn = column.ColumnName.Replace(" ", "").Trim();

                    if (cleanColumn.Equals(cleanTarget, StringComparison.OrdinalIgnoreCase))
                    {
                        if (!row.IsNull(column))
                        {
                            return row[column].ToString()?.Trim();
                        }
                    }
                }
            }
            return null;
        }

        private static List<List<T>> PartitionList<T>(List<T> source, int chunkSize)
        {
            var result = new List<List<T>>();
            for (int i = 0; i < source.Count; i += chunkSize)
            {
                result.Add(source.GetRange(i, Math.Min(chunkSize, source.Count - i)));
            }
            return result;
        }
    }
}