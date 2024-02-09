using System;
using System.Collections.Generic;
using System.IO;
using DbfDataReader;
using OfficeOpenXml;

namespace DBFToXLSXConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Prompt the user to enter the directory path
            Console.WriteLine("Enter the directory path containing DBF files:");
            string dirPath = Console.ReadLine();

            // Check if the directory path is valid
            if (string.IsNullOrEmpty(dirPath) || !Directory.Exists(dirPath))
            {
                Console.WriteLine("Invalid directory path. Exiting...");
                return;
            }

            // Define the subdirectory name and create it
            string subdir = "Converted_XLSX";
            string subdirPath = Path.Combine(dirPath, subdir);
            Directory.CreateDirectory(subdirPath);

            // List all DBF files in the directory
            string[] dbfFiles = Directory.GetFiles(dirPath, "*.dbf");

            // Initialize a list to store files that had errors
            List<string> errorLog = new List<string>();

            // Process each DBF file
            foreach (string dbfFile in dbfFiles)
            {
                try
                {
                    // Read the DBF file
                    using (var reader = new DbfDataReader.DbfDataReader(dbfFile))
                    {
                        // Define XLSX file path
                        string fileName = Path.GetFileName(dbfFile);
                        string xlsxFile = Path.Combine(subdirPath, Path.ChangeExtension(fileName, ".xlsx"));

                        // Write the records to an XLSX file
                        using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(xlsxFile)))
                        {
                            var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                            int rowIndex = 1;
                            while (reader.Read())
                            {
                                for (int i = 0; i < reader.FieldCount; i++)
                                {
                                    worksheet.Cells[rowIndex, i + 1].Value = reader.GetValue(i);
                                }
                                rowIndex++;
                            }
                            package.Save();
                        }

                        // Output progress
                        Console.WriteLine($"Converted: {dbfFile} to {xlsxFile}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error reading {dbfFile}: {ex.Message}");
                    errorLog.Add(dbfFile);
                }
            }

            // Display the error log at the end
            if (errorLog.Count > 0)
            {
                Console.WriteLine("The following files had errors and were skipped:");
                foreach (var errorFile in errorLog)
                {
                    Console.WriteLine(errorFile);
                }
            }
            else
            {
                Console.WriteLine("All files processed successfully with no errors!");
            }
        }
    }
}
