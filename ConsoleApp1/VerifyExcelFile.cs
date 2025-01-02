using System;
using Sylvan.Data.Excel;

namespace AzureWriter
{
    public class VerifyExcelFile
    {
        public bool VerifyFile(string filename)
        {
            bool status = true;

            // Open the file for processing
            ExcelDataReader edr = ExcelDataReader.Create(filename);
            do
            {
                // Validate the common Data
                edr.Read();
                status = ValidateCommonDataRow(status, edr);

                // Move to the start of the Data rows
                edr.Read();
                edr.Read();



                // Process all data rows
                while (edr.Read())
                {
                    VerifyRow(edr, ref status);
                    int a = 1;
                }
                // iterate sheets
                
            } while (edr.NextResult());

            return status;
        }

        private static bool ValidateCommonDataRow(bool status, ExcelDataReader edr)
        {
            // Verify the server location param
            string serverLocation = edr.GetString((int)Column.ServerLocation);
            if (!string.IsNullOrEmpty(serverLocation))
            {
                if (!Uri.IsWellFormedUriString(serverLocation, UriKind.Absolute))
                {
                    Console.WriteLine("Server location param is an invalid web location");
                    status = false;
                }
            }
            else
            {
                Console.WriteLine("Server Location not specified");
            }

            // Verify the PAT param
            string pat = edr.GetString((int)Column.Pat);
            if (string.IsNullOrEmpty(pat))
            {
                Console.WriteLine("Personal Access Token not specified");
                status = false;
            }

            // Verify the Project Name param
            string project = edr.GetString((int)Column.ProjectName);
            if (string.IsNullOrEmpty(project))
            {
                Console.WriteLine("Project name not specified");
                status = false;
            }

            // Verify the AreaPath param
            string areaPath = edr.GetString((int)Column.AreaPath);
            if (string.IsNullOrEmpty(areaPath))
            {
                Console.WriteLine("Area Path not specified");
                status = false;
            }

            return status;
        }

        // Validate a Data Row
        private void VerifyRow(ExcelDataReader edr, ref bool status)
        {
            // Validate the operation string is one of the excepted values
            if (!(edr.GetString((int)Column.Operation).Equals("Create") ||
                  edr.GetString((int)Column.Operation).Equals("Update") ||
                  edr.GetString((int)Column.Operation).Equals("Delete") ||
                  edr.GetString((int)Column.Operation).Equals("Rem") ||
                  string.IsNullOrEmpty(edr.GetString((int)Column.Operation))))
            {
                Console.WriteLine("Invalid Operation value");
                status = false;
            }
            else
            {
                // Only verify these operations as Rem and blank should be ignored
                var fred = edr.GetString((int)Column.Title);
                switch (edr.GetString((int)Column.Operation))
                {
                    case "Create":
                        // Verify that the mandatory Title is present
                        if (string.IsNullOrEmpty(edr.GetString((int)Column.Title)))
                        {
                            Console.WriteLine("A title is mandatory for a create operation");
                            status = false;
                        }

                        break;
                    case "Update":
                        // Verify that the WorkItem ID to update is present
                        if (!int.TryParse(edr.GetString((int)Column.ItemId), out _))
                        {
                            Console.WriteLine("WorkItem ID is not set for an update operation");
                            status = false;
                        }

                        break;
                    case "Delete":
                        // Verify that the WorkItem ID to update is present
                        if (!int.TryParse(edr.GetString((int)Column.ItemId), out _))
                        {
                            Console.WriteLine("WorkItem ID is not set for update operation");
                            status = false;
                        }

                        break;
                    default:
                        break;
                }

            }
        }
    }
}
