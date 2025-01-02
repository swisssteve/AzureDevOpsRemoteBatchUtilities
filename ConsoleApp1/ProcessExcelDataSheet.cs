
using Sylvan.Data.Excel;
using Microsoft.VisualStudio.Services.WebApi.Patch.Json;
using Microsoft.VisualStudio.Services.WebApi.Patch;


namespace AzureWriter
{
    public enum Column
    {
        ServerLocation = 1,
        Pat = 3,
        ProjectName = 5,
        AreaPath = 7,
        Operation = 0,
        WorkItemType = 1,
        Title = 2,
        Description = 3,
        Severity = 4,
        Priority = 5,
        ParentId = 6,
        ItemId = 7,
        Tags = 8,
        State = 9,
        Reason = 10,
        AssignedTo = 11,
        IterationId = 12
    };

    public class ProcessExcelDataSheet
    {
        public AzureDevOpsDao GetExcelData(string filename)
        {
            AzureDevOpsDao azu = new AzureDevOpsDao();

            ExcelDataReader edr = ExcelDataReader.Create(filename);
            do
            {
                // First line in File is blank
                edr.Read();
                edr.NextResult();

                // Extract Common Info
                azu.ServerPath = edr.GetString((int)Column.ServerLocation);
                azu.Pat = edr.GetString((int)Column.Pat);
                azu.Project = edr.GetString((int)Column.ProjectName);
                azu.AreaPath = edr.GetString((int)Column.AreaPath);

                // Move to location in file for start of data to be processed, one row were Azure action
                edr.Read();
                edr.Read();

                // enumerate rows in current sheet.
                while (edr.Read())
                {
                    var type = edr.GetString((int)Column.Operation);

                    DevOpsEntryDao doe = new DevOpsEntryDao
                    {
                        Operation = type,
                        WorkItemType = edr.GetString((int)Column.WorkItemType)
                    };
                    JsonPatchDocument patchDocument = new JsonPatchDocument();

                    switch (type)
                    {
                        case "Create":
                            GeneratePatchDocument(azu, edr, patchDocument);
                            doe.PatchDocument = patchDocument;
                            azu.AzureData.Add(doe);
                            break;

                        case "Update":
                            doe.Operation = type;
                            doe.WorkItemId = edr.GetInt32((int)Column.ItemId);
                            GeneratePatchDocument(azu, edr, patchDocument);
                            doe.PatchDocument = patchDocument;
                            azu.AzureData.Add(doe);
                            break;

                        case "Delete":
                            doe.Operation = type;
                            doe.WorkItemId = edr.GetInt32((int)Column.ItemId);

                            //add fields and their values to your patch document
                            // Delete does not need a field value but is included so that the delete operation is processed correctly
                            patchDocument.Add(
                                new JsonPatchOperation()
                                {
                                    Operation = Operation.Add,
                                    Path = "/fields/System.Title",
                                    Value = edr.GetString((int)Column.Title)
                                }
                            );

                            doe.PatchDocument = patchDocument;
                            azu.AzureData.Add(doe);
                            break;

                        default:
                            break;
                    }
                }
                // iterate sheets
            } while (edr.NextResult());
            return azu;
        }

 
        private void GeneratePatchDocument(AzureDevOpsDao azu, ExcelDataReader edr, JsonPatchDocument patchDocument)
        {
            // Add fields to your patch document
            // The correct fields per operation have been previously validated to ensure they exist 

            if (!string.IsNullOrEmpty(edr.GetString((int)Column.Title)))
            { 
                patchDocument.Add(
                new JsonPatchOperation()
                {
                    Operation = Operation.Add,
                    Path = "/fields/System.Title",
                    Value = edr.GetString((int)Column.Title)
                }
            );
            }

            if (!string.IsNullOrEmpty(azu.AreaPath))
            {
                patchDocument.Add(
                    new JsonPatchOperation()
                    {
                        Operation = Operation.Add,
                        Path = "/fields/System.AreaPath",
                        Value = azu.AreaPath
                    }
                );
            }

            if (!string.IsNullOrEmpty(edr.GetString((int)Column.Description)))
            {
                patchDocument.Add(
                new JsonPatchOperation()
                {
                    Operation = Operation.Add,
                    Path = "/fields/System.Description",
                    Value = edr.GetString((int)Column.Description)
                }
            );
            }

            if (!string.IsNullOrEmpty(edr.GetString((int)Column.Priority)))
            {
                patchDocument.Add(
                    new JsonPatchOperation()
                    {
                        Operation = Operation.Add,
                        Path = "/fields/Microsoft.VSTS.Common.Priority",
                        Value = edr.GetString((int)Column.Priority)
                    }
                );
            }

            // If a Parent ID is specified then link this created Work Item
            if (edr.GetString((int)Column.ParentId).Length > 0)
            {
                patchDocument.Add(
                    new JsonPatchOperation
                    {// If a Parent ID is specified then link this created Work Item
                        Operation = Operation.Add,
                        Path = "/relations/-",
                        Value = new
                        {
                            rel = "System.LinkTypes.Hierarchy-Reverse",
                            url = azu.ServerPath + "_apis/wit/workItems/" + edr.GetString((int)Column.ParentId)
                        }
                    }
                );
            }

            if (!string.IsNullOrEmpty(edr.GetString((int)Column.Tags)))
            {
                patchDocument.Add(
                    new JsonPatchOperation()
                    {
                        Operation = Operation.Add,
                        Path = "/fields/System.Tags",
                        Value = edr.GetString((int)Column.Tags)
                    }
                );
            }

            if (!string.IsNullOrEmpty(edr.GetString((int)Column.Severity)))
            {
                patchDocument.Add(
                    new JsonPatchOperation()
                    {
                        Operation = Operation.Add,
                        Path = "/fields/System.Severity",
                        Value = edr.GetString((int)Column.Severity)
                    }
                );
            }
            if (!string.IsNullOrEmpty(edr.GetString((int)Column.State)))
            {
                patchDocument.Add(
                    new JsonPatchOperation()
                    {
                        Operation = Operation.Add,
                        Path = "/fields/System.State",
                        Value = edr.GetString((int)Column.State)
                    }
                );
            }

            if (!string.IsNullOrEmpty(edr.GetString((int)Column.Reason)))
            {
                patchDocument.Add(
                new JsonPatchOperation()
                {
                    Operation = Operation.Add,
                    Path = "/fields/System.Reason",
                    Value = edr.GetString((int)Column.Reason)
                }
            );
            }

            if (!string.IsNullOrEmpty(edr.GetString((int)Column.AssignedTo)))
            {
                patchDocument.Add(
                 new JsonPatchOperation()
                 {
                     Operation = Operation.Add,
                     Path = "/fields/System.AssignedTo",
                     Value = edr.GetString((int)Column.AssignedTo)
                 }
             );
            }

            if (!string.IsNullOrEmpty(edr.GetString((int)Column.IterationId)))
            {
                patchDocument.Add(
                 new JsonPatchOperation()
                 {
                     Operation = Operation.Add,
                     Path = "/fields/System.State",
                     Value = edr.GetString((int)Column.IterationId)
                 }
             );
            }
        }
        /*       private void GenerateModifyPatchDocument(AzureDevOpsDAO azu, ExcelDataReader edr, JsonPatchDocument patchDocument)
        {
            //add fields and their values to your patch document
            if (!string.IsNullOrEmpty(edr.GetString((int)Column.Title)))
            {
                patchDocument.Add(
                    new JsonPatchOperation()
                    {
                        Operation = Operation.Add,
                        Path = "/fields/System.Title",
                        Value = edr.GetString((int)Column.Title)
                    }
                );
            }

            if (!string.IsNullOrEmpty(edr.GetString((int)Column.Description)))
            {
                patchDocument.Add(
                    new JsonPatchOperation()
                    {
                        Operation = Operation.Add,
                        Path = "/fields/System.Description",
                        Value = edr.GetString((int)Column.Description)
                    }
                );
            }

            if (!string.IsNullOrEmpty(edr.GetString((int)Column.Priority)))
            {
                patchDocument.Add(
                    new JsonPatchOperation()
                    {
                        Operation = Operation.Add,
                        Path = "/fields/Microsoft.VSTS.Common.Priority",
                        Value = edr.GetString((int)Column.Priority)
                    }
                );
            }

            if (!string.IsNullOrEmpty(edr.GetString((int)Column.Severity)))
            {
                patchDocument.Add(
                    new JsonPatchOperation()
                    {
                        Operation = Operation.Add,
                        Path = "/fields/Microsoft.VSTS.Common.Severity",
                        Value = edr.GetString((int)Column.Severity)
                    }
                );
            }
            if (!string.IsNullOrEmpty(edr.GetString((int)Column.Tags)))
            {
                patchDocument.Add(
                    new JsonPatchOperation()
                    {
                        Operation = Operation.Add,
                        Path = "/fields/System.Tags",
                        Value = edr.GetString((int)Column.Tags)
                    }
                );
            }
        }
 */
    }
}


