using Microsoft.VisualStudio.Services.WebApi.Patch.Json;

namespace AzureWriter
{
    public class DevOpsEntryDao
    {
        public string WorkItemType { get; set; }

        public string Operation { get; set; }

        public int WorkItemId { get; set; }
            public JsonPatchDocument PatchDocument { get; set; }
        }
    }
