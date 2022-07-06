using System;
using System.Collections.Generic;

namespace MedicalAssurancePolicyData
{
    public class MedicalAssuranceHtmlResponse
    {
        public int code { get; set; }

        public string message { get; set; }

        public Content content { get; set; }
    }
    public class Item
    {
        public int id { get; set; }
        public int? createdUserId { get; set; }
        public DateTime? createdDate { get; set; }
        public int? modifiedUserId { get; set; }
        public DateTime? modifiedDate { get; set; }
        public bool? isDeleted { get; set; }
        public int? versionNumber { get; set; }
        public string province_code { get; set; }
        public string city_code { get; set; }
        public string content { get; set; }
        public string content_for_search { get; set; }
        public bool? is_hot { get; set; }
        public int? order { get; set; }
        public string catalog_type { get; set; }
        public string type_code { get; set; }
        public string display_name { get; set; }
    }

    public class Content
    {
        public List<Item> res;

    }
}
