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
    }

    public class Content
    {
        public List<Item> res;

    }
}
