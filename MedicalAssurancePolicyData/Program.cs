using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json.Serialization;
using Newtonsoft.Json;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Reflection;
using System.ComponentModel;

namespace MedicalAssurancePolicyData
{
    class Program
    {
        static List<NameCodePair> Provinces;
        static List<NameCodePair> Cities;

        static string _authorization = "Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ1c2VySWQiOjU3LCJ1c2VyIjp7ImlkIjo1NywiY3JlYXRlZFVzZXJJZCI6NDgsImNyZWF0ZWREYXRlIjoiMjAyMi0wMS0xOVQwMTozMzo0Ny40NjdaIiwibW9kaWZpZWRVc2VySWQiOm51bGwsIm1vZGlmaWVkRGF0ZSI6bnVsbCwiaXNEZWxldGVkIjpmYWxzZSwidmVyc2lvbk51bWJlciI6MSwidXNlcm5hbWUiOiJDMjE3MzU1IiwiZW1haWwiOiJ0aWFuX2ppZUBuZXR3b3JrLmxpbGx5LmNuIiwicmVhbG5hbWUiOiJUaWFuSmllIiwiY29tbWVudCI6bnVsbH0sInR5cGUiOiJBWlVSRSIsInN1YiI6IjNkOGI0MDc4LWNiZTMtNDY1Ni05NDQ1LWU3NzIyNDBhNDMzNiIsImlhdCI6MTY1MjQzMDU1MSwiZXhwIjoxNjUyNDQ0OTUxfQ.9QpIV3e253qLEgXPJEkxufIvsZHQaiYSrgIfIS5Nuuo";
        static string _cookies = "connect.sid=s%3AW2K6T-2IwQTtQXAgq6XwqnmNKYe8qQAS.OLUfhTtx3eK4pVLvItXxUWd6eWK5Eo8qHd6BfgAGgxE; passport-aad.1652428454837.dca97261b6b899957e21816b5bc473494bb83657bd564f2b06887b2b8f1f4730c66b0862594499affce57c7a452a3666cd645f204fe86e1023216dd868b783696efc2a228e67d39c9fac1e3b4e51f031854bd9c74f0560.6102863c817b095343e22e3dc7174982=0";
        static string _host = "https://ipatientadmin.qa.lilly.cn";

        //static string _authorization = "Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ1c2VySWQiOjQ1LCJ1c2VyIjp7ImlkIjo0NSwiY3JlYXRlZFVzZXJJZCI6NCwiY3JlYXRlZERhdGUiOiIyMDIxLTA3LTE0VDA1OjMxOjM2Ljk5N1oiLCJtb2RpZmllZFVzZXJJZCI6bnVsbCwibW9kaWZpZWREYXRlIjpudWxsLCJpc0RlbGV0ZWQiOmZhbHNlLCJ2ZXJzaW9uTnVtYmVyIjoxLCJ1c2VybmFtZSI6IkMyMTczNTUiLCJlbWFpbCI6InRpYW5famllQG5ldHdvcmsubGlsbHkuY24iLCJyZWFsbmFtZSI6IkppZSBUaWFuIiwiY29tbWVudCI6bnVsbH0sInR5cGUiOiJBWlVSRSIsInN1YiI6ImIyOTRkOWVlLTU5ZjctNDIwZi04ZTllLWFjOTk0OTRiNDliYSIsImlhdCI6MTY1MjQzMjAwOSwiZXhwIjoxNjUyNDUzNjA5fQ.Wgz8yfTxHQEM24GGZjMr-ZKNiROp_do7VC5DMmo8FKk";
        //static string _cookies = "connect.sid=s%3ABvxk3AFkBZqBG_KBGdGk3Kpbuywr5h_t.WAwdAn41eNgGNfMilCx8%2F%2BG%2B8uRozZKWCd00EoLPANs";
        //static string _host = "https://ipatientadmin.lilly.cn";

        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            //ReadNewExcelAndParseData(args[0]);
            ReadNewExcelAndParseData("/Users/tianjie/workspaces/LCMedicalAssurancePolicyData/MedicalAssurancePolicyData/data/20220513/DrugReimbursement.xlsx");
        }


        private static void ReadNewExcelAndParseData(string folder)
        {
            // 打开某个目录
            List<MedicalAssurancePolicy> medicalAssurancePolicies;

            var filename = "all.json";
            if (File.Exists(filename))
            {
                using (StreamReader sr = new StreamReader(filename))
                {
                    var str = sr.ReadToEnd();
                    medicalAssurancePolicies = JsonConvert.DeserializeObject<List<MedicalAssurancePolicy>>(str);
                }
            }
            else
            {
                medicalAssurancePolicies = AnalyzeNewExcels(folder);
            }

            using (StreamReader sr = new StreamReader("province.json"))
            {
                var str = sr.ReadToEnd();
                Provinces = JsonConvert.DeserializeObject<List<NameCodePair>>(str);
            }

            using (StreamReader sr = new StreamReader("city.json"))
            {
                var str = sr.ReadToEnd();
                Cities = JsonConvert.DeserializeObject<List<NameCodePair>>(str);
            }

            // 处理完了，然后就可以整理成需要的内容了
            // TODO: 处理html内容，然后直接post
            //Console.WriteLine("Deleting Olumiant contents....");
            //DeleteAll("Olumiant", "normal");
            //PostContentByProduct(medicalAssurancePolicies, "艾乐明", "Olumiant", "", "normal");
            //Console.WriteLine("Deleting Taltz contents....");
            //DeleteAll("Taltz", "normal");
            //PostContentByProduct(medicalAssurancePolicies, "拓咨", "Taltz", "", "normal");

            //Console.WriteLine("Deleting Ve contents....");
            //DeleteAll("Verzenios", "countryHealthInsurance");
            //PostContentByProduct(medicalAssurancePolicies, "唯择", "Verzenios", "true", "countryHealthInsurance");

        }

        private static void PostContentByProduct(List<MedicalAssurancePolicy> medicalAssurancePolicies, string productName, string productCode, string isHot, string typeCode)
        {
            foreach (var policy in medicalAssurancePolicies.Where(a => a.ProductName == productName))
            {
                var html = CreateNewMedicalAssurancePolicyContent(policy);
                var province = policy.Province;
                var city = policy.City;
                if (string.IsNullOrEmpty(province) || string.IsNullOrEmpty(city))
                {
                    Console.WriteLine(" post failed: {0} - {1}", province, city);
                    continue;
                }
                try
                {
                    var provinceCode = Provinces.FirstOrDefault(a => a.Name == province).Code;
                    var cityCode = Cities.FirstOrDefault(a => a.Name == city).Code;

                    // 做HttpClient去post
                    Console.WriteLine(" Simulate POST: {0} - {1}, HTML", province, city);

                    var url = _host + "/api/v1/wechatMedical/ueditor?accountid=1";
                    HttpClient httpClient = new HttpClient();
                    html = html.Replace("\"", "'").Replace("\n", "");
                    isHot = string.IsNullOrEmpty(isHot)?(policy.IsHot == "是" ? "true" : "false"):isHot;
                    httpClient.DefaultRequestHeaders.Add("authorization", _authorization);
                    if (policy.Order == "-")
                    {
                        policy.Order = "0";
                    }
                    var postString = $"{{\"cmd\":\"MEDICAL_INSURANCE_POLICY\",\"data\":{{\"catalog_type\":\"{productCode}\",\"province_code\":\"{provinceCode}\",\"city_code\":\"{cityCode}\",\"content\":\"{html}\",\"is_hot\":{isHot},\"type_code\":\"{typeCode}\",\"order\":{policy.Order}}}}}";
                    postString = postString.Replace("\\", "\\\\");
                    var stringContent = new StringContent(postString, Encoding.UTF8, "application/json");
                    stringContent.Headers.Add("cookie", _cookies);
                    //stringContent.Headers.Add("Content-Type", "application/json");
                    var response = httpClient.PostAsync(url, stringContent).Result;
                    var result = response.Content.ReadAsStringAsync().Result;

                    if(response.StatusCode != System.Net.HttpStatusCode.Created)
                    {
                        Console.WriteLine(" post failed: {0} - {1}", province, city);
                    }
                }
                catch (Exception ex)
                {
                    // do nothing
                    Console.WriteLine(" Exceptions: {0} - {1}, {2}", province, city, ex.Message);
                    continue;
                }
            }
        }

        private static List<MedicalAssurancePolicy> AnalyzeNewExcels(string filename)
        {
            List<MedicalAssurancePolicy> medicalAssurancePolicies = new List<MedicalAssurancePolicy>();

            // 开始处理数据，读取所有数据，最后再过滤不同品牌
            using (FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read))
            {
                var excelFile = new XSSFWorkbook(fs);
                var sheet = excelFile.GetSheetAt(0);
                // 忽略tab的名字，到表格里找对应的省市
                // Console.WriteLine("  - " + sheet.SheetName);

                // 第一行数据
                var firstRow = sheet.GetRow(0);
                var province = "";
                var city = "";
                for (var rowNumber = 1; rowNumber < 10000; rowNumber++)
                {
                    IRow row = null;
                    try
                    {
                        row = sheet.GetRow(rowNumber);
                        // 判断文件结束
                        if (row == null || String.IsNullOrEmpty(row.Cells[1].StringCellValue))
                        {
                            break;
                        }
                    }
                    catch (Exception)
                    {
                        break;
                    }

                    // 每一行，数据
                    var medicalAssurancePolicy = new MedicalAssurancePolicy();

                    try
                    {
                        medicalAssurancePolicy.TreatmentArea = GetCellValue(row, 1);
                        medicalAssurancePolicy.ChemicalName = GetCellValue(row, 2);
                        medicalAssurancePolicy.ProductName = GetCellValue(row, 3);
                        medicalAssurancePolicy.DosageForm = GetCellValue(row, 4);
                        medicalAssurancePolicy.ApprovaledText = GetCellValue(row, 5);
                        medicalAssurancePolicy.NRDLStatus = GetCellValue(row, 6);
                        medicalAssurancePolicy.Specification = GetCellValue(row, 7);
                        medicalAssurancePolicy.ProductionEnterpriseName = GetCellValue(row, 8);
                        medicalAssurancePolicy.Indications = GetCellValue(row, 9);
                        medicalAssurancePolicy.Province = GetCellValue(row, 10);
                        medicalAssurancePolicy.City = GetCellValue(row, 11);
                        medicalAssurancePolicy.ReimbursementRatio = GetCellValue(row, 12);
                        medicalAssurancePolicy.InsuranceType = GetCellValue(row, 13);
                        medicalAssurancePolicy.ReimbursementTreatment = GetCellValue(row, 14);
                        medicalAssurancePolicy.OutpatientStartPayLine = GetCellValue(row, 15);
                        medicalAssurancePolicy.OutpatientEndPayLine = GetCellValue(row, 16);
                        medicalAssurancePolicy.HospitalizationStartPayLine = GetCellValue(row, 17);
                        medicalAssurancePolicy.HospitalizationEndPayLine = GetCellValue(row, 18);
                        medicalAssurancePolicy.PolicyLink = GetCellValue(row, 19);
                        medicalAssurancePolicy.InsuranceTypeTwo = GetCellValue(row, 20);
                        medicalAssurancePolicy.ReimbursementTreatmentTwo = GetCellValue(row, 21);
                        medicalAssurancePolicy.OutpatientStartPayLineTwo = GetCellValue(row, 22);
                        medicalAssurancePolicy.OutpatientEndPayLineTwo = GetCellValue(row, 23);
                        medicalAssurancePolicy.HospitalizationStartPayLineTwo = GetCellPercentageValue(row, 24);
                        medicalAssurancePolicy.HospitalizationEndPayLineTwo = GetCellValue(row, 25);
                        medicalAssurancePolicy.PolicyLinkTwo = GetCellValue(row, 26);
                        medicalAssurancePolicy.InsuranceTypeThree = GetCellValue(row, 27);
                        medicalAssurancePolicy.CrowdType = GetCellValue(row, 28);
                        medicalAssurancePolicy.ReimbursementTreatmentThree = GetCellValue(row, 29);
                        medicalAssurancePolicy.OutpatientStartPayLineThree = GetCellValue(row, 30);
                        medicalAssurancePolicy.OutpatientEndPayLineThree = GetCellValue(row, 31);
                        medicalAssurancePolicy.PolicyLinkThree = GetCellValue(row, 32);
                        medicalAssurancePolicy.IsHot = GetCellValue(row, 33);
                        medicalAssurancePolicy.Order = GetCellValue(row, 34);

                        province = medicalAssurancePolicy.Province;
                        city = medicalAssurancePolicy.City;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("  Unexpected Exception at line {0}: {1}", rowNumber, ex.Message);
                    }
                    medicalAssurancePolicies.Add(medicalAssurancePolicy);


                }

            }


            // 直接序列化下来，不要每次处理这么多文件了
            var str = JsonConvert.SerializeObject(medicalAssurancePolicies);
            using (StreamWriter sw = new StreamWriter("all.json"))
            {
                sw.Write(str);
            }

            return medicalAssurancePolicies;
        }

        private static string GetCellPercentageValue(IRow row, int cellNumber)
        {
            if(row.Cells[cellNumber].CellType == CellType.Numeric)
            {
                return row.Cells[cellNumber].NumericCellValue * 100 + "%";
            }
            else if (row.Cells[cellNumber].CellType == CellType.String)
            {
                if (!string.IsNullOrEmpty(row.Cells[cellNumber].StringCellValue))
                {
                    return row.Cells[cellNumber].StringCellValue;
                }
                else
                {
                    return "-";
                }
            }
            return "-";
        }

        private static string GetCellValue(IRow row, int cellNumber)
        {
            try
            {
                return string.IsNullOrEmpty(row.Cells[cellNumber].StringCellValue.Trim()) ? "-" : row.Cells[cellNumber].StringCellValue;
            }
            catch (Exception ex)
            {
                return row.Cells[cellNumber].NumericCellValue.ToString();
            }
        }

        private static string GetPropertyDescription(PropertyInfo[] properties, string propertyName)
        {
            var property = properties.FirstOrDefault(a => a.Name == propertyName);
            object[] objs = property.GetCustomAttributes(typeof(DescriptionAttribute), true);
            return ((DescriptionAttribute)objs[0]).Description;
        }

        private static string CreateNewMedicalAssurancePolicyContent(MedicalAssurancePolicy policy)
        {
            var htmlRateTemplate = @"    <div style=""width: 100%;""><B>报销比率：</B>【RATE】</div>";
            var contentTitleTemplate = @"<br/>            <div style=""width: 100%; font-weight: 800;"">
                人群/险种：【PEOPLE_TYPE】
            </div>";
            var contentKVTemplate = @"            <div style=""width: 100%;"">
                <B>【KEY】：</B>【VALUE】
            </div>";

            var htmlTitle = ""; // htmlTitleTemplate.Replace("【CITY】", policy.City);
            var htmlRate = htmlRateTemplate.Replace("【RATE】", policy.ReimbursementRatio);

            var html = htmlTitle + @"<div style=""width: 100%;"">【CONTENT】</div>";
            var htmlContent = htmlRate;


            PropertyInfo[] properties = (((object)policy).GetType()).GetProperties(BindingFlags.Public | BindingFlags.Instance);

            htmlContent += contentTitleTemplate.Replace("【PEOPLE_TYPE】", policy.InsuranceType);
            var contentKV = contentKVTemplate.Replace("【KEY】", GetPropertyDescription(properties, "ReimbursementTreatment")).Replace("【VALUE】", policy.ReimbursementTreatment);

            htmlContent += contentKV;
            htmlContent += contentKVTemplate.Replace("【KEY】", GetPropertyDescription(properties, "OutpatientStartPayLine")).Replace("【VALUE】", policy.OutpatientStartPayLine);
            htmlContent += contentKVTemplate.Replace("【KEY】", GetPropertyDescription(properties, "OutpatientEndPayLine")).Replace("【VALUE】", policy.OutpatientEndPayLine);
            htmlContent += contentKVTemplate.Replace("【KEY】", GetPropertyDescription(properties, "HospitalizationStartPayLine")).Replace("【VALUE】", policy.HospitalizationStartPayLine);
            htmlContent += contentKVTemplate.Replace("【KEY】", GetPropertyDescription(properties, "HospitalizationEndPayLine")).Replace("【VALUE】", policy.HospitalizationEndPayLine);
            htmlContent += contentKVTemplate.Replace("【KEY】", GetPropertyDescription(properties, "PolicyLink")).Replace("【VALUE】", policy.PolicyLink);

            htmlContent += contentTitleTemplate.Replace("【PEOPLE_TYPE】", policy.InsuranceTypeTwo);
            htmlContent += contentKVTemplate.Replace("【KEY】", GetPropertyDescription(properties, "ReimbursementTreatmentTwo")).Replace("【VALUE】", policy.ReimbursementTreatmentTwo);
            htmlContent += contentKVTemplate.Replace("【KEY】", GetPropertyDescription(properties, "OutpatientStartPayLineTwo")).Replace("【VALUE】", policy.OutpatientStartPayLineTwo);
            htmlContent += contentKVTemplate.Replace("【KEY】", GetPropertyDescription(properties, "OutpatientEndPayLineTwo")).Replace("【VALUE】", policy.OutpatientEndPayLineTwo);
            htmlContent += contentKVTemplate.Replace("【KEY】", GetPropertyDescription(properties, "HospitalizationStartPayLineTwo")).Replace("【VALUE】", policy.HospitalizationStartPayLineTwo);
            htmlContent += contentKVTemplate.Replace("【KEY】", GetPropertyDescription(properties, "HospitalizationEndPayLineTwo")).Replace("【VALUE】", policy.HospitalizationEndPayLineTwo);
            htmlContent += contentKVTemplate.Replace("【KEY】", GetPropertyDescription(properties, "PolicyLinkTwo")).Replace("【VALUE】", policy.PolicyLinkTwo);

            htmlContent += "<br/>";
            htmlContent += contentKVTemplate.Replace("【KEY】", GetPropertyDescription(properties, "InsuranceTypeThree")).Replace("【VALUE】", policy.InsuranceTypeThree);
            htmlContent += contentKVTemplate.Replace("【KEY】", GetPropertyDescription(properties, "CrowdType")).Replace("【VALUE】", policy.CrowdType);
            htmlContent += contentKVTemplate.Replace("【KEY】", GetPropertyDescription(properties, "OutpatientStartPayLineThree")).Replace("【VALUE】", policy.OutpatientStartPayLineThree);
            htmlContent += contentKVTemplate.Replace("【KEY】", GetPropertyDescription(properties, "OutpatientEndPayLineThree")).Replace("【VALUE】", policy.OutpatientEndPayLineThree);
            htmlContent += contentKVTemplate.Replace("【KEY】", GetPropertyDescription(properties, "PolicyLinkThree")).Replace("【VALUE】", policy.PolicyLinkThree);

            html = html.Replace("【CONTENT】", htmlContent);

            return html;
        }


        private static void DeleteAll(string product, string typeCode)
        {
            while (true)
            {
                // 先获取所有的列表
                var url = $"{_host}/api/v1/wechatMedical/medicalInsurancePolicy/manage/andCount?catalog_type={product}&type_code={typeCode}&keyword=&pageNumber=1&pageSize=10&province_code=&city_code=&timestamp=1642569560925&accountid=1";
                HttpClient httpClient = new HttpClient();
                httpClient.DefaultRequestHeaders.Add("authorization", _authorization);
                var message = new HttpRequestMessage(HttpMethod.Get, url);
                message.Headers.Add("cookie", _cookies);
                message.Headers.Referrer = new Uri(_host+"/wechat/wechatManage/medicalInsurancePolicy?accountManageId=1");
                //message.Headers.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                httpClient.DefaultRequestHeaders.Add("user-agent", "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.95 Safari/537.11");

                var response = httpClient.SendAsync(message).Result;
                var result = response.Content.ReadAsStringAsync().Result;

                // 解析result里的内容，给id拿出来挨个删除，然后循环再拿，直到result里没有id为止
                var obj = JsonConvert.DeserializeObject<MedicalAssuranceHtmlResponse>(result);
                if (obj.content.res == null || obj.content.res.Count == 0)
                {
                    return;
                }
                foreach (var o in obj.content.res)
                {
                    DeleteById(o.id);
                }
            }
        }

        private static void DeleteById(int id)
        {
            var url = $"{_host}/api/v1/wechatMedical/medicalInsurancePolicy/manage/{id}?accountid=1";
            HttpClient httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Add("authorization", _authorization);
            httpClient.DefaultRequestHeaders.Add("user-agent", "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.95 Safari/537.11");
            var message = new HttpRequestMessage(HttpMethod.Delete, url);
            message.Headers.Add("cookie", _cookies);
            var response = httpClient.SendAsync(message).Result;
            var result = response.Content.ReadAsStringAsync().Result;
        }

    }
}