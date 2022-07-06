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

        static string _authorization = "Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ1c2VySWQiOjU3LCJ1c2VyIjp7ImlkIjo1NywiY3JlYXRlZFVzZXJJZCI6NDgsImNyZWF0ZWREYXRlIjoiMjAyMi0wMS0xOVQwMTozMzo0Ny40NjdaIiwibW9kaWZpZWRVc2VySWQiOm51bGwsIm1vZGlmaWVkRGF0ZSI6bnVsbCwiaXNEZWxldGVkIjpmYWxzZSwidmVyc2lvbk51bWJlciI6MSwidXNlcm5hbWUiOiJDMjE3MzU1IiwiZW1haWwiOiJ0aWFuX2ppZUBuZXR3b3JrLmxpbGx5LmNuIiwicmVhbG5hbWUiOiJUaWFuSmllIiwiY29tbWVudCI6bnVsbH0sInR5cGUiOiJBWlVSRSIsInN1YiI6IjNjZTZiOTE2LTc0NDctNDBkYy05YzBjLTQyMTUzZjk2YWI0ZiIsImlhdCI6MTY1NzEwNDcwNSwiZXhwIjoxNjU3MTE5MTA1fQ.VcmHfE855xMslPXjDpmn-NzNgIMgkrKg2iFXWtp5RiY";
        static string _cookies = "connect.sid=s:VRfmIyZNXZUC5McMF-pDVG8rZ1wpUcbz.jj5f+2IQeVTHFXiUhF4+Me8ksl8tdslOmTdzru7xXI8; passport-aad.1657104609071.dca97261b6b899957e219b2e61ef645648a41a5ab27d591506a27e398f2c55519d160d674a70f9e7a8cb7c7a452a3666cd645f204ffd6a13101849fd4587bb0f07f12307ad69fbcaa19337141960df34c311febc480560.420ddb7043f972231effe966527fa1f9=0";
        static string _host = "https://ipatientadmin.qa.lilly.cn";

        //static string _authorization = "Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ1c2VySWQiOjQ1LCJ1c2VyIjp7ImlkIjo0NSwiY3JlYXRlZFVzZXJJZCI6NCwiY3JlYXRlZERhdGUiOiIyMDIxLTA3LTE0VDA1OjMxOjM2Ljk5N1oiLCJtb2RpZmllZFVzZXJJZCI6bnVsbCwibW9kaWZpZWREYXRlIjpudWxsLCJpc0RlbGV0ZWQiOmZhbHNlLCJ2ZXJzaW9uTnVtYmVyIjoxLCJ1c2VybmFtZSI6IkMyMTczNTUiLCJlbWFpbCI6InRpYW5famllQG5ldHdvcmsubGlsbHkuY24iLCJyZWFsbmFtZSI6IkppZSBUaWFuIiwiY29tbWVudCI6bnVsbH0sInR5cGUiOiJBWlVSRSIsInN1YiI6ImIyOTRkOWVlLTU5ZjctNDIwZi04ZTllLWFjOTk0OTRiNDliYSIsImlhdCI6MTY1MjQzMjAwOSwiZXhwIjoxNjUyNDUzNjA5fQ.Wgz8yfTxHQEM24GGZjMr-ZKNiROp_do7VC5DMmo8FKk";
        //static string _cookies = "connect.sid=s%3ABvxk3AFkBZqBG_KBGdGk3Kpbuywr5h_t.WAwdAn41eNgGNfMilCx8%2F%2BG%2B8uRozZKWCd00EoLPANs";
        //static string _host = "https://ipatientadmin.lilly.cn";

        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            //ReadNewExcelAndParseData(args[0]);
            ReadNewExcelAndParseDataV2("1.xlsx");
        }

        private static List<Item> GetCurrentList(string product)
        {
            HttpClient httpClient = new HttpClient();
            //httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _authorization.Substring(7));
            httpClient.DefaultRequestHeaders.Add("authorization", _authorization);
            httpClient.DefaultRequestHeaders.Add("user-agent", "Mozilla/4.0 (compatible; MSIE Version; Operating System)");
            var url = $"https://ipatientadmin.qa.lilly.cn/api/v1/wechatMedical/medicalInsurancePolicy/manage/andCount?catalog_type={product}&keyword=&pageNumber=1&pageSize=100000&province_code=&city_code=&type_code=&onlyHot=0&timestamp={(DateTime.Now - new DateTime(1970, 1, 1)).Ticks / 10000}&accountid=1";
            var message = new HttpRequestMessage(HttpMethod.Get, url);
            message.Headers.Add("cookie", _cookies);
            var result = httpClient.Send(message);
            //var result = httpClient.GetAsync(url).Result;
            var html = result.Content.ReadAsStringAsync().Result;
            var maResponse = JsonConvert.DeserializeObject<MedicalAssuranceHtmlResponse>(html);
            var maData = maResponse.content.res;
            return maData;
        }

        private static void ReadNewExcelAndParseData(string folder)
        {
            // 打开某个目录
            List<MedicalAssurancePolicy> medicalAssurancePolicies;

            var filename = "all.json";
            //if (File.Exists(filename))
            //{
            //    using (StreamReader sr = new StreamReader(filename))
            //    {
            //        var str = sr.ReadToEnd();
            //        medicalAssurancePolicies = JsonConvert.DeserializeObject<List<MedicalAssurancePolicy>>(str);
            //    }
            //}
            //else
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
            PostContentByProduct(medicalAssurancePolicies, "艾乐明", "Olumiant", "", "normal");
            //Console.WriteLine("Deleting Taltz contents....");
            //DeleteAll("Taltz", "normal");
            //PostContentByProduct(medicalAssurancePolicies, "拓咨", "Taltz", "", "normal");

            //Console.WriteLine("Deleting Ve contents....");
            //DeleteAll("Verzenios", "countryHealthInsurance");
            //PostContentByProduct(medicalAssurancePolicies, "唯择", "Verzenios", "true", "countryHealthInsurance");

        }

        private static void ReadNewExcelAndParseDataV2(string folder)
        {
            // 打开某个目录
            List<MedicalAssurancePolicyV2> medicalAssurancePolicies;

            var filename = "all.json";
            //if (File.Exists(filename))
            //{
            //    using (StreamReader sr = new StreamReader(filename))
            //    {
            //        var str = sr.ReadToEnd();
            //        medicalAssurancePolicies = JsonConvert.DeserializeObject<List<MedicalAssurancePolicy>>(str);
            //    }
            //}
            //else
            {
                medicalAssurancePolicies = AnalyzeNewExcelsV2(folder);
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
            PostContentByProductV2(medicalAssurancePolicies, "艾乐明", "Olumiant", "", "normal");
            //Console.WriteLine("Deleting Taltz contents....");
            //DeleteAll("Taltz", "normal");
            //PostContentByProduct(medicalAssurancePolicies, "拓咨", "Taltz", "", "normal");

            //Console.WriteLine("Deleting Ve contents....");
            //DeleteAll("Verzenios", "countryHealthInsurance");
            //PostContentByProduct(medicalAssurancePolicies, "唯择", "Verzenios", "true", "countryHealthInsurance");

        }

        private static void PostContentByProduct(List<MedicalAssurancePolicy> medicalAssurancePolicies, string productName, string productCode, string isHot, string typeCode)
        {
            // 获取已有的数据，对比excel的数据，如果有不相同的，则post更新
            var currentMA = GetCurrentList(productCode);

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

                    HttpClient httpClient = new HttpClient();
                    html = html.Replace("\"", "'").Replace("\n", "");

                    // 确认一下html跟原始的是否相同
                    var ma = currentMA.FirstOrDefault(a => a.province_code == provinceCode.ToString() && a.city_code == cityCode.ToString());
                    if (ma == null)
                    {
                        // 不存在就添加
                        isHot = string.IsNullOrEmpty(isHot) ? (policy.IsHot == "是" ? "true" : "false") : isHot;
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
                        var url = _host + "/api/v1/wechatMedical/ueditor?accountid=1";
                        var response = httpClient.PostAsync(url, stringContent).Result;
                        var result = response.Content.ReadAsStringAsync().Result;

                        if (response.StatusCode != System.Net.HttpStatusCode.Created)
                        {
                            Console.WriteLine(" post failed: {0} - {1}", province, city);
                        }
                    }
                    else
                    {
                        // 如果html相同，则跳过，html不同的话，就修改。
                        if (ma.content == html)
                        {
                            continue;
                        }
                        else
                        {
                            isHot = string.IsNullOrEmpty(isHot) ? (policy.IsHot == "是" ? "true" : "false") : isHot;
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
                            var url = $"{_host}/api/v1/wechatMedical/ueditor/{ma.id}?accountid=1";
                            var response = httpClient.PostAsync(url, stringContent).Result;
                            var result = response.Content.ReadAsStringAsync().Result;

                            if (response.StatusCode != System.Net.HttpStatusCode.Created)
                            {
                                Console.WriteLine(" post failed: {0} - {1}", province, city);
                            }
                        }
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
        private static void PostContentByProductV2(List<MedicalAssurancePolicyV2> medicalAssurancePolicies, string productName, string productCode, string isHot, string typeCode)
        {
            // 获取已有的数据，对比excel的数据，如果有不相同的，则post更新
            var currentMA = GetCurrentList(productCode);

            foreach (var policy in medicalAssurancePolicies.Where(a => a.Product == productName).GroupBy(a => new { a.Province, a.City, a.IsMVP }))
            {
                var province = policy.Key.Province;
                var city = policy.Key.City;
                var html = "";
                var policyOrder = "-";
                foreach (var p in policy)
                {
                    html += p.ToHtml();
                    html += "<br/>";
                    policyOrder = p.Order;
                }

                if (string.IsNullOrEmpty(province) || string.IsNullOrEmpty(city))
                {
                    Console.WriteLine(" post failed: {0} - {1}, province or city cannot be empty.", province, city);
                    continue;
                }

                var provinceCode = -1;
                var cityCode = -1;
                try
                {
                    provinceCode = Provinces.FirstOrDefault(a => a.Name == province).Code;
                    cityCode = Cities.FirstOrDefault(a => a.Name == city).Code;
                }
                catch
                {
                    Console.WriteLine(" post failed: {0} - {1}, province or city not found.", province, city);
                    continue;
                }

                try
                {

                    // 做HttpClient去post
                    Console.Write(" Simulate POST: {0} - {1}, HTML ... ", province, city);

                    HttpClient httpClient = new HttpClient();
                    html = html.Replace("\"", "'");

                    // 确认一下html跟原始的是否相同
                    var ma = currentMA.FirstOrDefault(a => a.province_code == provinceCode.ToString() && a.city_code == cityCode.ToString());
                    if (ma == null)
                    {
                        Console.Write("new, inserting...");
                        // 不存在就添加
                        isHot = policy.Key.IsMVP ? "true" : "false";
                        httpClient.DefaultRequestHeaders.Add("authorization", _authorization);
                        if (policyOrder == "-")
                        {
                            policyOrder = "0";
                        }
                        var postString = $"{{\"cmd\":\"MEDICAL_INSURANCE_POLICY\",\"data\":{{\"catalog_type\":\"{productCode}\",\"province_code\":\"{provinceCode}\",\"city_code\":\"{cityCode}\",\"content\":\"{html}\",\"is_hot\":{isHot},\"type_code\":\"{typeCode}\",\"order\":{policyOrder}}}}}";
                        postString = postString.Replace("\\", "\\\\");
                        var stringContent = new StringContent(postString, Encoding.UTF8, "application/json");
                        stringContent.Headers.Add("cookie", _cookies);
                        //stringContent.Headers.Add("Content-Type", "application/json");
                        var url = _host + "/api/v1/wechatMedical/ueditor?accountid=1";
                        var response = httpClient.PostAsync(url, stringContent).Result;
                        var result = response.Content.ReadAsStringAsync().Result;

                        if (response.StatusCode != System.Net.HttpStatusCode.Created)
                        {
                            Console.WriteLine(" post failed: {0} - {1}", province, city);
                        }
                    }
                    else
                    {
                        // 如果html相同，则跳过，html不同的话，就修改。
                        if (ma.content == html)
                        {
                            Console.Write("same, skipping");
                            Console.WriteLine();
                            continue;
                        }
                        else
                        {
                            Console.Write("different, updating...");

                            isHot = policy.Key.IsMVP ? "true" : "false";
                            httpClient.DefaultRequestHeaders.Add("authorization", _authorization);
                            if (policyOrder == "-")
                            {
                                policyOrder = "0";
                            }
                            var postString = $"{{\"cmd\":\"MEDICAL_INSURANCE_POLICY\",\"data\":{{\"catalog_type\":\"{productCode}\",\"province_code\":\"{provinceCode}\",\"city_code\":\"{cityCode}\",\"content\":\"{html}\",\"is_hot\":{isHot},\"type_code\":\"{typeCode}\",\"order\":{policyOrder}}}}}";
                            postString = postString.Replace("\\", "\\\\");
                            var stringContent = new StringContent(postString, Encoding.UTF8, "application/json");
                            stringContent.Headers.Add("cookie", _cookies);
                            //stringContent.Headers.Add("Content-Type", "application/json");
                            var url = $"{_host}/api/v1/wechatMedical/ueditor/{ma.id}?accountid=1";
                            var response = httpClient.PutAsync(url, stringContent).Result;
                            var result = response.Content.ReadAsStringAsync().Result;

                            if (response.StatusCode != System.Net.HttpStatusCode.Created
                                && response.StatusCode != System.Net.HttpStatusCode.Accepted
                                && response.StatusCode != System.Net.HttpStatusCode.OK)
                            {
                                Console.WriteLine(" post failed: {0} - {1}", province, city);
                            }
                            Console.WriteLine();
                        }
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

        private static List<MedicalAssurancePolicyV2> AnalyzeNewExcelsV2(string filename)
        {
            List<MedicalAssurancePolicyV2> medicalAssurancePolicies = new List<MedicalAssurancePolicyV2>();

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
                for (var rowNumber = 3; rowNumber < 10000; rowNumber++)
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
                    var medicalAssurancePolicy = new MedicalAssurancePolicyV2();

                    try
                    {
                        medicalAssurancePolicy.Order = GetCellValue(row, 0);
                        medicalAssurancePolicy.Product = GetCellValue(row, 1);
                        medicalAssurancePolicy.Region = GetCellValue(row, 2);
                        medicalAssurancePolicy.Province = GetCellValue(row, 3);
                        medicalAssurancePolicy.City = GetCellValue(row, 4);
                        medicalAssurancePolicy.IsMVP = GetCellValue(row, 5) == "Y" ? true : false;
                        medicalAssurancePolicy.InsuranceCitizenType = GetCellValue(row, 6);
                        medicalAssurancePolicy.InsuranceType = GetCellValue(row, 7);
                        medicalAssurancePolicy.HospitalClassification = GetCellValue(row, 8);
                        medicalAssurancePolicy.ProductPolicyType_Standard = GetCellValue(row, 9);
                        medicalAssurancePolicy.ProductPolicyType_Actual = GetCellValue(row, 10);
                        medicalAssurancePolicy.SelfPayment = GetCellPercentageValue(row, 11);
                        medicalAssurancePolicy.DeductibleAmount = GetCellValue(row, 12);
                        medicalAssurancePolicy.ReimbersmentRate = GetCellPercentageValue(row, 13);
                        medicalAssurancePolicy.ReimbersmentLimit = GetCellValue(row, 14);
                        medicalAssurancePolicy.Memo = GetCellValue(row, 15);
                        medicalAssurancePolicy.HospitalizationReimbursementTreatment = GetCellValue(row, 16);

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
            if (row.Cells[cellNumber].CellType == CellType.Numeric)
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
                message.Headers.Referrer = new Uri(_host + "/wechat/wechatManage/medicalInsurancePolicy?accountManageId=1");
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
        private static string GetPropertyDescription(PropertyInfo[] properties, string propertyName)
        {
            var property = properties.FirstOrDefault(a => a.Name == propertyName);
            object[] objs = property.GetCustomAttributes(typeof(DescriptionAttribute), true);
            return ((DescriptionAttribute)objs[0]).Description;
        }
    }
}