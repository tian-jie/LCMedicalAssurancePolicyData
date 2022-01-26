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

namespace MedicalAssurancePolicyData
{
    class Program
    {
        static List<NameCodePair> Provinces;
        static List<NameCodePair> Cities;


        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            // 打开某个目录
            var folder = args[0];

            List<MedicalAssurancePolicy> medicalAssurancePolicies;

            var filename = folder + "all.json";
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
                medicalAssurancePolicies = AnalyzeExcels(folder);
            }

            using (StreamReader sr = new StreamReader(folder + "province.json"))
            {
                var str = sr.ReadToEnd();
                Provinces = JsonConvert.DeserializeObject<List<NameCodePair>>(str);
            }

            using (StreamReader sr = new StreamReader(folder + "city.json"))
            {
                var str = sr.ReadToEnd();
                Cities = JsonConvert.DeserializeObject<List<NameCodePair>>(str);
            }

            // 处理完了，然后就可以整理成需要的内容了
            // TODO: 处理html内容，然后直接post
            DeleteAll();
            foreach (var policy in medicalAssurancePolicies.Where(a=>a.ProductName=="艾乐明"))
            {
                var html = CreateMedicalAssurancePolicyContent(policy);
                var province = policy.Province;
                var city = policy.City;
                if(string.IsNullOrEmpty(province) || string.IsNullOrEmpty(city))
                {
                    continue;
                }
                try
                {
                    var provinceCode = Provinces.FirstOrDefault(a => a.Name == province).Code;
                    var cityCode = Cities.FirstOrDefault(a => a.Name == city).Code;

                    // TODO: 做HttpClient去post
                    Console.WriteLine(" Simulate POST: {0} - {1}, HTML", province, city);

                    var url = "https://ipatientadmin.qa.lilly.cn/api/v1/wechatMedical/ueditor?accountid=1";
                    HttpClient httpClient = new HttpClient();
                    html = html.Replace("\"", "'").Replace("\n", "");
                    httpClient.DefaultRequestHeaders.Add("authorization", "Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ1c2VySWQiOjU3LCJ1c2VyIjp7ImlkIjo1NywiY3JlYXRlZFVzZXJJZCI6NDgsImNyZWF0ZWREYXRlIjoiMjAyMi0wMS0xOVQwMTozMzo0Ny40NjdaIiwibW9kaWZpZWRVc2VySWQiOm51bGwsIm1vZGlmaWVkRGF0ZSI6bnVsbCwiaXNEZWxldGVkIjpmYWxzZSwidmVyc2lvbk51bWJlciI6MSwidXNlcm5hbWUiOiJDMjE3MzU1IiwiZW1haWwiOiJ0aWFuX2ppZUBuZXR3b3JrLmxpbGx5LmNuIiwicmVhbG5hbWUiOiJUaWFuSmllIiwiY29tbWVudCI6bnVsbH0sInR5cGUiOiJBWlVSRSIsInN1YiI6IjJkM2IzNDIxLWYwNjctNDAzMy1iMDg1LWUyYzAyM2I0NDdkMiIsImlhdCI6MTY0MjU1NjAzNCwiZXhwIjoxNjQyNTk5MjM0fQ.rVr9ayjexOwcc_7EEDrbiE-JH_OJfOxyGCS1Sc21pTw");
                    var postString = $"{{\"cmd\":\"MEDICAL_INSURANCE_POLICY\",\"data\":{{\"catalog_type\":\"Olumiant\",\"province_code\":\"{provinceCode}\",\"city_code\":\"{cityCode}\",\"content\":\"{html}\",\"is_hot\":false}}}}";
                    var stringContent = new StringContent(postString, Encoding.UTF8, "application/json");
                    stringContent.Headers.Add("cookie", "connect.sid=s%3AED4h7TK4VbvvvMJfE-ydh8BkNCV_WmRf.Zroy%2F%2BeUwB7oQasC3ObnYfQZKnBuHrcibM8STnn5dIs");
                    //stringContent.Headers.Add("Content-Type", "application/json");
                    var response = httpClient.PostAsync(url, stringContent).Result;
                    var result = response.Content.ReadAsStringAsync().Result;

                }
                catch (Exception ex)
                {
                    // do nothing
                    Console.WriteLine(" Exceptions: {0} - {1}, {2}", province, city, ex.Message);
                    continue;
                }
            }
        }

        private static List<MedicalAssurancePolicy> AnalyzeExcels(string folder)
        {
            // 循环目录里的文件
            var files = Directory.GetFiles(folder);

            List<MedicalAssurancePolicy> medicalAssurancePolicies = new List<MedicalAssurancePolicy>();
            foreach (var file in files)
            {
                if (!file.EndsWith(".xlsx"))
                {
                    continue;
                }
                //Console.WriteLine(file);
                // 处理sheet，每个sheet是一个城市
                // 开始处理数据，读取所有数据，最后再过滤不同品牌
                using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read))
                {
                    var excelFile = new XSSFWorkbook(fs);
                    for (var i = 0; i < 100; i++)
                    {
                        try
                        {
                            var sheet = excelFile.GetSheetAt(i);
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
                                        Console.WriteLine("  {0}, {1}, {2}", province, city, rowNumber - 1);
                                        break;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("  {0}, {1}, {2}", province, city, rowNumber - 1);
                                    break;
                                }

                                // 每一行，数据
                                var medicalAssurancePolicy = new MedicalAssurancePolicy();
                                //try
                                //{
                                //    medicalAssurancePolicy.TreatmentArea = row.Cells[1].StringCellValue;
                                //    medicalAssurancePolicy.ChemicalName = row.Cells[2].StringCellValue;
                                //    medicalAssurancePolicy.ProductName = row.Cells[3].StringCellValue;
                                //    medicalAssurancePolicy.DosageForm = row.Cells[4].StringCellValue;
                                //    medicalAssurancePolicy.Specification = row.Cells[5].StringCellValue;
                                //    medicalAssurancePolicy.Indications = row.Cells[9].StringCellValue;
                                //    medicalAssurancePolicy.Province = row.Cells[10].StringCellValue;
                                //    medicalAssurancePolicy.City = row.Cells[11].StringCellValue;
                                //    province = medicalAssurancePolicy.Province;
                                //    city = medicalAssurancePolicy.City;
                                //}
                                //catch (Exception ex)
                                //{
                                //    Console.WriteLine("  Unexpected Exception at line {0}: {1}", rowNumber, ex.Message);
                                //}
                                //medicalAssurancePolicies.Add(medicalAssurancePolicy);

                                //// 以上是标准数据
                                //// 以下是处理各种动态数据咯
                                //if (DealDynamicData(row, medicalAssurancePolicy, firstRow))
                                //{
                                //    continue;
                                //}

                                try
                                {
                                    medicalAssurancePolicy.TreatmentArea = row.Cells[1].StringCellValue;
                                    medicalAssurancePolicy.ChemicalName = row.Cells[2].StringCellValue;
                                    medicalAssurancePolicy.ProductName = row.Cells[3].StringCellValue;
                                    medicalAssurancePolicy.DosageForm = row.Cells[4].StringCellValue;
                                    medicalAssurancePolicy.ApprovaledText = row.Cells[5].StringCellValue;
                                    medicalAssurancePolicy.NRDLStatus = row.Cells[6].StringCellValue;
                                    medicalAssurancePolicy.Specification = row.Cells[7].StringCellValue;
                                    medicalAssurancePolicy.ProductionEnterpriseName = row.Cells[8].StringCellValue;
                                    medicalAssurancePolicy.Indications = row.Cells[9].StringCellValue;
                                    medicalAssurancePolicy.Province = row.Cells[10].StringCellValue;
                                    medicalAssurancePolicy.City = row.Cells[11].StringCellValue;
                                    medicalAssurancePolicy.ReimbursementRatio = row.Cells[12].StringCellValue;
                                    medicalAssurancePolicy.InsuranceType = row.Cells[13].StringCellValue;
                                    medicalAssurancePolicy.ReimbursementTreatment = row.Cells[14].StringCellValue;
                                    medicalAssurancePolicy.OutpatientStartPayLine = row.Cells[15].StringCellValue;
                                    medicalAssurancePolicy.OutpatientEndPayLine = row.Cells[16].StringCellValue;
                                    medicalAssurancePolicy.HospitalizationStartPayLine = row.Cells[17].StringCellValue;
                                    medicalAssurancePolicy.HospitalizationEndPayLine = row.Cells[18].StringCellValue;
                                    medicalAssurancePolicy.PolicyLink = row.Cells[19].StringCellValue;
                                    medicalAssurancePolicy.InsuranceTypeTwo = row.Cells[20].StringCellValue;
                                    medicalAssurancePolicy.ReimbursementTreatmentTwo = row.Cells[21].StringCellValue;
                                    medicalAssurancePolicy.OutpatientStartPayLineTwo = row.Cells[22].StringCellValue;
                                    medicalAssurancePolicy.OutpatientEndPayLineTwo = row.Cells[23].StringCellValue;
                                    medicalAssurancePolicy.HospitalizationStartPayLineTwo = row.Cells[24].StringCellValue;
                                    medicalAssurancePolicy.HospitalizationEndPayLineTwo = row.Cells[25].StringCellValue;
                                    medicalAssurancePolicy.PolicyLinkTwo = row.Cells[26].StringCellValue;
                                    medicalAssurancePolicy.InsuranceTypeThree = row.Cells[27].StringCellValue;
                                    medicalAssurancePolicy.CrowdType = row.Cells[28].StringCellValue;
                                    medicalAssurancePolicy.ReimbursementTreatmentThree = row.Cells[29].StringCellValue;
                                    medicalAssurancePolicy.OutpatientStartPayLineThree = row.Cells[30].StringCellValue;
                                    medicalAssurancePolicy.OutpatientEndPayLineThree = row.Cells[31].StringCellValue;
                                    medicalAssurancePolicy.PolicyLinkThree = row.Cells[32].StringCellValue;

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
                        catch (Exception ex)
                        {
                            // do nothing, 没了
                            //Console.WriteLine("  Exceptions: {0}", ex.Message);
                            break;
                        }
                    }
                }

            }

            // 直接序列化下来，不要每次处理这么多文件了
            var str = JsonConvert.SerializeObject(medicalAssurancePolicies);
            using (StreamWriter sw = new StreamWriter(folder + "\\all.json"))
            {
                sw.Write(str);
            }

            return medicalAssurancePolicies;
        }

        private static string CreateMedicalAssurancePolicyContent(MedicalAssurancePolicy policy)
        {
            var htmlTitleTemplate = @"<div style=""width: 100%; text-align: center;""><h1>【CITY】市医保政策</h1></div>";
            var htmlRateTemplate = @"    <div style=""width: 100%;""><B>报销比率：</B>【RATE】</div><br/>";
            var contentTitleTemplate = @"            <div style=""width: 100%; font-weight: 800;"">
                人群/险种：【PEOPLE_TYPE】
            </div>";
            var contentKVTemplate = @"            <div style=""width: 100%;"">
                <B>【KEY】：</B>【VALUE】
            </div>";

            var htmlTitle = htmlTitleTemplate.Replace("【CITY】", policy.City);
            var htmlRate = htmlRateTemplate.Replace("【RATE】", policy.ChemicalName);

            var html = htmlTitle + @"<div style=""width: 100%;"">【CONTENT】</div>";
            var htmlContent = htmlRate;

            var htmlSectionTemplate = @"        <div style=""width: 100%;"">【SECTION_CONTENT】</div><br/>";
            var sectionContent = "";

            var htmlSections = "";
            foreach(var policyDetail in policy.MedicalAssurancePolicyDetails)
            {
                var contentTitle = contentTitleTemplate.Replace("【PEOPLE_TYPE】", policyDetail.CitizenType + policyDetail.AssuranceType);
                sectionContent = contentTitle;
                foreach (var p in policyDetail.Policies)
                {
                    var key = p.Key;
                    var value = p.Value;
                    var htmlKV = contentKVTemplate.Replace("【KEY】", key).Replace("【VALUE】", value);
                    sectionContent += htmlKV;
                }
                var htmlSection = htmlSectionTemplate.Replace("【SECTION_CONTENT】", sectionContent);
                htmlSections += htmlSection;
            }
            html = html.Replace("【CONTENT】", htmlSections);

            return html;
        }

        private static void DeleteAll()
        {
            while (true)
            {
                // 先获取所有的列表
                var url = "https://ipatientadmin.qa.lilly.cn/api/v1/wechatMedical/medicalInsurancePolicy/manage/andCount?catalog_type=Olumiant&keyword=&pageNumber=1&pageSize=10&province_code=&city_code=&timestamp=1642569560925&accountid=1";
                HttpClient httpClient = new HttpClient();
                httpClient.DefaultRequestHeaders.Add("authorization", "Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ1c2VySWQiOjU3LCJ1c2VyIjp7ImlkIjo1NywiY3JlYXRlZFVzZXJJZCI6NDgsImNyZWF0ZWREYXRlIjoiMjAyMi0wMS0xOVQwMTozMzo0Ny40NjdaIiwibW9kaWZpZWRVc2VySWQiOm51bGwsIm1vZGlmaWVkRGF0ZSI6bnVsbCwiaXNEZWxldGVkIjpmYWxzZSwidmVyc2lvbk51bWJlciI6MSwidXNlcm5hbWUiOiJDMjE3MzU1IiwiZW1haWwiOiJ0aWFuX2ppZUBuZXR3b3JrLmxpbGx5LmNuIiwicmVhbG5hbWUiOiJUaWFuSmllIiwiY29tbWVudCI6bnVsbH0sInR5cGUiOiJBWlVSRSIsInN1YiI6IjJkM2IzNDIxLWYwNjctNDAzMy1iMDg1LWUyYzAyM2I0NDdkMiIsImlhdCI6MTY0MjU1NjAzNCwiZXhwIjoxNjQyNTk5MjM0fQ.rVr9ayjexOwcc_7EEDrbiE-JH_OJfOxyGCS1Sc21pTw");
                var message = new HttpRequestMessage(HttpMethod.Get, url);
                message.Headers.Add("cookie", "connect.sid=s%3AED4h7TK4VbvvvMJfE-ydh8BkNCV_WmRf.Zroy%2F%2BeUwB7oQasC3ObnYfQZKnBuHrcibM8STnn5dIs");
                message.Headers.Referrer = new Uri("https://ipatientadmin.qa.lilly.cn/wechat/wechatManage/medicalInsurancePolicy?accountManageId=1");
                //message.Headers.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                httpClient.DefaultRequestHeaders.Add("user-agent", "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.95 Safari/537.11");

                var response = httpClient.SendAsync(message).Result;
                var result = response.Content.ReadAsStringAsync().Result;

                // 解析result里的内容，给id拿出来挨个删除，然后循环再拿，直到result里没有id为止
                var obj = JsonConvert.DeserializeObject<MedicalAssuranceHtmlResponse>(result);
                if (obj.content.res.Count == 0)
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
            var url = $"https://ipatientadmin.qa.lilly.cn/api/v1/wechatMedical/medicalInsurancePolicy/manage/{id}?accountid=1";
            HttpClient httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Add("authorization", "Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ1c2VySWQiOjU3LCJ1c2VyIjp7ImlkIjo1NywiY3JlYXRlZFVzZXJJZCI6NDgsImNyZWF0ZWREYXRlIjoiMjAyMi0wMS0xOVQwMTozMzo0Ny40NjdaIiwibW9kaWZpZWRVc2VySWQiOm51bGwsIm1vZGlmaWVkRGF0ZSI6bnVsbCwiaXNEZWxldGVkIjpmYWxzZSwidmVyc2lvbk51bWJlciI6MSwidXNlcm5hbWUiOiJDMjE3MzU1IiwiZW1haWwiOiJ0aWFuX2ppZUBuZXR3b3JrLmxpbGx5LmNuIiwicmVhbG5hbWUiOiJUaWFuSmllIiwiY29tbWVudCI6bnVsbH0sInR5cGUiOiJBWlVSRSIsInN1YiI6IjJkM2IzNDIxLWYwNjctNDAzMy1iMDg1LWUyYzAyM2I0NDdkMiIsImlhdCI6MTY0MjU1NjAzNCwiZXhwIjoxNjQyNTk5MjM0fQ.rVr9ayjexOwcc_7EEDrbiE-JH_OJfOxyGCS1Sc21pTw");
            httpClient.DefaultRequestHeaders.Add("user-agent", "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.95 Safari/537.11");
            var message = new HttpRequestMessage(HttpMethod.Delete, url);
            message.Headers.Add("Cookie", "connect.sid=s%3AED4h7TK4VbvvvMJfE-ydh8BkNCV_WmRf.Zroy%2F%2BeUwB7oQasC3ObnYfQZKnBuHrcibM8STnn5dIs");
            var response = httpClient.SendAsync(message).Result;
            var result = response.Content.ReadAsStringAsync().Result;
        }

        private static bool DealDynamicData(IRow row, MedicalAssurancePolicy medicalAssurancePolicy, IRow firstRow )
        {
            // 第一组数据
            try
            {
                if (row.Cells.Count <= 13)
                {
                    return true;
                }
                var medicalAssurancePolicyDetail = new MedicalAssurancePolicyDetail();
                medicalAssurancePolicyDetail.CitizenType = row.Cells[13].StringCellValue;
                medicalAssurancePolicyDetail.AssuranceType = "";
                for (var x = 14; x <= 19; x++)
                {
                    var value = "";
                    try
                    {
                        value = row.Cells[x].StringCellValue;
                    }
                    catch
                    {
                        try
                        {
                            value = row.Cells[x].NumericCellValue.ToString();
                        }
                        catch (Exception ex1)
                        {
                            value = "-";
                        }
                    }

                    medicalAssurancePolicyDetail.Policies.Add(
                        new System.Collections.Generic.KeyValuePair<string, string>(
                            firstRow.Cells[x].StringCellValue,
                            string.IsNullOrEmpty(value) ? "-" : value
                    ));
                }
                medicalAssurancePolicy.MedicalAssurancePolicyDetails.Add(medicalAssurancePolicyDetail);
            }
            catch (Exception ex)
            {
                return true;
            }
            // 第二组数据
            try
            {
                if (row.Cells.Count <= 20)
                {
                    return true;
                }
                var medicalAssurancePolicyDetail = new MedicalAssurancePolicyDetail();
                medicalAssurancePolicyDetail.CitizenType = row.Cells[20].StringCellValue;
                medicalAssurancePolicyDetail.AssuranceType = "";
                for (var x = 21; x <= 26; x++)
                {
                    var value = "";
                    try
                    {
                        value = row.Cells[x].StringCellValue;
                    }
                    catch
                    {
                        try
                        {
                            value = row.Cells[x].NumericCellValue.ToString();
                        }
                        catch (Exception ex1)
                        {
                            value = "-";
                        }
                    }

                    medicalAssurancePolicyDetail.Policies.Add(
                        new System.Collections.Generic.KeyValuePair<string, string>(
                            firstRow.Cells[x].StringCellValue,
                            string.IsNullOrEmpty(value) ? "-" : value
                    ));
                }
                medicalAssurancePolicy.MedicalAssurancePolicyDetails.Add(medicalAssurancePolicyDetail);
            }
            catch (Exception ex)
            {
                return true;
            }
            // 第三组数据
            try
            {
                if (row.Cells.Count <= 27)
                {
                    return true;
                }
                var medicalAssurancePolicyDetail = new MedicalAssurancePolicyDetail();
                medicalAssurancePolicyDetail.CitizenType = row.Cells[28].StringCellValue;
                medicalAssurancePolicyDetail.AssuranceType = row.Cells[27].StringCellValue; ;
                for (var x = 29; x <= 32; x++)
                {
                    var value = "";
                    try
                    {
                        value = row.Cells[x].StringCellValue;
                    }
                    catch (Exception ex)
                    {
                        try
                        {
                            value = row.Cells[x].NumericCellValue.ToString();
                        }
                        catch (Exception ex1)
                        {
                            value = "-";
                        }
                    }

                    medicalAssurancePolicyDetail.Policies.Add(
                        new System.Collections.Generic.KeyValuePair<string, string>(
                            firstRow.Cells[x].StringCellValue,
                            string.IsNullOrEmpty(value) ? "-" : value
                    ));
                }
                medicalAssurancePolicy.MedicalAssurancePolicyDetails.Add(medicalAssurancePolicyDetail);
            }
            catch (Exception ex)
            {
                return true;
            }
            return false;
        }
    }
}
