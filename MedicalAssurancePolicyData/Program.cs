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

        static string _authorization = "Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ1c2VySWQiOjU3LCJ1c2VyIjp7ImlkIjo1NywiY3JlYXRlZFVzZXJJZCI6NDgsImNyZWF0ZWREYXRlIjoiMjAyMi0wMS0xOVQwMTozMzo0Ny40NjdaIiwibW9kaWZpZWRVc2VySWQiOm51bGwsIm1vZGlmaWVkRGF0ZSI6bnVsbCwiaXNEZWxldGVkIjpmYWxzZSwidmVyc2lvbk51bWJlciI6MSwidXNlcm5hbWUiOiJDMjE3MzU1IiwiZW1haWwiOiJ0aWFuX2ppZUBuZXR3b3JrLmxpbGx5LmNuIiwicmVhbG5hbWUiOiJUaWFuSmllIiwiY29tbWVudCI6bnVsbH0sInR5cGUiOiJBWlVSRSIsInN1YiI6IjlmNzU5Y2ZkLWUwY2ItNDJiMS1iYTgzLTA1ZTExMjhlNTI1ZCIsImlhdCI6MTY1MjA5NDE3NCwiZXhwIjoxNjUyMTA4NTc0fQ.xD67_t53X6kqyTgWDVQlhd2UTTHx-fzb95xKRHYv2uI";
        static string _cookies = "connect.sid=s%3A_TWAbUgQxxkxuix9sDBdLQ_VZsE3hprj.PtIV9R1YeXd5SrQbBpTVYls%2FP3aszRUQMZGshBV0Mjs";



        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            ReadNewExcelAndParseData(args[0]);
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
            //DeleteAll("Olumiant");
            //Console.WriteLine("Deleting Taltz contents....");
            //DeleteAll("Taltz");
            //PostContentByProduct(medicalAssurancePolicies, "艾乐明", "Olumiant");
            //PostContentByProduct(medicalAssurancePolicies, "拓咨", "Taltz");
            Console.WriteLine("Deleting Ve contents....");
            DeleteAll("Verzenios");
            PostContentByProduct(medicalAssurancePolicies, "唯择", "Verzenios");

        }

        private static void PostContentByProduct(List<MedicalAssurancePolicy> medicalAssurancePolicies, string productName, string productCode)
        {
            foreach (var policy in medicalAssurancePolicies.Where(a => a.ProductName == productName))
            {
                var html = CreateNewMedicalAssurancePolicyContent(policy);
                var province = policy.Province;
                var city = policy.City;
                if (string.IsNullOrEmpty(province) || string.IsNullOrEmpty(city))
                {
                    continue;
                }
                try
                {
                    var provinceCode = Provinces.FirstOrDefault(a => a.Name == province).Code;
                    var cityCode = Cities.FirstOrDefault(a => a.Name == city).Code;

                    // 做HttpClient去post
                    Console.WriteLine(" Simulate POST: {0} - {1}, HTML", province, city);

                    var url = "https://ipatientadmin.qa.lilly.cn/api/v1/wechatMedical/ueditor?accountid=1";
                    HttpClient httpClient = new HttpClient();
                    html = html.Replace("\"", "'").Replace("\n", "");
                    httpClient.DefaultRequestHeaders.Add("authorization", _authorization);
                    var postString = $"{{\"cmd\":\"MEDICAL_INSURANCE_POLICY\",\"data\":{{\"catalog_type\":\"{productCode}\",\"province_code\":\"{provinceCode}\",\"city_code\":\"{cityCode}\",\"content\":\"{html}\",\"is_hot\":true,\"type_code\":\"businessHealthInsurance\"}}}}";
                    var stringContent = new StringContent(postString, Encoding.UTF8, "application/json");
                    stringContent.Headers.Add("cookie", _cookies);
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
            var htmlRate = htmlRateTemplate.Replace("【RATE】", policy.ChemicalName);

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


        //private static List<MedicalAssurancePolicy> ReadCsvAndParseData(string filename)
        //{
        //    var csvHelper = new CsvHelper(filename);

        //    var csvArray = csvHelper.ReadCsvToArray();

        //    List<MedicalAssurancePolicy> medicalAssurancePolicies = new List<MedicalAssurancePolicy>();

        //    foreach(var row in csvArray)
        //    {
        //        // 每一行，数据
        //        var medicalAssurancePolicy = new MedicalAssurancePolicy();
        //        string province = "";
        //        string city = "";

        //        try
        //        {
        //            medicalAssurancePolicy.TreatmentArea = row[1];
        //            medicalAssurancePolicy.ChemicalName = row[2];
        //            medicalAssurancePolicy.ProductName = row[3];
        //            medicalAssurancePolicy.DosageForm = row[4];
        //            medicalAssurancePolicy.ApprovaledText = row[5];
        //            medicalAssurancePolicy.NRDLStatus = row[6];
        //            medicalAssurancePolicy.Specification = row[7];
        //            medicalAssurancePolicy.ProductionEnterpriseName = row[8];
        //            medicalAssurancePolicy.Indications = row[9];
        //            medicalAssurancePolicy.Province = row[10];
        //            medicalAssurancePolicy.City = row[11];
        //            medicalAssurancePolicy.ReimbursementRatio = row[12];
        //            medicalAssurancePolicy.InsuranceType = row[13];
        //            medicalAssurancePolicy.ReimbursementTreatment = row[14];
        //            medicalAssurancePolicy.OutpatientStartPayLine = row[15];
        //            medicalAssurancePolicy.OutpatientEndPayLine = row[16];
        //            medicalAssurancePolicy.HospitalizationStartPayLine = row[17];
        //            medicalAssurancePolicy.HospitalizationEndPayLine = row[18];
        //            medicalAssurancePolicy.PolicyLink = row[19];
        //            medicalAssurancePolicy.InsuranceTypeTwo = row[20];
        //            medicalAssurancePolicy.ReimbursementTreatmentTwo = row[21];
        //            medicalAssurancePolicy.OutpatientStartPayLineTwo = row[22];
        //            medicalAssurancePolicy.OutpatientEndPayLineTwo = row[23];
        //            medicalAssurancePolicy.HospitalizationStartPayLineTwo = row[24];
        //            medicalAssurancePolicy.HospitalizationEndPayLineTwo = row[25];
        //            medicalAssurancePolicy.PolicyLinkTwo = row[26];
        //            medicalAssurancePolicy.InsuranceTypeThree = row[27];
        //            medicalAssurancePolicy.CrowdType = row[28];
        //            medicalAssurancePolicy.ReimbursementTreatmentThree = row[29];
        //            medicalAssurancePolicy.OutpatientStartPayLineThree = row[30];
        //            medicalAssurancePolicy.OutpatientEndPayLineThree = row[31];
        //            medicalAssurancePolicy.PolicyLinkThree = row[32];

        //            province = medicalAssurancePolicy.Province;
        //            city = medicalAssurancePolicy.City;
        //        }
        //        catch (Exception ex)
        //        {
        //            Console.WriteLine("  Unexpected Exception at line: {1}\n{0}", row[0], ex.Message);
        //        }
        //        medicalAssurancePolicies.Add(medicalAssurancePolicy);

        //    }
        //    return medicalAssurancePolicies;
        //}

        private static void CreateContentAndPostData(List<MedicalAssurancePolicy> medicalAssurancePolicies, string productName)
        {
            List<NameCodePair> provinces = null;
            List<NameCodePair> cities = null;
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

            foreach (var medicalAssurancePolicie in medicalAssurancePolicies.Where(a => a.ProductName == productName))
            {
                PostDataToServer(medicalAssurancePolicie, provinces, cities);
            }
        }

        private static void PostDataToServer(MedicalAssurancePolicy policy, List<NameCodePair> provinces, List<NameCodePair> cities)
        {
            var html = CreateMedicalAssurancePolicyContent(policy);

            var province = policy.Province;
            var city = policy.City;
            if (string.IsNullOrEmpty(province) || string.IsNullOrEmpty(city))
            {
                return;
            }
            try
            {
                var provinceCode = provinces.FirstOrDefault(a => a.Name == province).Code;
                var cityCode = cities.FirstOrDefault(a => a.Name == city).Code;

                // 做HttpClient去post
                Console.WriteLine(" Simulate POST: {0} - {1}, HTML", province, city);

                var url = "https://ipatientadmin.qa.lilly.cn/api/v1/wechatMedical/ueditor?accountid=1";
                HttpClient httpClient = new HttpClient();
                html = html.Replace("\"", "'").Replace("\n", "");
                httpClient.DefaultRequestHeaders.Add("authorization", _authorization);
                var postString = $"{{\"cmd\":\"MEDICAL_INSURANCE_POLICY\",\"data\":{{\"catalog_type\":\"Olumiant\",\"province_code\":\"{provinceCode}\",\"city_code\":\"{cityCode}\",\"content\":\"{html}\",\"is_hot\":false}}}}";
                var stringContent = new StringContent(postString, Encoding.UTF8, "application/json");
                stringContent.Headers.Add("cookie", _cookies);
                //stringContent.Headers.Add("Content-Type", "application/json");
                var response = httpClient.PostAsync(url, stringContent).Result;
                var result = response.Content.ReadAsStringAsync().Result;

            }
            catch (Exception ex)
            {
                // do nothing
                Console.WriteLine(" Exceptions: {0} - {1}, {2}", province, city, ex.Message);
                return;
            }
        }


        private static void ReadExcelAndParseData(string folder)
        {
            // 打开某个目录

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
            DeleteAll("Olumiant");
            PostContentByProduct(medicalAssurancePolicies, "艾乐明", "Olumiant");

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
                                        break;
                                    }
                                }
                                catch (Exception)
                                {
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

            var htmlTitle = ""; // htmlTitleTemplate.Replace("【CITY】", policy.City);
            var htmlRate = htmlRateTemplate.Replace("【RATE】", policy.ChemicalName);

            var html = htmlTitle + @"<div style=""width: 100%;"">【CONTENT】</div>";
            var htmlContent = htmlRate;

            var htmlSectionTemplate = @"        <div style=""width: 100%;"">【SECTION_CONTENT】</div><br/>";
            var sectionContent = "";

            var htmlSections = "";
            foreach (var policyDetail in policy.MedicalAssurancePolicyDetails)
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

        private static void DeleteAll(string product)
        {
            while (true)
            {
                // 先获取所有的列表
                var url = $"https://ipatientadmin.qa.lilly.cn/api/v1/wechatMedical/medicalInsurancePolicy/manage/andCount?catalog_type={product}&keyword=&pageNumber=1&pageSize=10&province_code=&city_code=&timestamp=1642569560925&accountid=1";
                HttpClient httpClient = new HttpClient();
                httpClient.DefaultRequestHeaders.Add("authorization", _authorization);
                var message = new HttpRequestMessage(HttpMethod.Get, url);
                message.Headers.Add("cookie", _cookies);
                message.Headers.Referrer = new Uri("https://ipatientadmin.qa.lilly.cn/wechat/wechatManage/medicalInsurancePolicy?accountManageId=1");
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
            var url = $"https://ipatientadmin.qa.lilly.cn/api/v1/wechatMedical/medicalInsurancePolicy/manage/{id}?accountid=1";
            HttpClient httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Add("authorization", _authorization);
            httpClient.DefaultRequestHeaders.Add("user-agent", "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.95 Safari/537.11");
            var message = new HttpRequestMessage(HttpMethod.Delete, url);
            message.Headers.Add("cookie", _cookies);
            var response = httpClient.SendAsync(message).Result;
            var result = response.Content.ReadAsStringAsync().Result;
        }

        private static bool DealDynamicData(IRow row, MedicalAssurancePolicy medicalAssurancePolicy, IRow firstRow)
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