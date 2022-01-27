//using System;
//using System.Collections.Generic;
//using System.IO;
//using System.Text;

//namespace MedicalAssurancePolicyData
//{
//    public class CsvHelper
//    {
//        string _filename;
//        List<List<string>> _csvArray;
//        List<string> _firstRow = null;
//        int _columnNumber = 0;
//        List<string> FirstRow { get { return _firstRow; } }
//        public CsvHelper(string filename)
//        {
//            _filename = filename;
//        }

//        public List<List<string>> ReadCsvToArray()
//        {
//            _csvArray = new List<List<string>>();
//            using (StreamReader sr = new StreamReader(_filename))
//            {
//                var line = sr.ReadLine();
//                _firstRow = ReadCsvLine(line);
//                _columnNumber = _firstRow.Count;
//                line = sr.ReadLine();

//                while (line != null)
//                {
//                    var row = ReadCsvLine(line);
//                    _csvArray.Add(row);
//                    line = sr.ReadLine();
//                }
//            }
//            return _csvArray;
//        }
//        private List<MedicalAssurancePolicy> ReadCsvAndParseData(string filename)
//        {
//            using (StreamReader sr = new StreamReader(filename))
//            {
//                var line = sr.ReadLine();
//                var firstRow = ReadCsvLine(sr.ReadLine());
//                line = sr.ReadLine();

//                List<MedicalAssurancePolicy> medicalAssurancePolicies = new List<MedicalAssurancePolicy>();

//                while (line != null)
//                {
//                    var row = ReadCsvLine(line);
//                    // 每一行，数据
//                    var medicalAssurancePolicy = new MedicalAssurancePolicy();
//                    string province = "";
//                    string city = "";

//                    try
//                    {
//                        medicalAssurancePolicy.TreatmentArea = row[1];
//                        medicalAssurancePolicy.ChemicalName = row[2];
//                        medicalAssurancePolicy.ProductName = row[3];
//                        medicalAssurancePolicy.DosageForm = row[4];
//                        medicalAssurancePolicy.ApprovaledText = row[5];
//                        medicalAssurancePolicy.NRDLStatus = row[6];
//                        medicalAssurancePolicy.Specification = row[7];
//                        medicalAssurancePolicy.ProductionEnterpriseName = row[8];
//                        medicalAssurancePolicy.Indications = row[9];
//                        medicalAssurancePolicy.Province = row[10];
//                        medicalAssurancePolicy.City = row[11];
//                        medicalAssurancePolicy.ReimbursementRatio = row[12];
//                        medicalAssurancePolicy.InsuranceType = row[13];
//                        medicalAssurancePolicy.ReimbursementTreatment = row[14];
//                        medicalAssurancePolicy.OutpatientStartPayLine = row[15];
//                        medicalAssurancePolicy.OutpatientEndPayLine = row[16];
//                        medicalAssurancePolicy.HospitalizationStartPayLine = row[17];
//                        medicalAssurancePolicy.HospitalizationEndPayLine = row[18];
//                        medicalAssurancePolicy.PolicyLink = row[19];
//                        medicalAssurancePolicy.InsuranceTypeTwo = row[20];
//                        medicalAssurancePolicy.ReimbursementTreatmentTwo = row[21];
//                        medicalAssurancePolicy.OutpatientStartPayLineTwo = row[22];
//                        medicalAssurancePolicy.OutpatientEndPayLineTwo = row[23];
//                        medicalAssurancePolicy.HospitalizationStartPayLineTwo = row[24];
//                        medicalAssurancePolicy.HospitalizationEndPayLineTwo = row[25];
//                        medicalAssurancePolicy.PolicyLinkTwo = row[26];
//                        medicalAssurancePolicy.InsuranceTypeThree = row[27];
//                        medicalAssurancePolicy.CrowdType = row[28];
//                        medicalAssurancePolicy.ReimbursementTreatmentThree = row[29];
//                        medicalAssurancePolicy.OutpatientStartPayLineThree = row[30];
//                        medicalAssurancePolicy.OutpatientEndPayLineThree = row[31];
//                        medicalAssurancePolicy.PolicyLinkThree = row[32];

//                        province = medicalAssurancePolicy.Province;
//                        city = medicalAssurancePolicy.City;
//                    }
//                    catch (Exception ex)
//                    {
//                        Console.WriteLine("  Unexpected Exception at line: {1}\n{0}", line, ex.Message);
//                    }
//                    medicalAssurancePolicies.Add(medicalAssurancePolicy);


//                    line = sr.ReadLine();
//                }

//                return medicalAssurancePolicies;
//            }
//        }

//        private string ReadCsvLine(StreamReader sr)
//        {
//            // 一个一个字母地读取，如果碰到引号，则两个引号之间的内容可以换行
//            StringBuilder line = new StringBuilder();
//            char c = (char)sr.Read();
//            while (true)
//            {
//                if(c == '\"')
//                {
//                    while (true)
//                    {
//                        var lastC = c;
//                        c = (char)sr.Read();
//                        // 只找下一个，非双引号连在一起的

//                    }
//                }
//            }
//        }

//        private List<string> ParseCsvLine(string line)
//        {
//            var strings = new List<string>();
//            var sb = new StringBuilder();
//            for (var i = 0; i < line.Length; i++)
//            {
//                var c = line[i];
//                if (c == '\"')
//                {
//                    i++;
//                    var last = line[i];
//                    // 向后找到下一个引号 （非双引号）
//                    while (line[i] != '\"' && last != '\"')
//                    {
//                        sb.Append(line[i]);
//                        last = line[i];
//                        i++;
//                    }
//                }
//                else if (c == ',')
//                {
//                    strings.Add(sb.ToString());
//                    sb.Clear();
//                }
//                else
//                {
//                    sb.Append(c);
//                }

//            }
//            if (sb.Length > 0)
//            {
//                strings.Add(sb.ToString());
//            }

//            // 补充，万一一行内容不全
//            if (_columnNumber > 0 && strings.Count<_columnNumber)
//            {
//                for(var i=0; i< _columnNumber-strings.Count; i++)
//                {
//                    strings.Add("");
//                }
//            }
//            return strings;
//        }

//    }
//}
