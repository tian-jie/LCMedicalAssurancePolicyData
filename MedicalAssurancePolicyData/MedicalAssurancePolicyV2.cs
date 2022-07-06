using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;

namespace MedicalAssurancePolicyData
{
    public class MedicalAssurancePolicyV2
    {
        public MedicalAssurancePolicyV2()
        {
        }

        public string Order { get; set; }

        public string Product { get; set; }

        public string Region { get; set; }

        public string Province { get; set; }

        public string City { get; set; }

        public bool IsMVP { get; set; }

        public string InsuranceCitizenType { get; set; }

        public string InsuranceType { get; set; }

        public string HospitalClassification { get; set; }

        public string ProductPolicyType_Standard { get; set; }

        public string ProductPolicyType_Actual { get; set; }

        public string SelfPayment { get; set; }

        public string DeductibleAmount { get; set; }

        public string ReimbersmentRate { get; set; }

        public string ReimbersmentLimit { get; set; }

        public string Memo { get; set; }

        public string HospitalizationReimbursementTreatment { get; set; }

        public string ToHtml()
        {
            var htmlRateTemplate=@"
<div style=""width: 100%;""><B>参保人群：</B>【参保人群】</div>
<div style=""width: 100%;""><B>类别：</B>【类别】</div>
<div style=""width: 100%;""><B>医院级别：</B>【医院级别】</div>
<div style=""width: 100%;""><B>产品政策类型-标准化：</B>【产品政策类型-标准化】</div>
<div style=""width: 100%;""><B>产品政策类型-实际：</B>【产品政策类型-实际】</div>
<div style=""width: 100%;""><B>首自付：</B>【首自付】</div>
<div style=""width: 100%;""><B>起付线（元 / 年）：</B>【起付线】</div>
<div style=""width: 100%;""><B>报销比率 / 额度区间：</B>【报销比率】</div>
<div style=""width: 100%;""><B>报销限额（元 / 年）：</B>【报销限额】</div>
<div style=""width: 100%;""><B>备注：</B>【备注】</div>
<div style=""width: 100%;""><B>住院报销待遇：</B>【住院报销待遇】</div>
<div></div> ";

            var htmlContent = htmlRateTemplate.Replace("【参保人群】", InsuranceCitizenType)
                .Replace("【类别】", InsuranceType)
                .Replace("【医院级别】", HospitalClassification)
                .Replace("【产品政策类型-标准化】", ProductPolicyType_Standard)
                .Replace("【产品政策类型-实际】", ProductPolicyType_Actual)
                .Replace("【首自付】", SelfPayment)
                .Replace("【起付线】", DeductibleAmount)
                .Replace("【报销比率】", ReimbersmentRate)
                .Replace("【报销限额】", ReimbersmentLimit)
                .Replace("【备注】", Memo)
                .Replace("【住院报销待遇】", HospitalizationReimbursementTreatment)
                .ReplaceLineEndings("");

            return htmlContent;
        }

    }
}
