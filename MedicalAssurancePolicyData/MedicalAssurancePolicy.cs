using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace MedicalAssurancePolicyData
{
    public class MedicalAssurancePolicy
    {
        public MedicalAssurancePolicy()
        {
            MedicalAssurancePolicyDetails = new List<MedicalAssurancePolicyDetail>();
        }


        /// <summary>
        ///  1. 治疗领域
        /// </summary>
        [Description("治疗领域")]
        public string TreatmentArea { get; set; }

        /// <summary>
        ///  2. 化学名
        /// </summary>
        [Description("化学名")]
        public string ChemicalName { get; set; }

        /// <summary>
        ///  3. 商品名
        /// </summary>
        [Description("商品名")]
        public string ProductName { get; set; }

        /// <summary>
        ///  4. 剂型
        /// </summary>
        [Description("剂型")]
        public string DosageForm { get; set; }

        /// <summary>
        ///  5. 规格
        /// </summary>
        [Description("规格")]
        public string Specification { get; set; }

        /// <summary>
        ///  6. 生产注册批文
        /// </summary>
        [Description("生产注册批文")]
        public string ApprovaledText { get; set; }

        /// <summary>
        ///  7. NRDL现状
        /// </summary>
        [Description("NRDL现状")]
        public string NRDLStatus { get; set; }

        /// <summary>
        ///  8. 生产企业名称
        /// </summary>
        [Description("生产企业名称")]
        public string ProductionEnterpriseName { get; set; }

        /// <summary>
        ///  9. 适应症
        /// </summary>
        [Description("适应症")]
        public string Indications { get; set; }

        /// <summary>
        ///  10. 省
        /// </summary>
        [Description("省")]
        public string Province { get; set; }

        /// <summary>
        ///  11. 市
        /// </summary>
        [Description("市")]
        public string City { get; set; }

        /// <summary>
        ///  12. 报销比率
        /// </summary>
        [Description("报销比率")]
        public string ReimbursementRatio { get; set; }

        #region 人群/险种： 职工
        /// <summary>
        ///  13. 人群/险种： 职工
        /// </summary>
        [Description("人群/险种")]
        public string InsuranceType { get; set; }

        /// <summary>
        ///  14. 报销待遇（详细）
        /// </summary>
        [Description("报销待遇（详细）")]
        public string ReimbursementTreatment { get; set; }

        /// <summary>
        ///  15. 统筹基金起付线（门诊）
        /// </summary>
        [Description("统筹基金起付线（门诊）")]
        public string OutpatientStartPayLine { get; set; }

        /// <summary>
        ///  16. 统筹基金封顶线（门诊）
        /// </summary>
        [Description("统筹基金分封顶线（门诊）")]
        public string OutpatientEndPayLine { get; set; }

        /// <summary>
        ///  17. 起付线(住院)
        /// </summary>
        [Description("起付线(住院)")]
        public string HospitalizationStartPayLine { get; set; }

        /// <summary>
        ///  18. 封顶线(住院)
        /// </summary>
        [Description("封顶线(住院)")]
        public string HospitalizationEndPayLine { get; set; }

        /// <summary>
        ///  19. 政策链接
        /// </summary>
        [Description("政策来源")]
        public string PolicyLink { get; set; }

        #endregion

        #region 人群/险种 居民
        /// <summary>
        ///  20. 人群/险种 居民
        /// </summary>
        [Description("人群/险种")]
        public string InsuranceTypeTwo { get; set; }

        /// <summary>
        ///  21. 报销待遇（详细）
        /// </summary>
        [Description("报销待遇（详细）")]
        public string ReimbursementTreatmentTwo { get; set; }

        /// <summary>
        ///  22. 门诊起付线
        /// </summary>
        [Description("门诊起付线")]
        public string OutpatientStartPayLineTwo { get; set; }

        /// <summary>
        ///  23. 门诊封顶线
        [Description("门诊封顶线")]
        /// </summary>
        public string OutpatientEndPayLineTwo { get; set; }

        /// <summary>
        ///  24. 住院起付线
        /// </summary>
        [Description("住院起付线")]
        public string HospitalizationStartPayLineTwo { get; set; }

        /// <summary>
        ///  25. 住院封顶线
        /// </summary>
        [Description("住院封顶线")]
        public string HospitalizationEndPayLineTwo { get; set; }

        /// <summary>
        ///  26. 政策原文链接
        /// </summary>
        [Description("政策来源")]
        public string PolicyLinkTwo { get; set; }
        #endregion


        #region 险种：门特门慢门诊大病特病
        /// <summary>
        ///  27. 险种
        /// </summary>
        [Description("险种")]
        public string InsuranceTypeThree { get; set; }

        /// <summary>
        ///  28. 人群
        /// </summary>
        [Description("人群")]
        public string CrowdType { get; set; }

        /// <summary>
        ///  29. 报销待遇（详细）
        /// </summary>
        [Description("报销待遇（详细）")]
        public string ReimbursementTreatmentThree { get; set; }

        /// <summary>
        ///  30. 门诊起付线
        /// </summary>
        [Description("门诊起付线")]
        public string OutpatientStartPayLineThree { get; set; }

        /// <summary>
        ///  31. 门诊封顶线
        /// </summary>
        [Description("门诊封顶线")]
        public string OutpatientEndPayLineThree { get; set; }

        /// <summary>
        ///  32. 政策原文链接
        /// </summary>
        [Description("政策来源")]
        public string PolicyLinkThree { get; set; }

        /// <summary>
        /// 33 是否热门城市
        /// </summary>
        [Description("是否热门城市")]
        public string IsHot { get; set; }

        /// <summary>
        /// 热门城市排序
        /// </summary>
        [Description("热门城市排序")]
        public string Order { get; set; }

        #endregion


        public List<MedicalAssurancePolicyDetail> MedicalAssurancePolicyDetails { get; }
    }

    public class MedicalAssurancePolicyDetail
    {

        public MedicalAssurancePolicyDetail()
        {
            Policies = new List<KeyValuePair<string, string>>();
        }

        /// <summary>
        ///  人群
        /// </summary>
        public string CitizenType { get; set; }

        /// <summary>
        ///  险种
        /// </summary>
        public string AssuranceType { get; set; }


        /// <summary>
        ///  政策，用KV模式存储
        /// </summary>
        public List<KeyValuePair<string, string>> Policies { get; }

    }
}
