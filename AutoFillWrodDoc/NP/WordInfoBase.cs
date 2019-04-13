namespace AutoFillWrodDoc
{
    public class WordInfoBase
    {
        /// <summary>
        /// 规格型号
        /// </summary>
        public string Spec { get; set; }

        /// <summary>
        /// 订单号
        /// </summary>
        public string WorkOrderNo { get; set; }

        /// <summary>
        /// 编号
        /// </summary>
        public string Number { get; set; }

        /// <summary>
        /// 编号范围
        /// </summary>
        public string NumberRange { get; set; }

        /// <summary>
        /// 数量
        /// </summary>
        public string Qty { get; set; }

        /// <summary>
        /// 温度 ℃
        /// </summary>
        public string Temperature { get; set; }

        /// <summary>
        /// 湿度 %RH
        /// </summary>
        public string Humidity { get; set; }

        /// <summary>
        /// 检验员
        /// </summary>
        public string Tester { get; set; }

        /// <summary>
        /// 检验日期
        /// </summary>
        public string TestDate { get; set; }

        /// <summary>
        /// 审核人
        /// </summary>
        public string Auditor { get; set; }

        /// <summary>
        /// 审核日期
        /// </summary>
        public string AuditDate { get; set; }
    }
}
