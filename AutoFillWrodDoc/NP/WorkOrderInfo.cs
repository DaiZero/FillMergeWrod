using System;

namespace AutoFillWrodDoc
{
    public class WorkOrderInfo
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
        /// 编号范围
        /// </summary>
        public string NumberRange { get; set; }

        /// <summary>
        /// 温度
        /// </summary>
        public double Temperature { get; set; }

        /// <summary>
        /// 湿度
        /// </summary>
        public double Humidity { get; set; }

        /// <summary>
        /// 检验员
        /// </summary>
        public string Tester { get; set; }

        /// <summary>
        /// 检验日期
        /// </summary>
        public DateTime TestDate { get; set; }

        /// <summary>
        /// 审核人
        /// </summary>
        public string Auditor { get; set; }

        /// <summary>
        /// 审核日期
        /// </summary>
        public DateTime AuditDate { get; set; }
    }
}
