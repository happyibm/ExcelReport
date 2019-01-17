using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelReport.Model
{
    /// <summary>
    /// 出库
    /// </summary>
    public class Outstock
    {
        /// <summary>
        /// 日期
        /// </summary>
        public string RQ { get; set; }
        /// <summary>
        /// 编号
        /// </summary>
        public string BH { get; set; }

        /// <summary>
        /// 名称
        /// </summary>
        public string MC { get; set; }

        /// <summary>
        /// 型号
        /// </summary>
        public string XH { get; set; }

        /// <summary>
        /// 数量
        /// </summary>
        public string SL { get; set; }

        /// <summary>
        /// 单价
        /// </summary>
        public string DJ { get; set; }

        /// <summary>
        /// 金额
        /// </summary>
        public string JE { get; set; }

        /// <summary>
        /// 保管员
        /// </summary>
        public string BGY { get; set; }

        /// <summary>
        /// 经手人
        /// </summary>
        public string JSR { get; set; }
    }
}
