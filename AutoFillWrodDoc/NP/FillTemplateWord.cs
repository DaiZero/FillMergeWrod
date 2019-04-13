using Aspose.Words;
using Aspose.Words.Replacing;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AutoFillWrodDoc
{
    public enum WordDataTemplate
    {
        Base = 0,
        SingleChannel = 1,
        Double = 2
    }

    public class FillTemplateWord
    {
        public static readonly Dictionary<string, WordDataTemplate> SpecDataDic = new Dictionary<string, WordDataTemplate>()
        {
            { "NPEXA-CM31", WordDataTemplate.SingleChannel },
            { "NPEXA-GM31", WordDataTemplate.SingleChannel },
            { "NPEXA-CM311", WordDataTemplate.Double },
            { "NPEXA-CM3D11", WordDataTemplate.Double },
            { "NPEXA-C5D111", WordDataTemplate.Base },
            { "NPEXA-C5111", WordDataTemplate.Base },
            { "NPEXA-C511", WordDataTemplate.Base },
        };

        public ResultInfo GetWordInfos(WorkOrderInfo orderInfo)
        {
            ResultInfo resultInfo = new ResultInfo();
            if (SpecDataDic.Keys.Contains(orderInfo.Spec))
            {
                resultInfo.WordDataTemplate = SpecDataDic[orderInfo.Spec];
            }

            List<ChannelWordInfo> channelWords = new List<ChannelWordInfo>();
            var norg = orderInfo.NumberRange.Trim().Replace("～", "~").Replace("，", ",").TrimEnd(',');
            var noarr = norg.Split(',');

            int sumCount = 0;
            foreach (var item in noarr)
            {
                if (item.Contains("~"))
                {
                    var nos = item.Split('~');
                    if (nos != null && nos.Length == 2)
                    {
                        var r1 = decimal.TryParse(nos[0], out decimal s1);
                        var r2 = decimal.TryParse(nos[1], out decimal s2);
                        if (r1 && r2 && s2 > s1)
                        {
                            var icount = Convert.ToInt32(s2 - s1) + 1;
                            sumCount = sumCount + icount;
                            for (int i = 0; i < icount; i++)
                            {
                                var cnWord = new ChannelWordInfo()
                                {
                                    WorkOrderNo = orderInfo.WorkOrderNo,
                                    Number = s1.ToString(),
                                    NumberRange = orderInfo.NumberRange,
                                    Temperature = orderInfo.Temperature.ToString().TrimEnd('0') + "℃",
                                    Humidity = orderInfo.Humidity.ToString().TrimEnd('0') + "%RH",
                                    Tester = orderInfo.Tester,
                                    TestDate = orderInfo.TestDate.ToString("yyyy.MM.dd"),
                                    Auditor = orderInfo.Auditor,
                                    AuditDate = orderInfo.AuditDate.ToString("yyyy.MM.dd"),
                                    Spec = orderInfo.Spec
                                };
                                SetChannelData(cnWord, resultInfo.WordDataTemplate);
                                channelWords.Add(cnWord);
                                s1++;
                            }
                        }
                        else
                        {
                            resultInfo.Message = "编号范围不正常";
                            return resultInfo;
                        }
                    }
                    else
                    {
                        resultInfo.Message = "编号范围不正常";
                        return resultInfo;
                    }
                }
                else
                {
                    sumCount++;
                    var cnWord = new ChannelWordInfo()
                    {
                        WorkOrderNo = orderInfo.WorkOrderNo,
                        Number = item,
                        NumberRange = orderInfo.NumberRange,
                        Temperature = orderInfo.Temperature.ToString().TrimEnd('0') + "℃",
                        Humidity = orderInfo.Humidity.ToString().TrimEnd('0') + "%RH",
                        Tester = orderInfo.Tester,
                        TestDate = orderInfo.TestDate.ToString("YYYY.MM.dd"),
                        Auditor = orderInfo.Auditor,
                        AuditDate = orderInfo.AuditDate.ToString("YYYY.MM.dd"),
                        Spec = orderInfo.Spec
                    };
                    SetChannelData(cnWord, resultInfo.WordDataTemplate);
                    channelWords.Add(cnWord);
                }
            }

            foreach (var item in channelWords)
            {
                item.Qty = sumCount.ToString();
            }
            resultInfo.ChannelWordInfos = channelWords;
            resultInfo.Succeed = true;
            return resultInfo;
        }

        private void SetChannelData(ChannelWordInfo channelWordInfo, WordDataTemplate wordDataTemplate)
        {
            switch (wordDataTemplate)
            {
                case WordDataTemplate.Base:
                    break;
                case WordDataTemplate.SingleChannel:
                    SetC1ChannelData(channelWordInfo);
                    break;
                case WordDataTemplate.Double:
                    SetDoubleChannelData(channelWordInfo);
                    break;
                default:
                    break;
            }
        }

        static int GetRandomSeed()
        {
            byte[] bytes = new byte[4];
            System.Security.Cryptography.RNGCryptoServiceProvider rng = new System.Security.Cryptography.RNGCryptoServiceProvider();
            rng.GetBytes(bytes);
            return BitConverter.ToInt32(bytes, 0);
        }


        private void SetC1ChannelData(ChannelWordInfo channelWordInfo)
        {
            Random random = new Random(GetRandomSeed());

            #region C1

            var C1_U_1_0 = (4 + 0.006 * (random.NextDouble() * 2 - 1));
            var C1_U_1_25 = (8 + 0.006 * (random.NextDouble() * 2 - 1));
            var C1_U_1_50 = (12 + 0.006 * (random.NextDouble() * 2 - 1));
            var C1_U_1_75 = (16 + 0.006 * (random.NextDouble() * 2 - 1));
            var C1_U_1_100 = (20 + 0.006 * (random.NextDouble() * 2 - 1));

            var C1_U_2_0 = (4 + 0.006 * (random.NextDouble() * 2 - 1));
            var C1_U_2_25 = (8 + 0.006 * (random.NextDouble() * 2 - 1));
            var C1_U_2_50 = (12 + 0.006 * (random.NextDouble() * 2 - 1));
            var C1_U_2_75 = (16 + 0.006 * (random.NextDouble() * 2 - 1));
            var C1_U_2_100 = (20 + 0.006 * (random.NextDouble() * 2 - 1));

            var C1_D_1_0 = (4 + 0.006 * (random.NextDouble() * 2 - 1));
            var C1_D_1_25 = (8 + 0.006 * (random.NextDouble() * 2 - 1));
            var C1_D_1_50 = (12 + 0.006 * (random.NextDouble() * 2 - 1));
            var C1_D_1_75 = (16 + 0.006 * (random.NextDouble() * 2 - 1));
            var C1_D_1_100 = (20 + 0.006 * (random.NextDouble() * 2 - 1));

            var C1_D_2_0 = (4 + 0.006 * (random.NextDouble() * 2 - 1));
            var C1_D_2_25 = (8 + 0.006 * (random.NextDouble() * 2 - 1));
            var C1_D_2_50 = (12 + 0.006 * (random.NextDouble() * 2 - 1));
            var C1_D_2_75 = (16 + 0.006 * (random.NextDouble() * 2 - 1));
            var C1_D_2_100 = (20 + 0.006 * (random.NextDouble() * 2 - 1));

            List<double> listu1 = new List<double>
            {
                (C1_U_1_0 - 4) / 16 * 100,
                (C1_U_1_25 - 8) / 16 * 100,
                (C1_U_1_50 - 12) / 16 * 100,
                (C1_U_1_75 - 16) / 16 * 100,
                (C1_U_1_100 - 20) / 16 * 100,

                (C1_U_2_0 - 4) / 16 * 100,
                (C1_U_2_25 -8) / 16 * 100,
                (C1_U_2_50 - 12) / 16 * 100,
                (C1_U_2_75 - 16) / 16 * 100,
                (C1_U_2_100 - 20) / 16 * 100
            };

            List<double> listd1 = new List<double>
            {
                (C1_D_1_0 - 4) / 16 * 100,
                (C1_D_1_25 - 8) / 16 * 100,
                (C1_D_1_50 - 12) / 16 * 100,
                (C1_D_1_75 - 16) / 16 * 100,
                (C1_D_1_100 - 20) / 16 * 100,

                (C1_D_2_0 - 4) / 16 * 100,
                (C1_D_2_25 - 8) / 16 * 100,
                (C1_D_2_50 - 12) / 16 * 100,
                (C1_D_2_75 - 16) / 16 * 100,
                (C1_D_2_100 -20) / 16 * 100
            };

            List<double> listud1 = listu1.ToList();
            listud1.AddRange(listd1);
            var min = listud1.Min();
            var max = listud1.Max();
            if (Math.Abs(min) > Math.Abs(max))
            {
                channelWordInfo.C1_Bjqd = min.ToString("0.00");
            }
            else
            {
                channelWordInfo.C1_Bjqd = max.ToString("0.00");
            }

            List<double> hcList = new List<double>();
            for (int i = 0; i < 10; i++)
            {
                hcList.Add(Math.Abs(listd1[i] - listu1[i]));
            }
            channelWordInfo.C1_Hc = hcList.Max().ToString("0.00");

            List<double> cfList = new List<double>
            {
                Math.Abs(C1_U_1_0 - C1_U_2_0),
                Math.Abs(C1_U_1_50 - C1_U_2_50),
                Math.Abs(C1_U_1_25 - C1_U_2_25),
                Math.Abs(C1_U_1_75 - C1_U_2_75),
                Math.Abs(C1_U_1_100 - C1_U_2_100),

                Math.Abs(C1_D_1_0 - C1_D_2_0),
                Math.Abs(C1_D_1_50 - C1_D_2_50),
                Math.Abs(C1_D_1_25 - C1_D_2_25),
                Math.Abs(C1_D_1_75 - C1_D_2_75),
                Math.Abs(C1_D_1_100 - C1_D_2_100)
            };

            channelWordInfo.C1_Cfxwc = cfList.Max().ToString("0.00");


            #region 赋值
            channelWordInfo.C1_U_1_0 = C1_U_1_0.ToString("0.000");
            channelWordInfo.C1_U_1_25 = C1_U_1_25.ToString("0.000");
            channelWordInfo.C1_U_1_50 = C1_U_1_50.ToString("0.000");
            channelWordInfo.C1_U_1_75 = C1_U_1_75.ToString("0.000");
            channelWordInfo.C1_U_1_100 = C1_U_1_100.ToString("0.000");

            channelWordInfo.C1_U_2_0 = C1_U_2_0.ToString("0.000");
            channelWordInfo.C1_U_2_25 = C1_U_2_25.ToString("0.000");
            channelWordInfo.C1_U_2_50 = C1_U_2_50.ToString("0.000");
            channelWordInfo.C1_U_2_75 = C1_U_2_75.ToString("0.000");
            channelWordInfo.C1_U_2_100 = C1_U_2_100.ToString("0.000");

            channelWordInfo.C1_D_1_0 = C1_D_1_0.ToString("0.000");
            channelWordInfo.C1_D_1_25 = C1_D_1_25.ToString("0.000");
            channelWordInfo.C1_D_1_50 = C1_D_1_50.ToString("0.000");
            channelWordInfo.C1_D_1_75 = C1_D_1_75.ToString("0.000");
            channelWordInfo.C1_D_1_100 = C1_D_1_100.ToString("0.000");

            channelWordInfo.C1_D_2_0 = C1_D_2_0.ToString("0.000");
            channelWordInfo.C1_D_2_25 = C1_D_2_25.ToString("0.000");
            channelWordInfo.C1_D_2_50 = C1_D_2_50.ToString("0.000");
            channelWordInfo.C1_D_2_75 = C1_D_2_75.ToString("0.000");
            channelWordInfo.C1_D_2_100 = C1_D_2_100.ToString("0.000");




            channelWordInfo.C1_SQ = "0.00";
            #endregion
            #endregion
        }

        private void SetC2ChannelData(ChannelWordInfo channelWordInfo)
        {
            Random random = new Random(GetRandomSeed());
            #region C2

            var C2_U_1_0 = (4 + 0.006 * (random.NextDouble() * 2 - 1));
            var C2_U_1_25 = (8 + 0.006 * (random.NextDouble() * 2 - 1));
            var C2_U_1_50 = (12 + 0.006 * (random.NextDouble() * 2 - 1));
            var C2_U_1_75 = (16 + 0.006 * (random.NextDouble() * 2 - 1));
            var C2_U_1_100 = (20 + 0.006 * (random.NextDouble() * 2 - 1));

            var C2_U_2_0 = (4 + 0.006 * (random.NextDouble() * 2 - 1));
            var C2_U_2_25 = (8 + 0.006 * (random.NextDouble() * 2 - 1));
            var C2_U_2_50 = (12 + 0.006 * (random.NextDouble() * 2 - 1));
            var C2_U_2_75 = (16 + 0.006 * (random.NextDouble() * 2 - 1));
            var C2_U_2_100 = (20 + 0.006 * (random.NextDouble() * 2 - 1));



            var C2_D_1_0 = (4 + 0.006 * (random.NextDouble() * 2 - 1));
            var C2_D_1_25 = (8 + 0.006 * (random.NextDouble() * 2 - 1));
            var C2_D_1_50 = (12 + 0.006 * (random.NextDouble() * 2 - 1));
            var C2_D_1_75 = (16 + 0.006 * (random.NextDouble() * 2 - 1));
            var C2_D_1_100 = (20 + 0.006 * (random.NextDouble() * 2 - 1));

            var C2_D_2_0 = (4 + 0.006 * (random.NextDouble() * 2 - 1));
            var C2_D_2_25 = (8 + 0.006 * (random.NextDouble() * 2 - 1));
            var C2_D_2_50 = (12 + 0.006 * (random.NextDouble() * 2 - 1));
            var C2_D_2_75 = (16 + 0.006 * (random.NextDouble() * 2 - 1));
            var C2_D_2_100 = (20 + 0.006 * (random.NextDouble() * 2 - 1));

            List<double> listu2 = new List<double>
            {
                (C2_U_1_0 - 4) / 16 * 100,
                (C2_U_1_25 - 8) / 16 * 100,
                (C2_U_1_50 - 12) / 16 * 100,
                (C2_U_1_75 - 16) / 16 * 100,
                (C2_U_1_100 - 20) / 16 * 100,

                (C2_U_2_0 - 4) / 16 * 100,
                (C2_U_2_25 -8) / 16 * 100,
                (C2_U_2_50 - 12) / 16 * 100,
                (C2_U_2_75 - 16) / 16 * 100,
                (C2_U_2_100 - 20) / 16 * 100
            };

            List<double> listd2 = new List<double>
            {
                (C2_D_1_0 - 4) / 16 * 100,
                (C2_D_1_25 - 8) / 16 * 100,
                (C2_D_1_50 - 12) / 16 * 100,
                (C2_D_1_75 - 16) / 16 * 100,
                (C2_D_1_100 - 20) / 16 * 100,

                (C2_D_2_0 - 4) / 16 * 100,
                (C2_D_2_25 - 8) / 16 * 100,
                (C2_D_2_50 - 12) / 16 * 100,
                (C2_D_2_75 - 16) / 16 * 100,
                (C2_D_2_100 -20) / 16 * 100
            };

            List<double> listud2 = listu2.ToList();
            listud2.AddRange(listd2);
            var min2 = listud2.Min();
            var max2 = listud2.Max();
            if (Math.Abs(min2) > Math.Abs(max2))
            {
                channelWordInfo.C2_Bjqd = min2.ToString("0.00");
            }
            else
            {
                channelWordInfo.C2_Bjqd = max2.ToString("0.00");
            }

            List<double> hcList2 = new List<double>();
            for (int i = 0; i < 10; i++)
            {
                hcList2.Add(Math.Abs(listd2[i] - listu2[i]));
            }
            channelWordInfo.C2_Hc = hcList2.Max().ToString("0.00");

            List<double> cfList2 = new List<double>
            {
                Math.Abs(C2_U_1_0 - C2_U_2_0),
                Math.Abs(C2_U_1_50 - C2_U_2_50),
                Math.Abs(C2_U_1_25 - C2_U_2_25),
                Math.Abs(C2_U_1_75 - C2_U_2_75),
                Math.Abs(C2_U_1_100 - C2_U_2_100),

                Math.Abs(C2_D_1_0 - C2_D_2_0),
                Math.Abs(C2_D_1_50 - C2_D_2_50),
                Math.Abs(C2_D_1_25 - C2_D_2_25),
                Math.Abs(C2_D_1_75 - C2_D_2_75),
                Math.Abs(C2_D_1_100 - C2_D_2_100)
            };

            channelWordInfo.C2_Cfxwc = cfList2.Max().ToString("0.00");
            channelWordInfo.C2_SQ = "0.00";

            #region 赋值
            channelWordInfo.C2_U_1_0 = C2_U_1_0.ToString("0.000");
            channelWordInfo.C2_U_1_25 = C2_U_1_25.ToString("0.000");
            channelWordInfo.C2_U_1_50 = C2_U_1_50.ToString("0.000");
            channelWordInfo.C2_U_1_75 = C2_U_1_75.ToString("0.000");
            channelWordInfo.C2_U_1_100 = C2_U_1_100.ToString("0.000");

            channelWordInfo.C2_U_2_0 = C2_U_2_0.ToString("0.000");
            channelWordInfo.C2_U_2_25 = C2_U_2_25.ToString("0.000");
            channelWordInfo.C2_U_2_50 = C2_U_2_50.ToString("0.000");
            channelWordInfo.C2_U_2_75 = C2_U_2_75.ToString("0.000");
            channelWordInfo.C2_U_2_100 = C2_U_2_100.ToString("0.000");

            channelWordInfo.C2_D_1_0 = C2_D_1_0.ToString("0.000");
            channelWordInfo.C2_D_1_25 = C2_D_1_25.ToString("0.000");
            channelWordInfo.C2_D_1_50 = C2_D_1_50.ToString("0.000");
            channelWordInfo.C2_D_1_75 = C2_D_1_75.ToString("0.000");
            channelWordInfo.C2_D_1_100 = C2_D_1_100.ToString("0.000");

            channelWordInfo.C2_D_2_0 = C2_D_2_0.ToString("0.000");
            channelWordInfo.C2_D_2_25 = C2_D_2_25.ToString("0.000");
            channelWordInfo.C2_D_2_50 = C2_D_2_50.ToString("0.000");
            channelWordInfo.C2_D_2_75 = C2_D_2_75.ToString("0.000");
            channelWordInfo.C2_D_2_100 = C2_D_2_100.ToString("0.000");
            #endregion
            #endregion
        }


        private void SetDoubleChannelData(ChannelWordInfo channelWordInfo)
        {
            SetC1ChannelData(channelWordInfo);
            SetC2ChannelData(channelWordInfo);
        }

        public void ReplaceWord(string filefullname, ChannelWordInfo wordInfo)
        {
            Document doc = new Document(filefullname);
            FindReplaceOptions findReplaceOptions = new FindReplaceOptions();
            void ReplaceString(string oldstr, string newstr)
            {
                if (string.IsNullOrWhiteSpace(newstr))
                {
                    return;
                }
                doc.Range.Replace(oldstr, newstr, findReplaceOptions);
            }
            #region 替换模板数据
            #region 正常
            ReplaceString("{wono}", wordInfo.WorkOrderNo);
            ReplaceString("{no}", wordInfo.Number);
            ReplaceString("{norg}", wordInfo.NumberRange);
            ReplaceString("{qty}", wordInfo.Qty);
            ReplaceString("{tp}", wordInfo.Temperature);
            ReplaceString("{hd}", wordInfo.Humidity);
            ReplaceString("{ter}", wordInfo.Tester);
            ReplaceString("{tdate}", wordInfo.TestDate);
            ReplaceString("{aer}", wordInfo.Auditor);
            ReplaceString("{adate}", wordInfo.AuditDate);
            #endregion

            #region 1通道
            ReplaceString("{c111u}", wordInfo.C1_U_1_0);
            ReplaceString("{c112u}", wordInfo.C1_U_1_25);
            ReplaceString("{c113u}", wordInfo.C1_U_1_50);
            ReplaceString("{c114u}", wordInfo.C1_U_1_75);
            ReplaceString("{c115u}", wordInfo.C1_U_1_100);

            ReplaceString("{c111d}", wordInfo.C1_D_1_0);
            ReplaceString("{c112d}", wordInfo.C1_D_1_25);
            ReplaceString("{c113d}", wordInfo.C1_D_1_50);
            ReplaceString("{c114d}", wordInfo.C1_D_1_75);
            ReplaceString("{c115d}", wordInfo.C1_D_1_100);

            ReplaceString("{c121u}", wordInfo.C1_U_2_0);
            ReplaceString("{c122u}", wordInfo.C1_U_2_25);
            ReplaceString("{c123u}", wordInfo.C1_U_2_50);
            ReplaceString("{c124u}", wordInfo.C1_U_2_75);
            ReplaceString("{c125u}", wordInfo.C1_U_2_100);

            ReplaceString("{c121d}", wordInfo.C1_D_2_0);
            ReplaceString("{c122d}", wordInfo.C1_D_2_25);
            ReplaceString("{c123d}", wordInfo.C1_D_2_50);
            ReplaceString("{c124d}", wordInfo.C1_D_2_75);
            ReplaceString("{c125d}", wordInfo.C1_D_2_100);

            ReplaceString("{b1}", wordInfo.C1_Bjqd);
            ReplaceString("{cf1}", wordInfo.C1_Cfxwc);
            ReplaceString("{hc1}", wordInfo.C1_Hc);
            ReplaceString("{sq1}", wordInfo.C1_SQ);

            #endregion

            #region 2通道
            ReplaceString("{c211u}", wordInfo.C2_U_1_0);
            ReplaceString("{c212u}", wordInfo.C2_U_1_25);
            ReplaceString("{c213u}", wordInfo.C2_U_1_50);
            ReplaceString("{c214u}", wordInfo.C2_U_1_75);
            ReplaceString("{c215u}", wordInfo.C2_U_1_100);

            ReplaceString("{c211d}", wordInfo.C2_D_1_0);
            ReplaceString("{c212d}", wordInfo.C2_D_1_25);
            ReplaceString("{c213d}", wordInfo.C2_D_1_50);
            ReplaceString("{c214d}", wordInfo.C2_D_1_75);
            ReplaceString("{c215d}", wordInfo.C2_D_1_100);

            ReplaceString("{c221u}", wordInfo.C2_U_2_0);
            ReplaceString("{c222u}", wordInfo.C2_U_2_25);
            ReplaceString("{c223u}", wordInfo.C2_U_2_50);
            ReplaceString("{c224u}", wordInfo.C2_U_2_75);
            ReplaceString("{c225u}", wordInfo.C2_U_2_100);

            ReplaceString("{c221d}", wordInfo.C2_D_2_0);
            ReplaceString("{c222d}", wordInfo.C2_D_2_25);
            ReplaceString("{c223d}", wordInfo.C2_D_2_50);
            ReplaceString("{c224d}", wordInfo.C2_D_2_75);
            ReplaceString("{c225d}", wordInfo.C2_D_2_100);

            ReplaceString("{b2}", wordInfo.C2_Bjqd);
            ReplaceString("{cf2}", wordInfo.C2_Cfxwc);
            ReplaceString("{hc2}", wordInfo.C2_Hc);
            ReplaceString("{sq2}", wordInfo.C2_SQ);
            #endregion
            #endregion
            doc.Save(filefullname);
        }
    }
}
