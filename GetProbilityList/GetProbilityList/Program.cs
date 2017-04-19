using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.SS.Formula.Functions;

namespace GetProbilityList
{
    internal class Program
    {
        public static string FilePath { get; set; }
        public static DataTable DataTableAward { get; set; }
        public static DataTable DataTableKu { get; set; }
        public static readonly string[] Name1 = { "必给物品1类ID", "必给物品2类ID", "必给物品3类ID", "必给物品4类ID" };
        public static readonly string[] Name2 = { "物品1类", "物品2类", "物品3类", "物品4类", "物品5类", "物品6类", "物品7类", "物品8类" };
        public static readonly string[] Number1 = { "必给物品1类数量", "必给物品2类数量", "必给物品3类数量", "必给物品4类数量" };
        public static readonly string[] Number2 = { "物品1类数量", "物品2类数量", "物品3类数量", "物品4类数量", "物品5类数量", "物品6类数量", "物品7类数量", "物品8类数量" };
        public static readonly string[] Problity = { "物品1类掉率", "物品2类掉率", "物品3类掉率", "物品4类掉率", "物品5类掉率", "物品6类掉率", "物品7类掉率", "物品8类掉率" };
        public static List<string> NameList = new List<string>();//名字列表
        public static List<string> NumberList = new List<string>();//数量列表
        public static List<string> ProbilityList = new List<string>();//概率列表
        public static List<int> HideRowList = new List<int>();//被调用的奖励列表，最后统计时不会显示
        public static float ProbilitySum { get; set; }//概率和

        private static void Main()
        {
            //读表
            Program program = new Program();
            program.InitializeConfig();

            //获取每行的数据
            if (DataTableAward == null) return;

            //写入文件
            FileStream fs = new FileStream(GetAppConfig(nameof(FilePath)) + @"\1.txt", FileMode.Create);
            StreamWriter sw = new StreamWriter(fs);

            for (int i = 0; i < DataTableAward.Rows.Count; i++)
            {
                ProbilitySum = Convert.ToInt32(DataTableAward.Rows[i]["总概率"]);
                GetList(i, out NameList, out NumberList, out ProbilityList);

                if (HideRowList.Contains(i))
                {
                    continue;
                }

                //开始写入
                //Console.WriteLine($"执行到第{i}行");
                sw.WriteLine(DataTableAward.Rows[i]["id说明"]);
                sw.WriteLine("道具" + "    " +"数量" + "    " +"概率");
                for (int j = 0; j < NameList.Count; j++)
                {
                    sw.WriteLine(NameList[j] + "    " + NumberList[j] + "    " + ProbilityList[j]);
                }
                sw.WriteLine();
            }
            //清空缓冲区
            sw.Flush();
            //关闭流
            sw.Close();
            fs.Close();
            Console.WriteLine("开饭啦！");
            Console.ReadKey();
        }

        //从表里读奖励和库
        private void InitializeConfig()
        {
            FilePath = GetAppConfig(nameof(FilePath));

            if (!Directory.Exists(FilePath))
            {
                Console.WriteLine($"\"{FilePath}\"未找到，请修改FestivalActivity.exe.config中的文件夹路径！");
                return;
            }

            FilePath = FilePath + @"\1.xlsx";

            ExcelHelper excelHelper = new ExcelHelper(FilePath);
            DataTableAward = excelHelper.ExcelToDataTable("Sheet1", true);
            DataTableKu = excelHelper.ExcelToDataTable("Sheet2", true);
        }

        //读FilePath
        private static string GetAppConfig(string strKey)
        {
            string file = Process.GetCurrentProcess().MainModule.FileName;
            Configuration config = ConfigurationManager.OpenExeConfiguration(file);
            return config.AppSettings.Settings.AllKeys.Any(key => key == strKey) ? config.AppSettings.Settings[strKey].Value : null;
        }

        //获取List
        private static void GetList(int row, out List<string> nameList, out List<string> numberList, out List<string> probilityList)
        {
            nameList = new List<string>();
            numberList = new List<string>();
            probilityList = new List<string>();

            //必得道具
            for (int i = 0; i < 4; i++)
            {
                if (GetData(row, Name1[i], out string _, out int kuId) != 2) break;
                nameList.Add(GetItemName(kuId.ToString()));
                numberList.Add(DataTableAward.Rows[row][Number1[i]].ToString());
                probilityList.Add("1");
            }
            //随机道具
            for (int i = 0; i < 8; i++)
            {
                if (GetData(row, Name2[i], out string str, out int awardId) == 3)
                {
                    //记录不显示的row,找不到的话，记录下奖励id
                    if (GetRow(awardId.ToString()) == 0)
                    {
                        Console.WriteLine($"awardId:{awardId}");
                    }
                    else
                    {
                        HideRowList.Add(GetRow(awardId.ToString()));
                    }
                   
                    GetList(GetRow(awardId.ToString()), out List<string> nameList2, out List<string> numberList2,
                        out List<string> probilityList2);
                    nameList.AddRange(nameList2);
                    numberList.AddRange(numberList2);
                    probilityList.AddRange(probilityList2);
                    break;
                }
                if (GetData(row, Name2[i], out str, out int kuId) != 2) continue;
                nameList.Add(GetItemName(kuId.ToString()));
                numberList.Add(DataTableAward.Rows[row][Number2[i]].ToString());

                if (GetData(row, Problity[i], out string probilityStr, out int probilityNum) == 3 || GetData(row, Problity[i], out probilityStr, out probilityNum) == 2)
                {
                    probilityList.Add((probilityNum / ProbilitySum).ToString(CultureInfo.InvariantCulture));
                }
                else
                {
                    probilityList.Add(probilityStr + "//" + DataTableAward.Rows[row]["总概率"]);
                }
            }
        }

        //根据库id查找名字
        private static string GetItemName(string kuId)
        {
            return (from DataRow row in DataTableKu.Rows where row["id"].ToString() == kuId select row["名称"].ToString()).FirstOrDefault();
        }

        //根据奖励id查找row
        private static int GetRow(string awardId)
        {
            return Convert.ToInt32((from DataRow row in DataTableAward.Rows where row["id"].ToString() == awardId select DataTableAward.Rows.IndexOf(row)).FirstOrDefault());
        }

        //0为空或0
        //1为string，data以str变量返回
        //2为>10000的值，data以num变量返回
        //3为<=10000的值，data以num变量返回
        private static int GetData(int row, string col, out string str, out int num)
        {
            if (DataTableAward.Rows[row][col] == null)
            {
                str = null;
                num = 0;
                return 0;
            }
            string data = DataTableAward.Rows[row][col].ToString();
            if (int.TryParse(data, out num))
            {
                str = null;
                return num > 10000 ? 2 : (num > 0 ? 3 : 0);
            }
            str = data;
            num = 0;
            return 1;
        }
    }
}