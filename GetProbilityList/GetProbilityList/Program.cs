using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace GetProbilityList
{
    class Program
    {
        public static string FilePath { get; set; }
        public static DataTable DataTableAward { get; set; }
        public static DataTable DataTableKu { get; set; }
        public static readonly string[] Name1 = {"必给物品1类ID","必给物品2类ID", "必给物品3类ID", "必给物品4类ID" };
        public static readonly string[] Name2 = {"物品1类", "物品2类", "物品3类", "物品4类", "物品5类", "物品6类", "物品7类", "物品8类" };
        public static readonly string[] Number1 = {"必给物品1类数量","必给物品2类数量", "必给物品3类数量", "必给物品4类数量" };
        public static readonly string[] Number2 = {"物品1类数量", "物品2类数量", "物品3类数量", "物品4类数量", "物品5类数量", "物品6类数量", "物品7类数量", "物品8类数量" };
        public static readonly string[] Problity = {"物品1类掉率", "物品2类掉率", "物品3类掉率", "物品4类掉率", "物品5类掉率", "物品6类掉率", "物品7类掉率", "物品8类掉率" };

        static void Main()
        {
            //读表
            Program program = new Program();
            program.InitializeConfig();
            
            //获取每行的数据
            for (int i = 0; i < DataTableAward.Rows.Count; i++)
            {
                var nameList = GetNameList(i);

                Console.WriteLine(DataTableAward.Rows[i]["id说明"]);
                foreach (string t in nameList)
                {
                    Console.WriteLine(t);
                }
            }
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

            
            //int count = excelHelper.DataTableToExcel(data, "Sheet1", true);
            //if (count > 0)
            //    Console.WriteLine("Number of imported data is {0} ", count);

            //Console.ReadKey();
        }

        //读FilePath
        private static string GetAppConfig(string strKey)
        {
            string file = Process.GetCurrentProcess().MainModule.FileName;
            Configuration config = ConfigurationManager.OpenExeConfiguration(file);
            return config.AppSettings.Settings.AllKeys.Any(key => key == strKey) ? config.AppSettings.Settings[strKey].Value : null;
        }

        //获取List
        private static void GetNameList(int row,out List<string> nameList, out List<string> numberList, out List<string> probilityList)
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
                if (GetData(row, Name2[i], out string str, out int kuId) == 2)
                {
                    nameList.Add(GetItemName(kuId.ToString()));
                    numberList.Add(DataTableAward.Rows[row][Number2[i]].ToString());
                    probilityList.Add("1");
                }
                else if (GetData(row, name, out str, out int awardId) == 3)
                {
                    nameList.AddRange(GetNameList(Convert.ToInt32(GetRow(awardId.ToString()))));
                }
            }
        }

        //根据库id查找名字
        private static string GetItemName(string kuId)
        {
            return (from DataRow row in DataTableKu.Rows where row["id"].ToString() == kuId select row["名称"].ToString()).FirstOrDefault();
        }

        //根据奖励id查找row
        private static string GetRow(string awardId)
        {
            return (from DataRow row in DataTableAward.Rows where row["id"].ToString() == awardId select row["id"].ToString()).FirstOrDefault();
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
                return num > 10000 ? 2 : 3;
            }
            str = data;
            num = 0;
            return 1;
        }

        private static double GetProbility(int row, string weigh)
        {
        }



    }
}
