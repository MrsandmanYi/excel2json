using System;
using System.IO;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Runtime.Remoting.Contexts;


namespace excel2json
{

    public class ConfigParams
    {
        public string clsName = "";
        public bool isDoubleKey = false;
    }

    /// <summary>
    /// 应用程序
    /// </summary>
    sealed partial class Program
    {
        /// <summary>
        /// 应用程序入口
        /// </summary>
        /// <param name="args">命令行参数</param>
        [STAThread]
        static void Main(string[] args)
        {
            if (args.Length <= 0)
            {
                //-- GUI MODE ----------------------------------------------------------
                Console.WriteLine("Launch excel2json GUI Mode...");
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new GUI.MainForm());
            }
            else
            {
                //-- COMMAND LINE MODE -------------------------------------------------

                //-- 分析命令行参数
                var options = new Options();
                var parser = new CommandLine.Parser(with => with.HelpWriter = Console.Error);

                if (parser.ParseArgumentsStrict(args, options, () => Environment.Exit(-1)))
                {
                    //-- 执行导出操作
                    try
                    {
                        DateTime startTime = DateTime.Now;
                        Run(options);
                        //-- 程序计时
                        DateTime endTime = DateTime.Now;
                        TimeSpan dur = endTime - startTime;
                        Console.WriteLine(
                            string.Format("[{0}]：\tConversion complete in [{1}ms].",
                            Path.GetFileName(options.ExcelPath),
                            dur.TotalMilliseconds)
                            );
                    }
                    catch (Exception exp)
                    {
                        Console.WriteLine("Error: " + exp.Message);
                    }
                }
            }// end of else
        }

        /// <summary>
        /// 根据命令行参数，执行Excel数据导出工作
        /// </summary>
        /// <param name="options">命令行参数</param>
        private static void Run(Options options)
        {
            string excelPath = options.ExcelPath;

            if (options.ExportMode == 0) //-- Excel File 
            {
                ExportExcel(options, excelPath);
            }
            else if (options.ExportMode == 1) //-- Excel Folder 
            {
                // 获取文件夹下所有的Excel文件，包括子文件夹下的
                string[] files = Directory.GetFiles(excelPath, "*.xlsx", SearchOption.AllDirectories);
                List<ConfigParams> configParamsList = new List<ConfigParams>();
                foreach (string file in files)
                {
                    configParamsList.Add(ExportExcel(options, file));
                }
                // 生成ClassConfig

                try
                {
                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine("using System;");
                    sb.AppendLine("using System.Collections.Generic;");
                    sb.AppendLine();
                    sb.AppendLine("namespace GameConfigClass");
                    sb.AppendLine("{");
                    sb.AppendLine("\tpublic class ConfigClassStruct");
                    sb.AppendLine("\t{");
                    sb.AppendLine("\t\tpublic string name;");
                    sb.AppendLine("\t\tpublic Func<string,object> jConvertMethod;");
                    sb.AppendLine("\t\tpublic bool doubleKey;");
                    sb.AppendLine("\t}");
                    sb.AppendLine();

                    sb.AppendLine("\tpublic static class ExcelClassConfig");
                    sb.AppendLine("\t{");
                    sb.AppendLine("\t\tpublic static Dictionary<string, ConfigClassStruct> configClassMap = new Dictionary<string, ConfigClassStruct>()");
                    sb.AppendLine("\t\t{");
                    foreach (ConfigParams cp in configParamsList)
                    {
                        sb.AppendLine($"\t\t\t{{\"{cp.clsName}\", new ConfigClassStruct(){{");
                        sb.AppendLine($"\t\t\t\tname = \"{cp.clsName}\",");
                        sb.AppendLine("\t\t\t\tjConvertMethod = (json) => {{");
                        if (cp.isDoubleKey)
                        {
                            sb.AppendLine(string.Format("\t\t\t\t\tvar data = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<int, Dictionary<int, {0}>>>(json);", cp.clsName));
                        }
                        else
                        {
                            sb.AppendLine($"\t\t\t\t\tvar data = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<int, {cp.clsName}>>(json);");
                        }
                        sb.AppendLine("\t\t\t\t\treturn data;");
                        sb.AppendLine("\t\t\t\t}},");
                        sb.AppendLine(string.Format("\t\t\t\tdoubleKey = {0},", cp.isDoubleKey ? "true" : "false"));
                        sb.AppendLine("\t\t\t\t}");
                        sb.AppendLine("\t\t\t},");  
                    }
                    sb.AppendLine("\t\t};");
                    sb.AppendLine("\t}");
                    sb.AppendLine("}");
                    sb.AppendLine();

                    string classConfigPath = Path.Combine(options.ClsConfigPath, "ExcelClassConfig.cs");
                    Console.WriteLine("ClassConfigPath: " + classConfigPath);
                    // 保存配置文件
                    File.WriteAllText(classConfigPath, sb.ToString(), new UTF8Encoding(false));
                }
                catch (Exception e)
                {
                    Console.WriteLine("Error: 导出ClassConfigPath失败!!!");
                    Console.WriteLine(e);
                    throw;
                }
            }
        }

        private static ConfigParams ExportExcel(Options options, string excelPath)
        {
            ConfigParams configParams = new ConfigParams();

            string excelName = Path.GetFileNameWithoutExtension(excelPath);

            //-- Header
            int header = options.HeaderRows;

            //-- Encoding
            Encoding cd = new UTF8Encoding(false);
            if (options.Encoding != "utf8-nobom")
            {
                foreach (EncodingInfo ei in Encoding.GetEncodings())
                {
                    Encoding e = ei.GetEncoding();
                    if (e.HeaderName == options.Encoding)
                    {
                        cd = e;
                        break;
                    }
                }
            }

            //-- Date Format
            string dateFormat = options.DateFormat;

            //-- Export path
            string exportPath;
            if (options.JsonPath != null && options.JsonPath.Length > 0)
            {
                if (options.ExportMode == 1)
                {
                    exportPath = Path.Combine(options.JsonPath, excelName + ".json");
                }
                else
                {
                    exportPath = options.JsonPath;
                }
            }
            else
            {
                exportPath = Path.ChangeExtension(excelPath, ".json");
            }

            //-- Load Excel
            ExcelLoader excel = new ExcelLoader(excelPath, header);

            //-- export
            JsonExporter exporter = new JsonExporter(excel, options.Lowcase, options.ExportArray, dateFormat, options.ForceSheetName, header, options.ExcludePrefix, options.CellJson, options.AllString);
            exporter.SaveToFile(exportPath, cd);

            //-- 生成C#定义文件
            if (options.CSharpPath != null && options.CSharpPath.Length > 0)
            {
                var cSharpPath = options.CSharpPath;
                if (options.ExportMode == 1)
                {
                    cSharpPath = Path.Combine(options.CSharpPath, excelName + ".cs");
                }

                CSDefineGenerator generator = new CSDefineGenerator(excelName, excel, options.ExcludePrefix);
                generator.SaveToFile(cSharpPath, cd);
            }
            configParams.clsName = excelName;
            configParams.isDoubleKey = exporter.isDoubleKey;
            return configParams;
        }
    }
}
