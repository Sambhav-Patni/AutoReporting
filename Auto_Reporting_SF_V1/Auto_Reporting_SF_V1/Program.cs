//SF TLV Reporting Automation StateFarm Specific weekly reports
// By Sambhav Patni
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;

using ExcelLibrary.SpreadSheet;
using System.Data;
using System.IO;
using System.Data.OracleClient;
using Microsoft.Vbe.Interop;
using System.Data.OleDb;
using System.Reflection;
using Microsoft.Exchange.WebServices.Data;

//using Microsoft.Office.Interop;

namespace Auto_Reporting_SF_V1
{
    class Program
    {
        //static string path = @"C:\Users\sambhav.patni\Documents\Visual Studio 2012\Projects\Auto_Reporting_SF_V1\Auto_Reporting_SF_V1\bin\Debug\book.xls";
        //static string Output = @"C:\Users\sambhav.patni\Documents\Visual Studio 2012\Projects\Auto_Reporting_SF_V1\Auto_Reporting_SF_V1\bin\Debug\SF_AllStates_Weekly_20-26_October.xlsx";
        static string[] config;
        static int TSDays = 7;
        static string dateTo = DateTime.Now.Date.ToString("dd-MMM-yyyy");//"12-JAN-2015";
        static string dateFrom = DateTime.Now.Date.Subtract(TimeSpan.FromDays(TSDays)).ToString("dd-MMM-yyyy");//"05-JAN-2015";        
        static string path_pre = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\";
        static string path = path_pre + "book.xls";
        static string Output_name = "SF_AllStates_PlaceHolder<>.xls";
        static string Output = path_pre + "SF_AllStates_PlaceHolder<>.xls";
        static string spLocation = "";
        static string spLocationUri = "";
        static bool upload = true;
        static string NmapPass = "XXX";
        static string EpfpPass = "XXX";
        static string[] Months = { "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" };
        static void Main(string[] args)
        {
            string temp1 = Convert.ToDateTime(dateFrom).ToShortDateString();
            string temp2 = Convert.ToDateTime(dateTo).ToShortDateString();
            string bookQuery = " select c.orig_valreq_id, c.claim_nbr||'-'||c.exposure_nbr as CLAIM_NBR, c.exposure_nbr exp, upper(c.created_by), dbms_lob.substr(e.rpt_content,100, dbms_lob.instr(e.rpt_content,'FirstName', (dbms_lob.instr(e.rpt_content,'<UserInfo',1,1)),1)) as FirstName, dbms_lob.substr(e.rpt_content,100, dbms_lob.instr(e.rpt_content,'LastName', (dbms_lob.instr(e.rpt_content,'<UserInfo',1,1)),1)) as LastName, d.VIN, d.model_year, d.vcd_make_desc, d.vcd_model_desc, d.mileage,d.loss_dt, d.loss_location_zip loss_zip,  e.adj_market_value market_value, dbms_lob.substr(e.rpt_content,105,dbms_lob.instr(e.rpt_content,'<adj:MainSubCategoryWeight>',1,1)-50) as OverallConditionScorewithTags, dbms_lob.substr(e.rpt_content,110,dbms_lob.instr(e.rpt_content,'<adj:AdjustmentType>OVERALL</adj:AdjustmentType>',1,1)) as ConditionAdjAmtwithTags, dbms_lob.substr(e.rpt_content,102,dbms_lob.instr(e.rpt_content,'<adj:EquipmentAdjustment>',1,1))as AfterMarketTotal, d.loss_location_state, dbms_lob.substr(e.rpt_content,105,dbms_lob.instr(e.rpt_content,'<adj:RefurbishmentAdjustment>',1,1))as RefurbishmentAdjwithTags, dbms_lob.substr(e.rpt_content,98,dbms_lob.instr(e.rpt_content,'<adj:PriorDamageAdjustment>',1,1)) as PriorDamageAdjwithTags, e.location_url as location_url, (case when dbms_lob.substr(e.rpt_content,85,dbms_lob.instr(e.rpt_content,'Level=\"OFFICE\">',1,1)-65) is null then dbms_lob.substr(e.rpt_content,85,dbms_lob.instr(e.rpt_content,'<HierNode Level=\"OFFICE\"',1,1)) else dbms_lob.substr(e.rpt_content,85,dbms_lob.instr(e.rpt_content,'Level=\"OFFICE\">',1,1)-65)    end) as Office from  tlv_valuation_request c, tlv_vr_vehicle d, tlv_nada_reports e where c.co_cd in ('SF','S3') and c.id = d.valreq_id (+) and (e.created_dt >= trunc(TO_DATE('" + Convert.ToDateTime(dateFrom).ToShortDateString() + "', 'MM-DD-YYYY')) and e.created_dt < trunc(TO_DATE('" + Convert.ToDateTime(dateTo).ToShortDateString() + "', 'MM-DD-YYYY'))) and d.valreq_id = e.valreq_id (+)  UNION ALL  select c.orig_valreq_id, c.claim_nbr||'-'||c.exposure_nbr as CLAIM_NBR, c.exposure_nbr exp, upper(c.created_by), dbms_lob.substr(e.rpt_content,100, dbms_lob.instr(e.rpt_content,'FirstName', (dbms_lob.instr(e.rpt_content,'<UserInfo',1,1)),1)) as FirstName, dbms_lob.substr(e.rpt_content,100, dbms_lob.instr(e.rpt_content,'LastName', (dbms_lob.instr(e.rpt_content,'<UserInfo',1,1)),1)) as LastName, d.VIN, d.model_year, d.vcd_make_desc, d.vcd_model_desc, d.mileage,d.loss_dt, d.loss_location_zip loss_zip,  e.adj_market_value market_value, dbms_lob.substr(e.rpt_content,105,dbms_lob.instr(e.rpt_content,'<adj:MainSubCategoryWeight>',1,1)-50) as OverallConditionScorewithTags, dbms_lob.substr(e.rpt_content,110,dbms_lob.instr(e.rpt_content,'<adj:AdjustmentType>OVERALL</adj:AdjustmentType>',1,1)) as ConditionAdjAmtwithTags, dbms_lob.substr(e.rpt_content,102,dbms_lob.instr(e.rpt_content,'<adj:EquipmentAdjustment>',1,1))as AfterMarketTotal, d.loss_location_state, dbms_lob.substr(e.rpt_content,105,dbms_lob.instr(e.rpt_content,'<adj:RefurbishmentAdjustment>',1,1))as RefurbishmentAdjwithTags, dbms_lob.substr(e.rpt_content,98,dbms_lob.instr(e.rpt_content,'<adj:PriorDamageAdjustment>',1,1)) as PriorDamageAdjwithTags, e.location_url as location_url, (case when dbms_lob.substr(e.rpt_content,85,dbms_lob.instr(e.rpt_content,'Level=\"OFFICE\">',1,1)-65) is null then dbms_lob.substr(e.rpt_content,85,dbms_lob.instr(e.rpt_content,'<HierNode Level=\"OFFICE\"',1,1)) else dbms_lob.substr(e.rpt_content,85,dbms_lob.instr(e.rpt_content,'Level=\"OFFICE\">',1,1)-65)    end) as Office from  tlv_valuation_request c, tlv_vr_vehicle d, tlv_redbook_reports e where c.co_cd in ('SF','S3') and c.id = d.valreq_id (+) and (e.created_dt >= trunc(TO_DATE('" + Convert.ToDateTime(dateFrom).ToShortDateString() + "', 'MM-DD-YYYY')) and e.created_dt < trunc(TO_DATE('" + Convert.ToDateTime(dateTo).ToShortDateString() + "', 'MM-DD-YYYY'))) and d.valreq_id = e.valreq_id (+) ";
            string cvdbQuery = " select c.orig_valreq_id, c.claim_nbr||'-'||c.exposure_nbr as CLAIM_NBR, c.exposure_nbr exp, upper(c.created_by), dbms_lob.substr(e.rpt_content,100, dbms_lob.instr(e.rpt_content,'FirstName', (dbms_lob.instr(e.rpt_content,'<UserInfo',1,1)),1)) as FirstName, dbms_lob.substr(e.rpt_content,100, dbms_lob.instr(e.rpt_content,'LastName', (dbms_lob.instr(e.rpt_content,'<UserInfo',1,1)),1)) as LastName, d.VIN, d.model_year, d.vcd_make_desc, d.vcd_model_desc, d.mileage,d.loss_dt, d.loss_location_zip loss_zip,  e.adj_market_value market_value, dbms_lob.substr(e.rpt_content,105,dbms_lob.instr(e.rpt_content,'<adj:MainSubCategoryWeight>',1,1)-50) as OverallConditionScorewithTags, dbms_lob.substr(e.rpt_content,110,dbms_lob.instr(e.rpt_content,'<adj:AdjustmentType>OVERALL</adj:AdjustmentType>',1,1)) as ConditionAdjAmtwithTags, dbms_lob.substr(e.rpt_content,102,dbms_lob.instr(e.rpt_content,'<adj:EquipmentAdjustment>',1,1))as AfterMarketTotal, d.loss_location_state, dbms_lob.substr(e.rpt_content,105,dbms_lob.instr(e.rpt_content,'<adj:RefurbishmentAdjustment>',1,1))as RefurbishmentAdjwithTags, dbms_lob.substr(e.rpt_content,98,dbms_lob.instr(e.rpt_content,'<adj:PriorDamageAdjustment>',1,1)) as PriorDamageAdjwithTags, e.location_url as location_url, (case when dbms_lob.substr(e.rpt_content,85,dbms_lob.instr(e.rpt_content,'Level=\"OFFICE\">',1,1)-65) is null then dbms_lob.substr(e.rpt_content,95,dbms_lob.instr(e.rpt_content,'<HierNode Level=\"OFFICE\"',1,1)) else dbms_lob.substr(e.rpt_content,85,dbms_lob.instr(e.rpt_content,'Level=\"OFFICE\">',1,1)-65)    end) as Office from  tlv_valuation_request c, tlv_vr_vehicle d, tlv_comp_vehicle_reports e where c.co_cd in ('SF','S3') and c.id = d.valreq_id (+) and (e.created_dt >= trunc(TO_DATE('" + Convert.ToDateTime(dateFrom).ToShortDateString() + "', 'MM-DD-YYYY')) and e.created_dt < trunc(TO_DATE('" + Convert.ToDateTime(dateTo).ToShortDateString() + "', 'MM-DD-YYYY'))) and d.valreq_id = e.valreq_id (+)";
            Output = Output.Replace("PlaceHolder<>", Convert.ToDateTime(dateFrom).Day + "-" + (Convert.ToDateTime(dateTo).Subtract(TimeSpan.FromDays(1)).Day) + "_" + Months[Convert.ToDateTime(dateTo).Month - 1] + "_" + Convert.ToDateTime(dateTo).Year);
            Output_name = Output_name.Replace("PlaceHolder<>", Convert.ToDateTime(dateFrom).Day + "-" + (Convert.ToDateTime(dateTo).Subtract(TimeSpan.FromDays(1)).Day) + "_" + Months[Convert.ToDateTime(dateTo).Month - 1] + "_" + Convert.ToDateTime(dateTo).Year);
            Console.Write("\n\n\t\t\tHello, Lets Get Started with Reporting\n");
            Console.WriteLine("Starting BOOK");
            File.WriteAllText(path_pre + "Log.txt", "\n\n\t\t\tHello, Lets Get Started with Reporting\n");
            File.AppendAllText(path_pre + "Log.txt", "Starting BOOK\n");
            try
            {
                config = File.ReadAllLines(path_pre + "DBConfig.sam");
                NmapPass = config[0].Substring(config[0].LastIndexOf(':')+1).Trim();
                EpfpPass = config[1].Substring(config[1].LastIndexOf(':')+1).Trim();
                Retrive_Data(bookQuery);
                File.Copy(path, path.Replace(".xls","_backup.xls"), true);
                Console.WriteLine("BOOK VBA");
                File.AppendAllText(path_pre + "Log.txt", "BOOK VBA");
                vba("book");
                Finalize();
                Console.WriteLine("Starting CVDB");
                File.AppendAllText(path_pre + "Log.txt", "Starting CVDB");
                //path = @"C:\Users\sambhav.patni\Documents\Visual Studio 2012\Projects\Auto_Reporting_SF_V1\Auto_Reporting_SF_V1\bin\Debug\cvdb.xls";
                path = path_pre + "cvdb.xls";
                Retrive_Data(cvdbQuery);
                File.Copy(path, path.Replace(".xls", "_backup.xls"), true);
                vba("comp");
                Finalize();
                if (upload == true)
                {
                    try
                    {
                        upload = false;
                        spUpload();
                        upload = true;
                    }
                    catch (Exception ex)
                    {
                        File.WriteAllText(path_pre + "SPErrorLog.txt", "Error: \n" + ex.Message + "\n" + ex.StackTrace);
                    }
                }

                mail(false);
            }
            catch (Exception ex)
            {
                File.WriteAllText(path_pre + "ErrorLog.txt", ex.Message + "\n" + ex.StackTrace);
                Console.WriteLine(ex.Message); mail(true, ex.Message);
            }
        }
        public static void Retrive_Data(string query)
        {
            try
            {
                using (OracleConnection connection = new OracleConnection("User Id=sp102532;Password="+NmapPass+";Data Source=nmap.mitchell.com"))
                    {
                        DataSet ds = new DataSet();
                        connection.Open();
                        OracleCommand command;
                        OracleDataAdapter adaptor; 
                        DataTable dtschema = connection.GetSchema();
                        try
                        {
                            //NADA REDBOOK COMBINED
                            command = new OracleCommand(query, connection);
                            adaptor = new OracleDataAdapter(command);
                            Console.WriteLine("\nCommand Set: "+command.CommandText);
                            File.AppendAllText(path_pre + "Log.txt", "\nCommand Set: " + command.CommandText);
                            adaptor.Fill(ds);
                            Console.WriteLine("\nGot Data From "+connection.DataSource);
                            File.AppendAllText(path_pre + "Log.txt", "\nGot Data From " + connection.DataSource);
                        }
                        catch (IndexOutOfRangeException)
                        {
                            //command = new OracleCommand("select * from [" + dtschema.Rows[2][2] + "] WHERE [Client ID] IN (" + ISTeam + ")", connection);
                            //command1 = new OleDbCommand("select * from [" + dtschema.Rows[2][2] + "] WHERE [Client ID] NOT IN (" + ISTeam + ")", connection);
                        }                                               
                        CreateWorkbook(path, ds);                      
                                              
                    }
            }
            //catch (OleDbException)
            //{
            //    //MessageBox.Show("Excel Not in Proper Format:\n Try Selecting \"Open\" CheckBox , \n\n Contact: Sambhav Patni\n At: Sambhav.patni@infogain.com", "Error Occured");
            //}
            //catch (FileNotFoundException ex)
            //{
            //    //if (ex.Message.Contains("ISOpsTeam"))
            //    //    //MessageBox.Show("Config File Does Not Exist...\n" + ex.Message);
            //    //else
            //    //    //MessageBox.Show("File Does Not Exist...\n" + ex.Message);
            //}
            catch (Exception ex)
            {
                //MessageBox.Show("Something Went Wrong...\n\n Contact: Sambhav Patni\n At: Sambhav.patni@infogain.com\n\n" + ex.Message, "Error Occured");
                Console.WriteLine("\nException: "+ex.Message);
                File.AppendAllText(path_pre + "Log.txt", "\nException: " + ex.Message);
            }
        }
        public static void CreateWorkbook(String filePath, DataSet dataset)
        {
            Console.Write("Writing Excel");
            File.AppendAllText(path_pre + "Log.txt", "Writing Excel\n");
            if (dataset.Tables.Count == 0)
                throw new ArgumentException("DataSet needs to have at least one DataTable", "dataset");            

            Workbook workbook = new Workbook();                       
            foreach (DataTable dt in dataset.Tables)
            {
                Worksheet worksheet = new Worksheet(dt.TableName);                
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    worksheet.Cells[0, i] = new Cell(dt.Columns[i].ColumnName);
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        try
                        {
                            //Convert.ToDateTime(dt.Rows[j][i]);
                            worksheet.Cells[j + 1, i] = new Cell(dt.Rows[j][i] == DBNull.Value ? string.Empty : dt.Rows[j][i]);
                        }
                        //catch (FormatException)
                        //{
                        //    worksheet.Cells[j + 1, i] = new Cell(dt.Rows[j][i] == DBNull.Value ? string.Empty : dt.Rows[j][i]);
                        //}
                        //catch (InvalidCastException)
                        //{
                        //    worksheet.Cells[j + 1, i] = new Cell(dt.Rows[j][i] == DBNull.Value ? string.Empty : dt.Rows[j][i]);
                        //    //worksheet.Cells[j + 1, i] = new Cell("");
                        //}
                        catch (Exception)
                        {
                            worksheet.Cells[j + 1, i] = new Cell(dt.Rows[j][i] == DBNull.Value ? string.Empty : dt.Rows[j][i]);
                        }
                    }
                }
                workbook.Worksheets.Add(worksheet);
            }            
            workbook.Save(filePath);
        }
        public static void InsertWorkbook(String filePath, DataSet dataset)
        {
            Console.WriteLine("\nWriting Excel");
            File.AppendAllText(path_pre + "Log.txt", "Insert Excel\n");
            if (dataset.Tables.Count == 0)
                throw new ArgumentException("DataSet needs to have at least one DataTable", "dataset");

            Workbook workbook = Workbook.Load(filePath);
            
            foreach (DataTable dt in dataset.Tables)
            {
                Worksheet worksheet = workbook.Worksheets[1];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    worksheet.Cells[0, i] = new Cell(dt.Columns[i].ColumnName);
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        try
                        {
                            //Convert.ToDateTime(dt.Rows[j][i]);
                            worksheet.Cells[j + 1, i] = new Cell(dt.Rows[j][i] == DBNull.Value ? string.Empty : dt.Rows[j][i]);
                        }
                        //catch (FormatException)
                        //{
                        //    worksheet.Cells[j + 1, i] = new Cell(dt.Rows[j][i] == DBNull.Value ? string.Empty : dt.Rows[j][i]);
                        //}
                        //catch (InvalidCastException)
                        //{
                        //    worksheet.Cells[j + 1, i] = new Cell(dt.Rows[j][i] == DBNull.Value ? string.Empty : dt.Rows[j][i]);
                        //    //worksheet.Cells[j + 1, i] = new Cell("");
                        //}
                        catch (Exception)
                        {
                            worksheet.Cells[j + 1, i] = new Cell(dt.Rows[j][i] == DBNull.Value ? string.Empty : dt.Rows[j][i]);
                        }
                    }
                }
                //workbook.Worksheets.Add(worksheet);
            }
            workbook.Save(filePath);
        }
        public static void vba(string type)
        {            
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;
            //FileStream temp = File.OpenRead(path);
            string workbookPath = path;//temp.Name;
            //temp.Close();
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(workbookPath,
                0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "",
                true, false, 0, true, false, false);
            var newStandardModule = excelWorkbook.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
            var codeModule = newStandardModule.CodeModule;
            var lineNum = codeModule.CountOfLines + 1;
            var codeText = File.ReadAllText(path_pre + "Process.vb");            
            codeModule.InsertLines(lineNum, codeText);
            excelWorkbook.Save();
            Console.WriteLine("macro_StartProcessing");
            File.AppendAllText(path_pre + "Log.txt", "macro_StartProcessing\n");
            var macro = string.Format("{0}!{1}.{2}", excelWorkbook.Name, newStandardModule.Name, "StartProcessing");
            excelApp.Run(macro);
            excelWorkbook.Save();
            Console.WriteLine("macro_SeparateOriginals");
            File.AppendAllText(path_pre + "Log.txt", "macro_SeparateOriginals\n");
            if(type.Equals("book"))
                macro = string.Format("{0}!{1}.{2}", excelWorkbook.Name, newStandardModule.Name, "SeperateOriginals_book");//SeperateOriginals_book
            else
                macro = string.Format("{0}!{1}.{2}", excelWorkbook.Name, newStandardModule.Name, "SeperateOriginals_comp");//SeperateOriginals_comp
            excelApp.Run(macro);
            excelWorkbook.Save();
            excelWorkbook.Close();                        
            Fill_MapUserId();
            excelWorkbook = excelApp.Workbooks.Open(workbookPath,
                0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "",
                true, false, 0, true, false, false);
            newStandardModule = excelWorkbook.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
            codeModule = newStandardModule.CodeModule;
            lineNum = codeModule.CountOfLines + 1;
            codeModule.InsertLines(lineNum, codeText);
            excelWorkbook.Save();
            Console.WriteLine("macro_mapUserId");
            File.AppendAllText(path_pre + "Log.txt", "macro_mapUserId\n");
            macro = string.Format("{0}!{1}.{2}", excelWorkbook.Name, newStandardModule.Name, "mapUserId");//mapUserId
            excelApp.Run(macro);
            Console.WriteLine("macro_Finalize");
            File.AppendAllText(path_pre + "Log.txt", "macro_Finalize\n");
            macro = string.Format("{0}!{1}.{2}", excelWorkbook.Name, newStandardModule.Name, "Finalize");//Finalize
            excelApp.Run(macro);
            Console.WriteLine("macro_Beautify");
            File.AppendAllText(path_pre + "Log.txt", "macro_Beautify\n");
            macro = string.Format("{0}!{1}.{2}", excelWorkbook.Name, newStandardModule.Name, "Beautify");//Finalize
            excelApp.Run(macro);
            excelApp.Visible = true;
            excelWorkbook.Save();
            excelWorkbook.Close();
            excelApp.Quit();
        }
        static void Fill_MapUserId()
        {
            Console.WriteLine("\n Mapping UserId");
            File.AppendAllText(path_pre + "Log.txt", "\n Mapping UserId\n");
            FileStream temp = File.OpenRead(path);
            string ExcelPath = temp.Name;
            temp.Close();
            //string ExcelPath = File.OpenRead(path).Name;            
            DataSet ds = new DataSet();            
            DataSet ds_orclFinal = new DataSet();
            OleDbDataAdapter adaptor;
            using (OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + ExcelPath + ";Extended Properties='Excel 12.0;HDR=Yes;';"/*con*/))
            {
                connection.Open();
                OleDbCommand command;
                command = new OleDbCommand("SELECT DISTINCT(ID) FROM (select [Original USER ID] AS ID from [OriginalCreators$] UNION ALL select [Revised USER ID] AS ID from [OriginalCreators$])", connection);
                adaptor = new OleDbDataAdapter(command);
                DataTable dtschema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                adaptor.Fill(ds);                
                int rowCount = 0;
                DataSet ds_orcl;
                    string QueryAdd = "'";                    
                    foreach (DataRow dr in ds.Tables[0].Rows)
                    {
                        if (rowCount < 1000)
                        {
                            try
                            {
                                if (dr["ID"] != null && !Convert.ToString(dr["ID"]).Equals(""))
                                {
                                    QueryAdd += dr["ID"] + "','";
                                }
                                rowCount++;
                            }
                            catch (Exception) { }
                        }
                        else
                        {
                            QueryAdd = QueryAdd.Remove(QueryAdd.Length - 2, 2);
                            ds_orcl = getUserName(QueryAdd);
                            ds_orclFinal.Tables.Add(new DataTable());
                            ds_orclFinal.Tables[0].Merge(ds_orcl.Tables[0]);
                            QueryAdd = "'";
                            rowCount = 0;
                            try
                            {
                                if (dr["ID"] != null && !Convert.ToString(dr["ID"]).Equals(""))
                                {
                                    QueryAdd += dr["ID"] + "','";
                                }
                                rowCount++;
                            }
                            catch (Exception) { }
                        }
                    }
                    QueryAdd = QueryAdd.Remove(QueryAdd.Length - 2, 2);
                    ds_orcl = getUserName(QueryAdd);
                    ds_orclFinal.Tables.Add(new DataTable());
                    ds_orclFinal.Tables[0].Merge(ds_orcl.Tables[0]);
            }
            InsertWorkbook(ExcelPath, ds_orclFinal);
        }
        static DataSet getUserName(string QueryAdd)
        {
            string QueryDistinct = "select org_cd as \"USER ID\",upper(org_name)as \"USER NAME\" from epd_org  where upper(org_cd) in ( PlaceHolder )  and org_type_name ='USER'  and co_cd in ('SF','S3')  order by org_cd";
            QueryDistinct = QueryDistinct.Replace("PlaceHolder", QueryAdd);
            DataSet ds_orcl = new DataSet();
            using (OracleConnection connection_orcl = new OracleConnection("User Id=sp102532;Password="+EpfpPass+";Data Source=epfp.mitchell.com"))
            {
                connection_orcl.Open();
                OracleCommand command_orcl;
                OracleDataAdapter adaptor_orcl;
                //DataTable dtschema_orcl = connection.GetSchema();
                try
                {
                    //MapUserId
                    command_orcl = new OracleCommand(QueryDistinct, connection_orcl);
                    adaptor_orcl = new OracleDataAdapter(command_orcl);
                    Console.WriteLine("\nCommand Set: " + command_orcl.CommandText);
                    File.AppendAllText(path_pre + "Log.txt", "\nCommand Set: " + command_orcl.CommandText);
                    adaptor_orcl.Fill(ds_orcl);
                    Console.WriteLine("\nGot Data From " + connection_orcl.DataSource);
                    File.AppendAllText(path_pre + "Log.txt", "\nGot Data From " + connection_orcl.DataSource);
                }
                catch (IndexOutOfRangeException)
                {
                    //command = new OracleCommand("select * from [" + dtschema.Rows[2][2] + "] WHERE [Client ID] IN (" + ISTeam + ")", connection);
                    //command1 = new OleDbCommand("select * from [" + dtschema.Rows[2][2] + "] WHERE [Client ID] NOT IN (" + ISTeam + ")", connection);
                }
            }
            return ds_orcl;
        }
        static void Finalize()
        {
            Console.WriteLine("Merging Book and Comparable");
            File.AppendAllText(path_pre + "Log.txt", "Merging Book and Comparable");
            FileStream temp = File.OpenRead(path);
            string ExcelPath = temp.Name;
            temp.Close();
            if (!File.Exists(Output))
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel._Workbook xlWB = (Microsoft.Office.Interop.Excel._Workbook)xlApp.Workbooks.Add(Missing.Value);
                Microsoft.Office.Interop.Excel._Worksheet xlSheet = (Microsoft.Office.Interop.Excel._Worksheet)xlWB.Sheets[1];
                Microsoft.Office.Interop.Excel._Worksheet xlSheet1 = (Microsoft.Office.Interop.Excel._Worksheet)xlWB.Sheets[2];
                ((Microsoft.Office.Interop.Excel._Worksheet)xlWB.Sheets[3]).Delete();
                //((Microsoft.Office.Interop.Excel._Worksheet)xlWB.Sheets[2]).Delete();

                xlSheet.Name = "Comparable";
                xlSheet1.Name = "Book";
                // Write a value into A1
                xlSheet.Cells[1, 1] = "Some value";
                xlSheet1.Cells[1, 1] = "Some value";                
                // Tell Excel to save your spreadsheet                
                xlWB.SaveAs(Output, Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel8, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);                
                xlWB.Close();
                xlApp.Quit();
            }
            //DataSet ds = new DataSet();
            //ds.Tables.Add(new DataTable("Comparable"));
            //ds.Tables.Add(new DataTable("Book"));
            //ds.Tables[0].Columns.Add("SAM");
            //DataRow dr = ds.Tables[0].NewRow();
            //dr[0] = "ME";
            //ds.Tables[0].Rows.Add(dr);            
            //ds.Tables[1].Columns.Add("SAM");
            //dr = ds.Tables[1].NewRow();
            //dr[0] = "ME";
            //ds.Tables[1].Rows.Add(dr);
            //CreateWorkbook(Output, ds);
            temp = File.OpenRead(Output);
            string pathFileDestination = temp.Name;
            temp.Close();
            //string pathFileDestination = Output;
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();            
            //excel.Visible = true;            
            Microsoft.Office.Interop.Excel.Workbook wbDest = excel.Workbooks.Open(pathFileDestination, 0, false, 1, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, 9, true, false, 0, true, false, false);            
            Microsoft.Office.Interop.Excel.Worksheet WorksheetDest = wbDest.Sheets[1];
            //Clear all contents in Destination workbook
            WorksheetDest.UsedRange.ClearContents();
            WorksheetDest.get_Range("A1").Value = "No Data";
            wbDest.Save();
            wbDest.Close();
            //Open the Source file
            Microsoft.Office.Interop.Excel.Workbook wbSource = excel.Workbooks.Open(ExcelPath, Type.Missing, true,Type.Missing,"", "", false, Type.Missing, Type.Missing, true, false, Type.Missing, true, Type.Missing, false);
            Microsoft.Office.Interop.Excel.Worksheet WorksheetSource = wbSource.Sheets[1];
            //Copy all range in this worksheet
            WorksheetSource.UsedRange.Copy();            
            //Open destination workbook            
            Microsoft.Office.Interop.Excel.Workbook wbDestination = excel.Workbooks.Open(pathFileDestination, 0, false, 1, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, 9, true, false, 0, true, false, false);
            Microsoft.Office.Interop.Excel.Worksheet WorksheetDestination;
            if (path.Contains("book"))
            {                
                WorksheetDestination = wbDestination.Sheets[2];
                //WorksheetSource.UsedRange.Copy(wbDestination.Sheets[2]);
            }
            else
            {
                WorksheetDestination = wbDestination.Sheets[1];
                //WorksheetSource.UsedRange.Copy(wbDestination.Sheets[1]);
            }
            //WorksheetSource = WorksheetDestination;            
            WorksheetDestination.UsedRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteColumnWidths, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, Missing.Value, Missing.Value);
            WorksheetDestination.UsedRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, Missing.Value, Missing.Value);
            wbDestination.Date1904 = true;
            wbDestination.Save();
            wbSource.Close();
            //wbDest.Save();
            //wbDest.Close();            
            wbDestination.Close();            
            excel.Quit();             
            GC.Collect();
        }
        //static void upload()
        //{
        //    String fileToUpload = @"C:\YourFile.txt";
        //    String sharePointSite = "http://yoursite.com/sites/Research/";
        //    String documentLibraryName = "Shared Documents";
        //    SPIntranet.DocumentsItem di = new SPIntranet.DocumentsItem();

        //    using (SPSite oSite = new SPSite(sharePointSite))
        //    {
        //        using (SPWeb oWeb = oSite.OpenWeb())
        //        {
        //            if (!System.IO.File.Exists(fileToUpload))
        //                throw new FileNotFoundException("File not found.", fileToUpload);

        //            SPFolder myLibrary = oWeb.Folders[documentLibraryName];

        //            // Prepare to upload
        //            Boolean replaceExistingFiles = true;
        //            String fileName = System.IO.Path.GetFileName(fileToUpload);
        //            FileStream fileStream = File.OpenRead(fileToUpload);

        //            // Upload document
        //            SPFile spfile = myLibrary.Files.Add(fileName, fileStream, replaceExistingFiles);

        //            // Commit 
        //            myLibrary.Update();
        //        }
        //    }
        //}

        static void spUpload()
        {            
            spLocation = @"\\intranet\apdops\Custom Reports\TLV\";
            if (TSDays > 27)
            {
                spLocation += "SF-All States (Monthly)";
            }
            else if (TSDays > 6)
            {
                spLocation += "SF -All States (Weekly)";
            }
            spLocation+="\\" + Convert.ToDateTime(dateFrom).Year;
            bool dExistx = Directory.Exists(spLocation);
            if(!dExistx)
                Directory.CreateDirectory(spLocation);
            spLocation += "\\" + Output_name;            
            File.Copy(Output, spLocation);
            spLocation.Replace('\\','/');
            spLocationUri += "http:" + spLocation;
        }
        static void mail(bool error,string msg="")
        {
            Console.WriteLine("Sending Mail");
            File.AppendAllText(path_pre + "Log.txt", "Sending Mail\n");
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
            service.Credentials = new WebCredentials("sp102532", "Mitchell12", "corp");
            service.AutodiscoverUrl("sambhav.patni@mitchell.com");
            //service.Url = new Uri("https://mail510ntv.mitchell.com/EWS/Exchange.asmx");            
            EmailMessage message = new EmailMessage(service);
            message.Subject = "TLV Report for SF-All States";
            if (error == true)
                message.Body = "Error Occured<br/><br/>" + msg + "<br/><br/>Thanks,\n<br/>Sambhav Patni<br/><br/>Auto Generated On " + DateTime.Now.ToString();
            else
            {
                message.Body = "Hi,\n\n<br/><br/>Please find the TLV Report “Report for SF-All States at below mentioned link:”\n<br/>";
                message.Body += "<br/><a href='" + spLocationUri + "'>" + spLocationUri + "</a><br/><br/>";
                message.Body += "The date range for the report is  " + Convert.ToDateTime(dateFrom).ToLongDateString() + " - " + Convert.ToDateTime(dateTo).Subtract(TimeSpan.FromDays(1)).ToLongDateString() + ".\n\n<br/><br/>Thanks,\n<br/>Sambhav Patni<br/><br/>Auto Generated On " + DateTime.Now.ToString();
                if(upload==false)
                    message.Attachments.AddFileAttachment(Output);
            }            
            message.ToRecipients.Add("sambhav.patni@mitchell.com");
            message.Save();

            message.SendAndSaveCopy();
        }       
    }
}
