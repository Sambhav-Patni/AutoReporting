//using ExcelLibrary.SpreadSheet;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OracleClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;
using Excel = Microsoft.Office.Interop.Excel;

namespace Auto_Reportin_SVV
{
    class Program
    {
        static string[] config;
        static string dateTo = DateTime.Now.Date.ToString("dd-MMM-yyyy");//"1-JAN-2015";
        //static string Mon;
        static string[] Months = { "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" };
        static string dateFrom = DateTime.Now.Date.ToString("dd-MMM-yyyy");//"SAM-Def";
        //static string dateTo = "16-JAN-2015";        
        //static string dateFrom = DateTime.Now.Date.Subtract(TimeSpan.FromDays(TSDays)).ToString("dd-MMM-yyyy");//"16-JAN-2015";
        static string path_pre = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\";
        static string path = path_pre + "SVV.xlsx";
        static string Output = path_pre + "CycleTime(SVV) bi-weekly_PlaceHolder.xlsx";
        static string Output_name = "CycleTime(SVV) bi-weekly_PlaceHolder.xlsx";
        static string spLocation = "";
        static string spLocationUri = "";
        static bool upload = true;
        static string NmapPass = "XXX";
        static string IFNmapPass = "XXX";
        static string MitchellPass = "XXX";

        static void Main(string[] args)
        {
            try
            {
                try
                {
                    dateFrom = args[0];
                    dateTo = args[1];
                }
                catch { }
                if (DateTime.Now.Day == 1)
                {
                    dateFrom = DateTime.Now.Subtract(TimeSpan.FromDays(15)).ToString("16-MMM-yyyy");
                }
                else if (DateTime.Now.Day == 16)
                {
                    dateFrom = DateTime.Now.ToString("1-MMM-yyyy");
                }
                config = File.ReadAllLines(path_pre + "DBConfig.sam");
                NmapPass = config[0].Substring(config[0].LastIndexOf(':') + 1).Trim();
                IFNmapPass = config[1].Substring(config[1].LastIndexOf(':') + 1).Trim();
                MitchellPass = config[2].Substring(config[2].LastIndexOf(':') + 1).Trim();
                //dateFrom = "16-" + Mon;// +DateTime.Now.Year;
                //string temp1 = Convert.ToDateTime(dateFrom).ToShortDateString();
                //string temp2 = Convert.ToDateTime(dateTo).ToShortDateString();
                string NmapQuery = "select b.co_cd \"CO CD\", a.valreq_id \"VALREQ ID\", b.claim_nbr||'-'||b.exposure_nbr \"CLAIM EXP NUMBER\",c.svv_type, b.created_by \"CREATED BY\",created_in_esrt \"ARRIVED IN ESRT\", assigned_timestamp \"ASSIGNED TO RESEARCHER\",ROUND(assigned_timestamp - created_in_esrt,7) \"ARRIVING TO ASSIGNED\",closed_timestamp \"SVV QUOTES SUBMITTED\", ROUND(closed_timestamp - assigned_timestamp,7) \"ASSIGNED TO SVV SUBMITTED\", ROUND(closed_timestamp - created_in_esrt,7) \"TOTAL TIME IN ESRT\" from tlv_exception_service_asgnmt a, tlv_valuation_request b, tlv.tlv_svv_type c where assignment_type = 'SVP' and a.valreq_id = b.id and b.svv_type_id=c.svv_type_id and b.co_cd not in ('X9', 'Y6', 'Z1', 'Z4', 'IF') and a.created_dt >= trunc(TO_DATE('" + Convert.ToDateTime(dateFrom).ToShortDateString() + "', 'MM-DD-YYYY'))  and a.created_dt < trunc(TO_DATE('" + Convert.ToDateTime(dateTo).ToShortDateString() + "', 'MM-DD-YYYY')) and a.dispncd_id = 36";
                string IFNmapQuery = "select b.co_cd \"CO CD\", a.valreq_id \"VALREQ ID\", b.claim_nbr||'-'||b.exposure_nbr \"CLAIM EXP NUMBER\",c.svv_type, b.created_by \"CREATED BY\",created_in_esrt \"ARRIVED IN ESRT\", assigned_timestamp \"ASSIGNED TO RESEARCHER\",ROUND(assigned_timestamp - created_in_esrt,7) \"ARRIVING TO ASSIGNED\",closed_timestamp \"SVV QUOTES SUBMITTED\", ROUND(closed_timestamp - assigned_timestamp,7) \"ASSIGNED TO SVV SUBMITTED\", ROUND(closed_timestamp - created_in_esrt,7) \"TOTAL TIME IN ESRT\" from tlv_exception_service_asgnmt a, tlv_valuation_request b, tlv.tlv_svv_type c where assignment_type = 'SVP' and a.valreq_id = b.id and b.svv_type_id=c.svv_type_id and b.co_cd = 'IF' and a.created_dt >= trunc(TO_DATE('" + Convert.ToDateTime(dateFrom).ToShortDateString() + "', 'MM-DD-YYYY'))  and a.created_dt < trunc(TO_DATE('" + Convert.ToDateTime(dateTo).ToShortDateString() + "', 'MM-DD-YYYY')) and a.dispncd_id = 36";
                Output = Output.Replace("PlaceHolder", Convert.ToDateTime(dateFrom).Day + "-" + (Convert.ToDateTime(dateTo).AddDays(-1).Day) + "_" + Months[Convert.ToDateTime(dateFrom).Month - 1] + "_" + Convert.ToDateTime(dateTo).Year);
                Output_name = Output_name.Replace("PlaceHolder", Convert.ToDateTime(dateFrom).Day + "-" + (Convert.ToDateTime(dateTo).AddDays(-1).Day) + "_" + Months[Convert.ToDateTime(dateFrom).Month - 1] + "_" + Convert.ToDateTime(dateTo).Year);
                Console.Write("\n\n\t\t\tHello, Lets Get Started with CycleTime Reporting\n");
                Console.WriteLine("Getting Data From Nmap");

                Retrive_Merge_Data(NmapQuery, IFNmapQuery);
                File.Copy(path, path.Replace(".xlsx", "_backup.xlsx"), true);
                Console.WriteLine("Running VBA");
                vba();
                File.Copy(path, Output, true);
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
        public static void Retrive_Merge_Data(string query1, string query2)
        {
            DataSet ds = new DataSet();
            DataSet ds1 = new DataSet();
            //try
            //{
                using (OracleConnection connection = new OracleConnection("User Id=sp102532;Password="+NmapPass+";Data Source=nmap.mitchell.com"))
                {                    
                    connection.Open();
                    OracleCommand command;
                    OracleDataAdapter adaptor;
                    DataTable dtschema = connection.GetSchema();
                    try
                    {                        
                        command = new OracleCommand(query1, connection);
                        adaptor = new OracleDataAdapter(command);
                        Console.WriteLine("\nCommand Set: " + command.CommandText);
                        adaptor.Fill(ds);
                        Console.WriteLine("\nGot Data From " + connection.DataSource);
                    }
                    catch (IndexOutOfRangeException)
                    {
                        
                    }                    
                }
                using (OracleConnection connection = new OracleConnection("User Id=sp102532;Password="+IFNmapPass+";Data Source=ifnmap.mitchell.com"))
                {                    
                    connection.Open();
                    OracleCommand command;
                    OracleDataAdapter adaptor;
                    DataTable dtschema = connection.GetSchema();
                    try
                    {
                        command = new OracleCommand(query2, connection);
                        adaptor = new OracleDataAdapter(command);
                        Console.WriteLine("\nCommand Set: " + command.CommandText);
                        adaptor.Fill(ds1);
                        Console.WriteLine("\nGot Data From " + connection.DataSource);
                    }
                    catch (IndexOutOfRangeException)
                    {

                    }
                }
                ds.Merge(ds1);
                CreateWorkbook(path, ds);
            //}            
            //catch (Exception ex)
            //{
            //    //MessageBox.Show("Something Went Wrong...\n\n Contact: Sambhav Patni\n At: Sambhav.patni@infogain.com\n\n" + ex.Message, "Error Occured");
            //    //File.WriteAllText(path_pre + "log.txt", "Error Log: \n" + ex.Message + "\n" + ex.StackTrace);
            //    Console.WriteLine("\nException: " + ex.Message);
            //}
        }
        public static void vba()
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;
            FileStream temp = File.OpenRead(path);
            string workbookPath = temp.Name;
            temp.Close();
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(workbookPath,
                0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "",
                true, false, 0, true, false, false);
            var newStandardModule = excelWorkbook.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
            var codeModule = newStandardModule.CodeModule;
            var lineNum = codeModule.CountOfLines + 1;
            var codeText = File.ReadAllText(path_pre + "ProcessSVV.vb");
            codeModule.InsertLines(lineNum, codeText);
            //excelWorkbook.Save();
            Console.WriteLine("macro_CreateSVV");
            var macro = string.Format("{0}!{1}.{2}", excelWorkbook.Name, newStandardModule.Name, "CreateSVV");
            excelApp.Run(macro);            
            excelApp.Visible = true;
            codeModule.DeleteLines(1, codeModule.CountOfLines);
            excelWorkbook.Save();            
            excelWorkbook.Close();
            excelApp.Quit();
        }
        //public static void CreateWorkbook_ExcelLibrary(String filePath, DataSet dataset)
        //{            
        //    CellFormat cfDate = new CellFormat(CellFormatType.Custom, "m/d/yyyy h:mm:ss AM/PM");
        //    CellFormat cfText = new CellFormat(CellFormatType.Text, "@");
        //    Console.Write("Writing Excel");
        //    if (dataset.Tables.Count == 0)
        //        throw new ArgumentException("DataSet needs to have at least one DataTable", "dataset");

        //    Workbook workbook = new Workbook();
        //    foreach (DataTable dt in dataset.Tables)
        //    {
        //        Worksheet worksheet = new Worksheet(dt.TableName);
        //        for (int i = 0; i < dt.Columns.Count; i++)
        //        {
        //            worksheet.Cells[0, i] = new Cell(dt.Columns[i].ColumnName);
        //            for (int j = 0; j < dt.Rows.Count; j++)
        //            {
        //                try
        //                {
        //                    //Convert.ToDateTime(dt.Rows[j][i]);
        //                    worksheet.Cells[j + 1, i] = new Cell(dt.Rows[j][i] == DBNull.Value ? string.Empty : dt.Rows[j][i].ToString(), cfText);
        //                }                        
        //                catch (Exception)
        //                {
        //                    worksheet.Cells[j + 1, i] = new Cell(dt.Rows[j][i] == DBNull.Value ? string.Empty : dt.Rows[j][i]);
        //                }
        //                //catch(InvalidCastException)
        //                //{
        //                //    try
        //                //    {
        //                //        Convert.ToDecimal(dt.Rows[j][i]);
        //                //        worksheet.Cells[j + 1, i] = new Cell(dt.Rows[j][i], new CellFormat(CellFormatType.Number, "0.0000000"));
        //                //    }
        //                //    catch (InvalidCastException)
        //                //    {
        //                //        worksheet.Cells[j + 1, i] = new Cell(dt.Rows[j][i] == DBNull.Value ? string.Empty : dt.Rows[j][i]);
        //                //    }
        //                //}
        //            }
        //        }
        //        workbook.Worksheets.Add(worksheet);
        //    }
        //    workbook.Save(filePath);
        //}
        public static void CreateWorkbook(String filePath, DataSet dataset)
        {            
            ExportToExcel(dataset.Tables[0],filePath);
        }
        // Export DataTable into an excel file with field names in the header line
        // - Save excel file without ever making it visible if filepath is given
        // - Don't save excel file, just make it visible if no filepath is given
        public static void ExportToExcel(DataTable Tbl, string ExcelFilePath = null)
        {
            Console.WriteLine("Saving Excel File....");
            try
            {
                FileStream temp = File.Open(path,FileMode.Create);
                string workbookPath = temp.Name;
                temp.Close();                

                if (Tbl == null || Tbl.Columns.Count == 0)
                    throw new Exception("ExportToExcel: Null or empty input table!\n");

                // load excel, and create a new workbook
                Excel.Application excelApp = new Excel.Application();
                excelApp.AlertBeforeOverwriting = false;
                excelApp.Workbooks.Add();

                // single worksheet
                excelApp.Sheets[1].Delete();
                excelApp.Sheets[2].Delete();
                Excel._Worksheet workSheet = excelApp.ActiveSheet;

                // column headings
                for (int i = 0; i < Tbl.Columns.Count; i++)
                {
                    workSheet.Cells[1, (i + 1)] = Tbl.Columns[i].ColumnName;
                }
                int countDivision = Tbl.Rows.Count / 100;
                int tmp = 0;
                Console.Write("Progress: ");
                // rows
                for (int i = 0; i < Tbl.Rows.Count; i++)
                {                    
                    // to do: format datetime values before printing
                    for (int j = 0; j < Tbl.Columns.Count; j++)
                    {
                        workSheet.Cells[(i + 2), (j + 1)] = Tbl.Rows[i][j];
                    }
                    tmp++;
                    if (tmp > countDivision)
                    {
                        Console.Write("-");
                        tmp = 0;
                    }
                }

                // check fielpath
                if (ExcelFilePath != null && ExcelFilePath != "")
                {
                    try
                    {                        
                        workSheet.Name = "Table";
                        File.Delete(workbookPath);    
                        workSheet.SaveAs(workbookPath);                        
                        excelApp.Quit();
                        //MessageBox.Show("Excel file saved!");
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n"
                            + ex.Message);
                    }
                }
                else    // no filepath is given
                {
                    excelApp.Visible = true;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("ExportToExcel: \n" + ex.Message);
            }
        }

        static void spUpload()
        {
            spLocation = @"\\intranet\apdops\Custom Reports\TLV\";

            spLocation += "Cycle time (SVV) biweekly report";
            
            spLocation += "\\" + Convert.ToDateTime(dateFrom).Year;
            bool dExistx = Directory.Exists(spLocation);
            if (!dExistx)
                Directory.CreateDirectory(spLocation);
            spLocation += "\\" + Output_name;
            File.Copy(Output, spLocation);
            spLocation.Replace('\\', '/');
            spLocationUri += "http:" + spLocation;
        }
        static void mail(bool error, string msg = "")
        {
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
            service.Credentials = new WebCredentials("sp102532", MitchellPass, "corp");
            service.AutodiscoverUrl("sambhav.patni@mitchell.com");
            //service.Url = new Uri("https://mail510ntv.mitchell.com/EWS/Exchange.asmx");            
            EmailMessage message = new EmailMessage(service);
            message.Subject = "TLV bi-Weekly Report for CycleTime(SVV)";
            if (error == true)
                message.Body = "Error Occured<br/><br/>" + msg + "<br/><br/>Thanks,\n<br/>Sambhav Patni<br/><br/>Auto Generated On " + DateTime.Now.ToString();
            else
            {
                message.Body = "Hi,\n\n<br/><br/>Please find the bi-Weekly TLV Report “Report for CycleTime(SVV)” at below mentioned link:\n<br/>";
                message.Body += "<br/><a href='" + spLocationUri + "'>" + spLocationUri + "</a><br/>";
                message.Body+="The date range for the report is  " + Convert.ToDateTime(dateFrom).ToLongDateString() + " - " + Convert.ToDateTime(dateTo).Subtract(TimeSpan.FromDays(1)).ToLongDateString() + ".\n\n<br/><br/>Thanks,\n<br/>Sambhav Patni<br/><br/>Auto Generated On " + DateTime.Now.ToString();
                if(upload==!true)
                    message.Attachments.AddFileAttachment(Output);
            }
            message.ToRecipients.Add("sambhav.patni@mitchell.com");
            message.Save();

            message.SendAndSaveCopy();
        }       
    }
}
