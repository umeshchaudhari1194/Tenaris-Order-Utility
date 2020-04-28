using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Mail;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using static System.Net.WebRequestMethods;
using Attachment = System.Net.Mail.Attachment;
using File = System.IO.File;

namespace Tenaris_Order
{
    class Program
    {
        static string ConnString = ConfigurationManager.AppSettings["ConnString"];
        static string strFileLoc = ConfigurationManager.AppSettings["FilePath"];
        static string strChkFileLoc = ConfigurationManager.AppSettings["ChkFilePath"];
        static string zipPath = ConfigurationSettings.AppSettings["zipPath"];
        static string errorPath = ConfigurationSettings.AppSettings["errorPath"];
        static FileInfo[] files;
        static string MailSubject;
        static string strMaxServiceDate;
        static string strFileNameServiceDate;
        static DateTime newdate;
        static string strFilePath;
        static void Main(string[] args)
        {
            try
            {
               // string strargs = "ck";
               string strargs = args[0].ToString();
                Program p = new Program();
                StringBuilder sb = new StringBuilder();
                DataTable dt = p.GetData(strargs);
                if (dt.Rows.Count > 0)
                {

                    var maxDate = dt.Select("ServiceDate=MAX(ServiceDate)");
                    if (strargs == "order")
                    {
                        newdate = Convert.ToDateTime(maxDate[0].ItemArray[0]);
                        strFileNameServiceDate = newdate.ToString("MM-dd-yyyy", CultureInfo.InvariantCulture);
                        strFilePath = strFileLoc + ConfigurationSettings.AppSettings["FileNameOrder"] + strFileNameServiceDate + ".xlsx";
                    }
                    else if (strargs == "ck")
                    {
                        newdate = Convert.ToDateTime(maxDate[0].ItemArray[3]);
                        strFileNameServiceDate = newdate.ToString("MM-dd-yyyy", CultureInfo.InvariantCulture);
                        strFilePath = strFileLoc + ConfigurationSettings.AppSettings["FileName_CKOrder"] + strFileNameServiceDate + ".xlsx";
                    }
                    //check folders
                    p.checkFolders();

                    if (strargs == "order")
                    {
                        //create excel file
                        createExcelFile(dt, strFilePath);
                    }
                    else if (strargs == "ck")
                    {
                        //create excel file
                        createExcelFileTenaris_CK_orders(dt, strFilePath);
                    }

                    //send mail
                    SendAlert(strargs);

                    Console.WriteLine("Mail send");

                    //Delete older files which is more than 7 days older.
                    deleteOlderFiles();
                }
                else // send blank excel
                {
                    SendBlankExcel(dt, strargs);
                }


            }
            catch (Exception ex)
            {
                ErrorLog(ex);
                SendAlertError(ex);
            }
        }

        public static void SendBlankExcel(DataTable dt, string strargs)
        {

            if (strargs == "order") // for daily get next day date
            {
                //Get Next day date
                strFileNameServiceDate = DateTime.Now.AddDays(1).ToString("MM-dd-yyyy");

                strFilePath = strFileLoc + ConfigurationSettings.AppSettings["FileNameOrder"] + strFileNameServiceDate + ".xlsx";
                //create excel file
                createExcelFile(dt, strFilePath);
                SendAlert(strargs);
            }
            else if (strargs == "ck") // get today's date
            {
                strFileNameServiceDate = DateTime.Now.ToString("MM-dd-yyyy");
                strFilePath = strFileLoc + ConfigurationSettings.AppSettings["FileName_CKOrder"] + strFileNameServiceDate + ".xlsx";
                //create excel file
                createExcelFileTenaris_CK_orders(dt, strFilePath);
                SendAlert(strargs);
            }
        }
        DataTable GetData(string argsChk)
        {
            DataTable dt = new DataTable();
            SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnString"].ToString());

            if (argsChk == "order")
            {
                SqlCommand cmd = new SqlCommand("TenarisOrders_Transactions_utility", connection);
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                connection.Open();
                dt.Load(cmd.ExecuteReader());
            }
            else if (argsChk == "ck")
            {
                SqlCommand cmd = new SqlCommand("TenarisOrders_CK_Transactions_utility", connection);
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                connection.Open();
                dt.Load(cmd.ExecuteReader());
            }
            return dt;
        }

        //create csv file.
        public static void ToCSV(DataTable dtDataTable, string strFilePath)
        {
            StreamWriter sw = new StreamWriter(strFilePath, false);
            //headers  
            for (int i = 0; i < dtDataTable.Columns.Count; i++)
            {
                sw.Write(dtDataTable.Columns[i]);
                if (i < dtDataTable.Columns.Count - 1)
                {
                    sw.Write(",");
                }
            }
            sw.Write(sw.NewLine);
            foreach (DataRow dr in dtDataTable.Rows)
            {
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                    {
                        string value = dr[i].ToString();
                        if (value.Contains(','))
                        {
                            value = String.Format("\"{0}\"", value);
                            sw.Write(value);
                        }
                        else
                        {
                            sw.Write(dr[i].ToString());
                        }
                    }
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }

        public static void createExcelFile(DataTable dtDataTable, string strFilePath)
        {
            var srcPath = ConfigurationManager.AppSettings["ChkFilePath"] + @"\TenerisOrder.xlsx";
            var Destpath = strFilePath;

            Boolean FileExist = File.Exists(Destpath);
            if (FileExist)
                File.Delete(Destpath);

            File.Copy(srcPath, Destpath);

            SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(Destpath, true);
            try
            {
                SharedStringTablePart shareStringPart;
                if (spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                {
                    shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                }
                else
                {
                    shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
                }
                WorkbookPart workbookPart = spreadSheet.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == "Orders").FirstOrDefault();
                WorksheetPart worksheetPart = workbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
                if (dtDataTable.Rows.Count > 0)
                {
                    for (var i = 0; i < dtDataTable.Rows.Count; i++)
                    {
                        int row = 2 + i;

                        //---------------Service Date--------------------------

                        string strQuantity = DateTime.Parse(dtDataTable.Rows[i][0].ToString()).ToString("MM-dd-yyyy");
                        int Qindex = InsertSharedStringItem(strQuantity, shareStringPart);
                        DocumentFormat.OpenXml.Spreadsheet.Cell cell10 = InsertCellInWorksheet("A", Convert.ToUInt32(row), worksheetPart);
                        cell10.CellValue = new CellValue(Qindex.ToString());
                        cell10.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                        //---------------LOCATION--------------------------
                        string convertdate = dtDataTable.Rows[i][1].ToString();
                        int TDindex = InsertSharedStringItem(convertdate, shareStringPart);
                        DocumentFormat.OpenXml.Spreadsheet.Cell cell1 = InsertCellInWorksheet("B", Convert.ToUInt32(row), worksheetPart);
                        cell1.CellValue = new CellValue(TDindex.ToString());
                        cell1.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                        //---------------USER NAME--------------------------

                        string strUserName = dtDataTable.Rows[i][2].ToString();
                        int UNindex = InsertSharedStringItem(strUserName, shareStringPart);
                        DocumentFormat.OpenXml.Spreadsheet.Cell cell2 = InsertCellInWorksheet("C", Convert.ToUInt32(row), worksheetPart);
                        cell2.CellValue = new CellValue(UNindex.ToString());
                        cell2.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                        //---------------MARKET CARD NUMBER--------------------------
                        string strCardNumber = dtDataTable.Rows[i][3].ToString();
                        int CNindex = InsertSharedStringItem(strCardNumber, shareStringPart);
                        DocumentFormat.OpenXml.Spreadsheet.Cell cell3 = InsertCellInWorksheet("D", Convert.ToUInt32(row), worksheetPart);
                        cell3.CellValue = new CellValue(CNindex.ToString());
                        cell3.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                        //---------------ITEM--------------------------
                        string strServiceDate = dtDataTable.Rows[i][4].ToString();
                        int SDindex = InsertSharedStringItem(strServiceDate, shareStringPart);
                        DocumentFormat.OpenXml.Spreadsheet.Cell cell4 = InsertCellInWorksheet("E", Convert.ToUInt32(row), worksheetPart);
                        cell4.CellValue = new CellValue(SDindex.ToString());
                        cell4.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                        //---------------DESSERT--------------------------
                        string strDeliveryPoint = dtDataTable.Rows[i][5].ToString();
                        int DPindex = InsertSharedStringItem(strDeliveryPoint, shareStringPart);
                        DocumentFormat.OpenXml.Spreadsheet.Cell cell5 = InsertCellInWorksheet("F", Convert.ToUInt32(row), worksheetPart);
                        cell5.CellValue = new CellValue(DPindex.ToString());
                        cell5.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                        //---------------FULL PRICE--------------------------
                        string strOrderType = dtDataTable.Rows[i][6].ToString();
                        int OTindex = InsertSharedStringItem(strOrderType, shareStringPart);
                        DocumentFormat.OpenXml.Spreadsheet.Cell cell6 = InsertCellInWorksheet("G", Convert.ToUInt32(row), worksheetPart);
                        cell6.CellValue = new CellValue(OTindex.ToString());
                        cell6.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                        //---------------Employee--------------------------
                        string strItemName = dtDataTable.Rows[i][7].ToString();
                        int ITindex = InsertSharedStringItem(strItemName, shareStringPart);
                        DocumentFormat.OpenXml.Spreadsheet.Cell cell7 = InsertCellInWorksheet("H", Convert.ToUInt32(row), worksheetPart);
                        cell7.CellValue = new CellValue(ITindex.ToString());
                        cell7.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                        //---------------SALES TAX--------------------------
                        string strSalesTax = dtDataTable.Rows[i][8].ToString();
                        int STndex = InsertSharedStringItem(strSalesTax, shareStringPart);
                        DocumentFormat.OpenXml.Spreadsheet.Cell cell11 = InsertCellInWorksheet("I", Convert.ToUInt32(row), worksheetPart);
                        cell11.CellValue = new CellValue(STndex.ToString());
                        cell11.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                        //---------------TENARIS--------------------------
                        string strDessert = dtDataTable.Rows[i][9].ToString();
                        int Dindex = InsertSharedStringItem(strDessert, shareStringPart);
                        DocumentFormat.OpenXml.Spreadsheet.Cell cell8 = InsertCellInWorksheet("J", Convert.ToUInt32(row), worksheetPart);
                        cell8.CellValue = new CellValue(Dindex.ToString());
                        cell8.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                        //---------------AFS2T--------------------------
                        string strItemCost = dtDataTable.Rows[i][10].ToString();
                        int ICindex = InsertSharedStringItem(strItemCost, shareStringPart);
                        DocumentFormat.OpenXml.Spreadsheet.Cell cell9 = InsertCellInWorksheet("K", Convert.ToUInt32(row), worksheetPart);
                        cell9.CellValue = new CellValue(ICindex.ToString());
                        cell9.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                    }
                }
                else // blank excel sheet
                {
                    //---------------No Orders--------------------------

                    string strOrder = "No Orders";
                    int Qindex = InsertSharedStringItem(strOrder, shareStringPart);
                    DocumentFormat.OpenXml.Spreadsheet.Cell cell1 = InsertCellInWorksheet("A", 2, worksheetPart);
                    cell1.CellValue = new CellValue(Qindex.ToString());
                    cell1.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                }
                //******************************************************************************
                worksheetPart.Worksheet.Save();
                // close excel file
                spreadSheet.Close();
            }
            catch (Exception ex) { }
            finally
            {
                foreach (Process process in Process.GetProcessesByName("Excel"))
                    process.Kill();
            }
        }

        public static void createExcelFileTenaris_CK_orders(DataTable dtDataTable, string strFilePath)
        {
            var srcPath = ConfigurationManager.AppSettings["ChkFilePath"] + @"\Teneris_CK_Order.xlsx";
            var Destpath = strFilePath;

            Boolean FileExist = File.Exists(Destpath);
            if (FileExist)
                File.Delete(Destpath);

            File.Copy(srcPath, Destpath);

            SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(Destpath, true);
            try
            {
                SharedStringTablePart shareStringPart;
                if (spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                {
                    shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                }
                else
                {
                    shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
                }
                WorkbookPart workbookPart = spreadSheet.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == "Orders").FirstOrDefault();
                WorksheetPart worksheetPart = workbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
                if (dtDataTable.Rows.Count > 0)
                {
                    for (var i = 0; i < dtDataTable.Rows.Count; i++)
                    {
                        int row = 2 + i;

                        //---------------KISOK--------------------------
                        string convertdate = dtDataTable.Rows[i][0].ToString();
                        int TDindex = InsertSharedStringItem(convertdate, shareStringPart);
                        DocumentFormat.OpenXml.Spreadsheet.Cell cell1 = InsertCellInWorksheet("A", Convert.ToUInt32(row), worksheetPart);
                        cell1.CellValue = new CellValue(TDindex.ToString());
                        cell1.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                        //---------------MARKET CARD--------------------------

                        string strUserName = dtDataTable.Rows[i][1].ToString();
                        int UNindex = InsertSharedStringItem(strUserName, shareStringPart);
                        DocumentFormat.OpenXml.Spreadsheet.Cell cell2 = InsertCellInWorksheet("B", Convert.ToUInt32(row), worksheetPart);
                        cell2.CellValue = new CellValue(UNindex.ToString());
                        cell2.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                        //---------------USER NAME--------------------------
                        string strCardNumber = dtDataTable.Rows[i][2].ToString();
                        int CNindex = InsertSharedStringItem(strCardNumber, shareStringPart);
                        DocumentFormat.OpenXml.Spreadsheet.Cell cell3 = InsertCellInWorksheet("C", Convert.ToUInt32(row), worksheetPart);
                        cell3.CellValue = new CellValue(CNindex.ToString());
                        cell3.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                        //---------------SERVICE DATE--------------------------

                        string strQuantity = DateTime.Parse(dtDataTable.Rows[i][3].ToString()).ToString("MM-dd-yyyy");
                        int Qindex = InsertSharedStringItem(strQuantity, shareStringPart);
                        DocumentFormat.OpenXml.Spreadsheet.Cell cell10 = InsertCellInWorksheet("D", Convert.ToUInt32(row), worksheetPart);
                        cell10.CellValue = new CellValue(Qindex.ToString());
                        cell10.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                        //---------------Employee--------------------------
                        string strItemName = dtDataTable.Rows[i][4].ToString();
                        int ITindex = InsertSharedStringItem(strItemName, shareStringPart);
                        DocumentFormat.OpenXml.Spreadsheet.Cell cell7 = InsertCellInWorksheet("E", Convert.ToUInt32(row), worksheetPart);
                        cell7.CellValue = new CellValue(ITindex.ToString());
                        cell7.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                        //---------------SALES TAX--------------------------
                        string strSalesTax = dtDataTable.Rows[i][5].ToString();
                        int STndex = InsertSharedStringItem(strSalesTax, shareStringPart);
                        DocumentFormat.OpenXml.Spreadsheet.Cell cell11 = InsertCellInWorksheet("F", Convert.ToUInt32(row), worksheetPart);
                        cell11.CellValue = new CellValue(STndex.ToString());
                        cell11.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                        //---------------Total--------------------------
                        string strDessert = dtDataTable.Rows[i][6].ToString();
                        int Dindex = InsertSharedStringItem(strDessert, shareStringPart);
                        DocumentFormat.OpenXml.Spreadsheet.Cell cell8 = InsertCellInWorksheet("G", Convert.ToUInt32(row), worksheetPart);
                        cell8.CellValue = new CellValue(Dindex.ToString());
                        cell8.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    }
                }
                else
                {
                    //---------------No Orders--------------------------

                    string strOrder = "No Orders";
                    int Qindex = InsertSharedStringItem(strOrder, shareStringPart);
                    DocumentFormat.OpenXml.Spreadsheet.Cell cell1 = InsertCellInWorksheet("A", 2, worksheetPart);
                    cell1.CellValue = new CellValue(Qindex.ToString());
                    cell1.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                }

                //******************************************************************************
                worksheetPart.Worksheet.Save();
                // close excel file
                spreadSheet.Close();
            }
            catch (Exception ex) { }
            finally
            {
                foreach (Process process in Process.GetProcessesByName("Excel"))
                    process.Kill();
            }
        }

        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {

            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }
            int i = 0;
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }
        private static DocumentFormat.OpenXml.Spreadsheet.Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;
            DocumentFormat.OpenXml.Spreadsheet.Row lastRow = sheetData.Elements<DocumentFormat.OpenXml.Spreadsheet.Row>().LastOrDefault();
            // If the worksheet does not contain a row with the specified row index, insert one.
            DocumentFormat.OpenXml.Spreadsheet.Row row;
            if (sheetData.Elements<DocumentFormat.OpenXml.Spreadsheet.Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<DocumentFormat.OpenXml.Spreadsheet.Row>().Where(r => r.RowIndex == rowIndex).First();

                //set auto height -- don't know how this line is worked
                sheetData.InsertAfter(new DocumentFormat.OpenXml.Spreadsheet.Row() { RowIndex = (lastRow.RowIndex + 1) }, lastRow);
            }
            else
            {
                row = new DocumentFormat.OpenXml.Spreadsheet.Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                DocumentFormat.OpenXml.Spreadsheet.Cell refCell = null;
                foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                DocumentFormat.OpenXml.Spreadsheet.Cell newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);
                worksheet.Save();
                return newCell;
            }
        }

        //Send Mail
        public static void SendAlert(string argsChk)
        {
            try
            {
                
                    if (argsChk == "order") // order excel
                    {
                        TransactionOrdersSendAlert();
                    }
                    else if (argsChk == "ck") // CK-Order excel
                    {
                        Transaction_CK_OrdersSendAlert();
                    }
            }
            catch (Exception e)
            {
                ErrorLog(e);
            }

        }


        public static void TransactionOrdersSendAlert()
        {
            try
            {
                var directory = new DirectoryInfo(zipPath);
                var myFile = (from f in directory.GetFiles()
                              orderby f.LastWriteTime descending
                              select f).First();

                ExchangeService service = new ExchangeService();
                string from = ConfigurationSettings.AppSettings["FromAddress"];
                string frompass = ConfigurationSettings.AppSettings["FromPassword"];
                service.Credentials = new NetworkCredential(from, frompass);

                service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
                EmailMessage message = new Microsoft.Exchange.WebServices.Data.EmailMessage(service);
                message.Subject = ConfigurationSettings.AppSettings["MailSubjectOrder"] + strFileNameServiceDate + "";
                message.Body = ConfigurationSettings.AppSettings["MailBodyOrder"] + strFileNameServiceDate + " excel";
                var path = System.IO.Path.Combine(directory + myFile.ToString());

                message.Attachments.AddFileAttachment(path);

                foreach (var address in ConfigurationSettings.AppSettings["ToAddressDaily"].Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries))
                {
                    message.ToRecipients.Add(address);
                }


                message.Send();

            }
            catch (Exception e)
            {
                ErrorLog(e);
                SendAlertError(e);
            }

        }

        public static void Transaction_CK_OrdersSendAlert()
        {
            try
            {
                var directory = new DirectoryInfo(zipPath);
                var myFile = (from f in directory.GetFiles()
                              orderby f.LastWriteTime descending
                              select f).First();

                ExchangeService service = new ExchangeService();
                string from = ConfigurationSettings.AppSettings["FromAddress"];
                string frompass = ConfigurationSettings.AppSettings["FromPassword"];
                service.Credentials = new NetworkCredential(from, frompass);

                service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
                EmailMessage message = new Microsoft.Exchange.WebServices.Data.EmailMessage(service);
                message.Subject = ConfigurationSettings.AppSettings["MailSubject_CKOrder"] + strFileNameServiceDate + "";
                message.Body = ConfigurationSettings.AppSettings["MailBody_CKOrder"] + strFileNameServiceDate + " excel";
                var path = System.IO.Path.Combine(directory + myFile.ToString());

                message.Attachments.AddFileAttachment(path);

                foreach (var address in ConfigurationSettings.AppSettings["ToAddressCK"].Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries))
                {
                    message.ToRecipients.Add(address);
                }
                message.Send();
            }
            catch (Exception ex)
            {
                SendAlertError(ex);
            }
        }

        public static void SendAlertError(Exception ex)
        {
            string strBody = "Message :" + ex.Message + "<br/>" + Environment.NewLine + "StackTrace :" + ex.StackTrace +
                   "" + Environment.NewLine + "Date :" + DateTime.Now.ToString();
            ExchangeService service = new ExchangeService();
            string from = ConfigurationSettings.AppSettings["FromAddress"];
            string frompass = ConfigurationSettings.AppSettings["FromPassword"];
            service.Credentials = new NetworkCredential(from, frompass);

            service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
            EmailMessage message = new Microsoft.Exchange.WebServices.Data.EmailMessage(service);
            message.Subject = "Error in tenaris order utility : " + ex.Message;
            message.Body = strBody;
            foreach (var address in ConfigurationSettings.AppSettings["ToAddressErrorNotification"].Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries))
            {
                message.ToRecipients.Add(address);
            }
            message.Send();
        }



        public static void ErrorLog(Exception ex)
        {
            string filePath = errorPath + @"\Error.txt";

            using (StreamWriter writer = new StreamWriter(filePath, true))
            {
                writer.WriteLine("Message :" + ex.Message == "Access to the path 'D:\\' is denied." ? "Email attachment Zip File Not Found." : ex.Message + "<br/>" + Environment.NewLine + "StackTrace :" + ex.StackTrace +
                   "" + Environment.NewLine + "Date :" + DateTime.Now.ToString());
                writer.WriteLine(Environment.NewLine + "-----------------------------------------------------------------------------" + Environment.NewLine);
            }
        }

        //check folder is exists or not
        public void checkFolders()
        {

            //check root folder 
            string path = strChkFileLoc;
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            //  check subfolders
            string pathsub = strChkFileLoc + "\\OrderArcheive";
            if (!Directory.Exists(pathsub))
            {
                Directory.CreateDirectory(pathsub);
            }
            //check 2nd subfolder
            string pathsubnew = strChkFileLoc + "\\OrderBkp";
            if (!Directory.Exists(pathsubnew))
            {
                Directory.CreateDirectory(pathsubnew);
            }

            string txtPath = errorPath + @"\Error.txt";
            if (!System.IO.File.Exists(txtPath))
            {
                System.IO.File.Create(txtPath);
            }
        }

        //Delete older files which is more than 7 days older.
        public static void deleteOlderFiles()
        {
            string[] files = Directory.GetFiles(ConfigurationSettings.AppSettings["FilePath"]);

            foreach (string file in files)
            {
                FileInfo fi = new FileInfo(file);
                if (fi.LastAccessTime < DateTime.Now.AddDays(-7))
                    fi.Delete();
            }
        }

        
       

        
    }

}
