using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace InstancePay
{



        //       If Left(wksTempSheet.Cells(lngRowCount, lngColCount).Value, 2) <> "0." And Left(wksTempSheet.Cells(lngRowCount, lngColCount).Value, 1) = 0 And Len(wksTempSheet.Cells(lngRowCount, lngColCount).Value) > 1 Then
        //      lngColCount = lngColCount + 1
        //    ElseIf InStr(1, wksTempSheet.Cells(lngRowCount, lngColCount).Value, "/") Then
        //      lngColCount = lngColCount + 1
        //    ElseIf InStr(1, wksTempSheet.Cells(lngRowCount, lngColCount).Value, "-") Then
        //        lngColCount = lngColCount + 1
        //    Else
        //        If IsNumeric(wksTempSheet.Cells(lngRowCount, lngColCount).Value) And InStr(1, wksTempSheet.Cells(lngRowCount, lngColCount).Value, "$") Then
        //            If wksTempSheet.Cells(lngRowCount, lngColCount).Value = "$0" Then
        //                    wksTempSheet.Cells(lngRowCount, lngColCount).NumberFormat = "$0"
        //            Else
        //                If InStr(1, wksTempSheet.Cells(lngRowCount, lngColCount).Value, ".") Then
        //                    wksTempSheet.Cells(lngRowCount, lngColCount).NumberFormat = "$#,##0.00"
        //                Else
        //                    wksTempSheet.Cells(lngRowCount, lngColCount).NumberFormat = "$#,##0"
        //                End If
        //            End If
        //        ElseIf IsNumeric(wksTempSheet.Cells(lngRowCount, lngColCount).Value) And InStr(1, wksTempSheet.Cells(lngRowCount, lngColCount).Value, ",") And Not InStr(1, wksTempSheet.Cells(lngRowCount, lngColCount).Value, "$") Then
        //               If InStr(1, wksTempSheet.Cells(lngRowCount, lngColCount).Value, ".") Then
        //                   wksTempSheet.Cells(lngRowCount, lngColCount).NumberFormat = "#,##0.00"
        //               Else
        //                   wksTempSheet.Cells(lngRowCount, lngColCount).NumberFormat = "#,##0"
        //               End If
        //        ElseIf IsNumeric(wksTempSheet.Cells(lngRowCount, lngColCount).Value) Then
        //            wksTempSheet.Cells(lngRowCount, lngColCount).NumberFormat = "0"
        //        End If
        //            wksTempSheet.Cells(lngRowCount, lngColCount).Value = wksTempSheet.Cells(lngRowCount, lngColCount).Value
        //            lngColCount = lngColCount + 1
        //    End If
        //Wend


    class Program
    {

        public DataTable BindDatatable()
        {
            DataTable dt = new DataTable();

            DataColumn Ids = new DataColumn("Ids");

            DataColumn Name = new DataColumn("Name");
            DataColumn ValueWithComma = new DataColumn("ValueWithComma");


            DataColumn ValueWithDoller = new DataColumn("ValueWithDoller");
            dt.Columns.Add(Ids);
            dt.Columns.Add(Name);
            dt.Columns.Add(ValueWithComma);
            dt.Columns.Add(ValueWithDoller);
            Random rr=new Random();
            for (int i = 0; i < 6500; i++)
            {

                DataRow rows = dt.NewRow();
                rows[Ids] = i;
                rows[Name] = "pasfjawf asd asdv asfoasdfn asdfas asdf'aosf;kasdfnasdf asdnas;dfa asdvn;fkasd asdfmasdflk";
                rows[ValueWithComma] ="30000";

                rows[ValueWithDoller] =   rr.Next(100, 656565);
                dt.Rows.Add(rows);
                
            }
            return dt;
        }

        static void Main(string[] args)
        {

            Program pp = new Program();
            int ColumnWidth = 10;
            DataTable ddd = pp.BindDatatable();
            var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Style NumberFormat");

            var a = 1;

            var b = 2;

            var c = 3;

            var d = 4;
            var ro = 1;

            for (int i = 1; i < ddd.Rows.Count-1; i++)
            {
                int temp = i;
                ws.Cell(temp, a).Value = ddd.Rows[i-1][ddd.Columns[0].ColumnName].ToString();
                //ws.Cell(temp, a).Value = "123456.789";
               
                //ws.Cell(temp, a).Style.NumberFormat.Format = "$ #,##0.00";
                       
                //ws.Cell(temp, b).Value = "12,345";
                ws.Cell(temp, b).Value = ddd.Rows[i-1][ddd.Columns[1].ColumnName].ToString();
                ws.Cell(temp, b).Style.NumberFormat.Format = "0000";
             
                // Using++i OpenXML's predefined formats
                //ws.Cell(temp, c).Value = "12333";

                ws.Cell(temp, c).Value = ddd.Rows[i-1][ddd.Columns[2].ColumnName].ToString();
                //ws.Cell(temp, c).Style.NumberFormat.Format = " #,##0.00";

               // ws.Cell(temp, c).DataType = XLDataType.Number;
                ws.Cell(temp, c).Style.NumberFormat.Format = "#,##0";
               // ws.Cell(temp, c).Style.NumberFormat.SetNumberFormatId(43);

                ws.Cell(temp, d).Value = ddd.Rows[i-1][ddd.Columns[3].ColumnName].ToString();
                ws.Cell(temp, d).Style.NumberFormat.Format = "$ ###0.00";
              
            }

            ws.Column(a).AdjustToContents().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

            ws.Column(a).AdjustToContents().Width = ColumnWidth;

            ws.Column(b).AdjustToContents().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            ws.Column(b).AdjustToContents().Width = ColumnWidth;

            ws.Column(c).AdjustToContents().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

            ws.Column(d).AdjustToContents().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right; 
            // Using a custom format
          
           

          
            workbook.SaveAs("StylesNumberFormat.xlsx");
            //Program pp = new Program();
            //Console.WriteLine("Send Data Pending Scheduler: " + DateTime.Now.ToString());
            //CustomLogs.LogWriter _Log = new CustomLogs.LogWriter("Start BackUp Scheduler: " + DateTime.Now.ToString());

            //try
            //{
            //    SqlParameter[] p = new SqlParameter[0];
            //    System.Data.DataTable Data = DAL.sqlDML.GetRecords("GetPendingDatatm_aeps_bc", p, System.Data.CommandType.StoredProcedure);
            //    pp.Bind(Data);
            //    Console.WriteLine("Complete BackUp Scheduler: " + DateTime.Now.ToString());

            //    //CustomLogs.LogWriter Complete = new CustomLogs.LogWriter("Complete BackUp Scheduler: " + DateTime.Now.ToString());
            //    //CustomLogs.SendMails mails = new CustomLogs.SendMails("joginder.banger19@gmail.com", "Redfort@2020", "joginder.banger19@gmail.com", "Harminder.singh@mahagram.in", "Compplete BackUP AEPS  :" + DateTime.Now.ToString());
            //    //mails.SendMail("Complete Database back sucessfully ");

            //}
            //catch (Exception ex)
            //{
            //    CustomLogs.SendMails mails = new CustomLogs.SendMails("joginder.banger19@gmail.com", "Redfort@2020", "joginder.banger19@gmail.com", "Harminder.singh@mahagram.in", "Faild Send pending data : " + DateTime.Now.ToString());
            //    mails.SendMail("Please check Some exception occured :" + ex.ToString());
            //}

        }


        public void Bind(DataTable _dt)
        {
            try
            {
                string TempExcelFileNameWithPath = @"BackUpFile\" + DateTime.Now.Day + "_" +DateTime.Now.Month+"_"+DateTime.Now.Year +"_"+ DateTime.Now.Minute + ".xlsx";
                XLWorkbook wb = new XLWorkbook();
                DataTable dt = _dt;
                wb.Worksheets.Add(dt, "WorksheetName");
                if(!Directory.Exists("BackUpFile"))
                {
                    Directory.CreateDirectory("BackUpFile");
                }
                else
                {
                   if( File.Exists(TempExcelFileNameWithPath))
                   File.Delete(TempExcelFileNameWithPath);
                }
                wb.SaveAs(TempExcelFileNameWithPath);
                email_send(TempExcelFileNameWithPath);
            }
            catch (Exception ex)
            {

                throw;
            }
        }

        public void email_send(string filePath)
        {
            MailMessage mail = new MailMessage();
            SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
            mail.From = new MailAddress("joginder.banger19@gmail.com");
            mail.To.Add("joginder.banger19@gmail.com");
            mail.Subject = "Pending Application Data";
            mail.Body = "Find the attachment file as per your requirments";
            System.Net.Mail.Attachment attachment;
            attachment = new System.Net.Mail.Attachment(filePath);
            mail.Attachments.Add(attachment);

            SmtpServer.Port = 587;
            SmtpServer.Credentials = new System.Net.NetworkCredential("joginder.banger19@gmail.com", "Redfort@2020");
            SmtpServer.EnableSsl = true;

            SmtpServer.Send(mail);

        }


    }
}
