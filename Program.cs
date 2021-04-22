using ExcelDataReader;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace AutomatedSendingMail
{
    public class ExcelDataVinaSun
    {
        public string VinasunCode { get; set; }
        public string OwnerName { get; set; }
        public string IsPayCard { get; set; }
        public string IsFreeCard { get; set; }
        public string VietinBank { get; set; }
        public string UserCode { get; set; }
        public string Quantity { get; set; }
        public string Code { get; set; }
        public string CustomerName { get; set; }
        public string Address { get; set; }
        public string Mail { get; set; }
        public string ReferrerUser { get; set; }
        public bool FreeCard => String.IsNullOrEmpty(IsPayCard) ? true : false;
    }
    class Program
    {

        static void Main(string[] args)
        {
            var templateFileName = "MauGuiChuyenDoiTheChip.xlsx";
            var templateFileNameRender = "MauGuiChuyenDoiTheChipDuLieu.xlsx";
            var fileData = "FileDuLieu.xlsx";
            var mailBody = "NoiDungMail.txt";
            var attachmentAll = "DinhKemChung";
            var attachFolder = "FileDinhKem";
            var mainFolder = "DuLieuMail";
            var keeperFolder = "DuLieuGuiDi";
            var errorUser = string.Empty;
            var listErrorUser = new List<string>();


            try
            {
                Console.WriteLine("Vui lòng nhập email : ");
                var mailSender = Console.ReadLine();
                Console.WriteLine("Vui lòng nhập password : ");
                var password = Console.ReadLine();
                Console.WriteLine("Vui lòng nhập Subject Mail : ");
                var mailSubject = Console.ReadLine();
                Console.WriteLine("Dữ liệu đang được xử lý, vui lòng đợi cho đến khi kết thúc và không tắt chương trình. ");

                #region ExcelData
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                var directoryPath = AppDomain.CurrentDomain.BaseDirectory;
                var directory = Environment.CurrentDirectory;
                var path = Directory.GetParent(directory).Parent.Parent.FullName + @"\" + mainFolder + @"\";
                var newFile = new FileInfo(String.Format("{0}" + "{1}", path, templateFileNameRender));
                var templateFile = new FileInfo(String.Format("{0}" + "{1}", path, templateFileName));
                //var fileBody = new FileInfo(String.Format("{0}" + "{1}", path, mailBody));
                var mailDataBody = File.ReadAllText(String.Format("{0}" + "{1}", path, mailBody), Encoding.UTF8);
                var streamData = new StreamReader(String.Format("{0}" + "{1}", path, fileData));
                var data = ExcelToListData(streamData.BaseStream);
                var pathAttachFolder = path + attachFolder + @"\" + attachmentAll;
                var listAttachment = GetAllAttachmentFile(pathAttachFolder);
                #endregion
                if (data.Any())
                {
                    var listGroupItem = data.GroupBy(x => x.UserCode).Select(x => x.First().UserCode).ToList();

                    if (listGroupItem.Any())
                    {
                        foreach (var item in listGroupItem)
                        {
                            var listGroupSubItem = data.Where(x => x.UserCode.Equals(item));
                            if (listGroupSubItem.Any())
                            {
                                try
                                {
                                    using (var pck = new ExcelPackage(newFile, templateFile))
                                    {
                                        var ws = pck.Workbook.Worksheets["Sheet1"];
                                        var firstElement = listGroupSubItem.FirstOrDefault();
                                        errorUser = firstElement.UserCode;

                                        var mailReceiver = new List<string>();
                                        if (!String.IsNullOrEmpty(firstElement.Mail))
                                        {
                                            if (firstElement.Mail.Contains(@"/"))
                                            {
                                                mailReceiver.AddRange(firstElement.Mail.Split(@"/").ToList());
                                            }
                                            else
                                            {
                                                mailReceiver.Add(firstElement.Mail);
                                            }
                                        }
                                        Console.WriteLine("Tiến hành xử lý mã " + errorUser);
                                        ws.Cells["G11"].Value = firstElement.UserCode;
                                        ws.Cells["B12"].Value = firstElement.UserCode;
                                        ws.Cells["G42"].Value = firstElement.OwnerName;
                                        ws.Cells["D22"].Value = "X";
                                        ws.Cells["D24"].Value = firstElement.Quantity;
                                        ws.Cells["G14"].Value = firstElement.CustomerName;
                                        ws.Cells["A17"].Value = firstElement.CustomerName;
                                        ws.Cells["A18"].Value = firstElement.CustomerName;
                                        ws.Cells["C19"].Value = firstElement.Address;
                                        ws.Cells["C20"].Value = firstElement.ReferrerUser;

                                        var listFreeCard = listGroupSubItem.Where(x => x.FreeCard);
                                        var listNotFreeCard = listGroupSubItem.Where(x => !x.FreeCard);

                                        var index = 42; // render data from index of row excel
                                        var length = 0;

                                        if (listFreeCard.Any())
                                        {
                                            length = listFreeCard.Count();
                                            ws.InsertRow(index, length);
                                            var i = index;
                                            var count = 1;
                                            foreach (var itemCard in listFreeCard)
                                            {
                                                var stringFormatSTT = String.Format("A{0}", i);
                                                var stringFormat = String.Format("B{0}:F{0}", i);
                                                var stringFormatB = String.Format("B{0}", i);
                                                var stringFormatG = String.Format("G{0}:K{0}", i);
                                                var stringFormatGData = String.Format("G{0}", i);
                                                ws.Cells[stringFormatSTT].Value = count;
                                                ws.Cells[stringFormatB].Value = itemCard.VinasunCode;
                                                ws.Cells[stringFormat].Merge = true;
                                                ws.Cells[stringFormatGData].Value = itemCard.OwnerName;
                                                ws.Cells[stringFormatG].Merge = true;
                                                i++;
                                                count++;
                                            }
                                        }

                                        index = index + length + 5;
                                        if (listNotFreeCard.Any())
                                        {
                                            length = listNotFreeCard.Count();
                                            ws.InsertRow(index, length);
                                            var i = index;
                                            var count = 1;
                                            foreach (var itemCard in listNotFreeCard)
                                            {
                                                var stringFormatSTT = String.Format("A{0}", i);
                                                var stringFormat = String.Format("B{0}:F{0}", i);
                                                var stringFormatB = String.Format("B{0}", i);
                                                var stringFormatG = String.Format("G{0}:K{0}", i);
                                                var stringFormatGData = String.Format("G{0}", i);
                                                ws.Cells[stringFormatSTT].Value = count;
                                                ws.Cells[stringFormatB].Value = itemCard.VinasunCode;
                                                ws.Cells[stringFormat].Merge = true;
                                                ws.Cells[stringFormatGData].Value = itemCard.OwnerName;
                                                ws.Cells[stringFormatG].Merge = true;
                                                i++;
                                                count++;
                                            }
                                        }
                                        var saveDirectory = String.Format("{0}" + "{1}", path, keeperFolder);
                                        var savePath = String.Format("{0}" + "{1}/" + "{2}.xlsx", path, keeperFolder, firstElement.UserCode);
                                        var attachmentUserFolder = String.Format("{0}" + "{1}/" + "{2}", path, attachFolder, firstElement.UserCode);
                                        if (!Directory.Exists(saveDirectory))
                                        {
                                            Directory.CreateDirectory(saveDirectory);
                                        }
                                        pck.SaveAs(new FileInfo(savePath));

                                        #region Mail Sender
                                        var mailModel = new SendEmailModel();
                                        mailModel.ListToEmails = new List<string>();
                                        mailModel.ListToEmails.Add("lamthanhphat98@gmail.com");
                                        mailModel.AttachmentFullPath = new List<string>();
                                        mailModel.AttachmentFullPath.Add(savePath);

                                        var listUserAttachment = GetAllAttachmentFile(attachmentUserFolder);
                                        if (listUserAttachment.Any())
                                            mailModel.AttachmentFullPath.AddRange(listUserAttachment);
                                        if (listAttachment.Any())
                                            mailModel.AttachmentFullPath.AddRange(listAttachment);

                                        string formatContent = mailDataBody;
                                        mailModel.Content = formatContent;
                                        mailModel.EmailSubject = "THÔNG BÁO THAY ĐỔI CHIP MỚI";//input
                                        SendMail(mailModel);
                                        Console.WriteLine("Xử lý thành công mã " + errorUser);

                                        #endregion
                                    }
                                }
                                catch (Exception)
                                {
                                    Console.WriteLine("Lỗi ở khách hàng " + errorUser);
                                    listErrorUser.Add(errorUser);

                                }
                            }
                        }
                    }
                }


                if (listErrorUser.Any())
                {
                    Console.WriteLine("Đang tiến hành ghi dữ liệu lỗi");
                    var errorMsg = String.Empty;
                    foreach (var item in listErrorUser)
                    {
                        errorMsg += item + "-";
                    }
                    var errorPathFile = String.Format("{0}" + "{1}", path, "ErrorLog.txt");
                    File.WriteAllText(errorPathFile, errorMsg);
                    Console.WriteLine("Ghi dữ liệu lỗi xong");
                }
                Console.WriteLine("Chương trình đã chạy xong");
                Console.ReadKey();
            }
            catch (Exception ex)
            {

            }

        }
        static List<string> GetAllAttachmentFile(string folderPath)
        {
            List<string> filePaths = Directory.GetFiles(folderPath, "*",
                                         SearchOption.TopDirectoryOnly).ToList();
            return filePaths;
        }
        static List<ExcelDataVinaSun> ExcelToListData(Stream dataStream)
        {
            var reader = ExcelReaderFactory.CreateReader(dataStream);
            var result = reader.AsDataSet(new ExcelDataSetConfiguration { ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = true } });
            var table = result.Tables[0];
            var datatable = DataTableToList<ExcelDataVinaSun>(table);
            return datatable;
        }

        static List<T> DataTableToList<T>(DataTable table)
        {
            try
            {
                const BindingFlags flags = BindingFlags.Public | BindingFlags.Instance;

                var count = table.Columns.Count;
                var objectProperties = typeof(T).GetProperties(flags);

                var targetList = table.AsEnumerable().Select(dataRow =>
                {

                    var instanceOfT = Activator.CreateInstance<T>();

                    for (var i = 0; i < table.Columns.Count; i++)
                    {
                        try
                        {
                            var propType = Nullable.GetUnderlyingType(objectProperties[i].PropertyType) ?? objectProperties[i].PropertyType;
                            var safeValue = String.IsNullOrEmpty(dataRow[i].ToString()) ? null : Convert.ChangeType(dataRow[i], propType);
                            objectProperties[i].SetValue(instanceOfT, safeValue, null);
                        }
                        catch (Exception ex)
                        {
                            //throw new Exception(dataRow[i].ToString());
                        }
                    }
                    return instanceOfT;
                }).ToList();

                return targetList;
            }
            catch (Exception ex)
            {
                return null;
            }

        }

        static String SignatureMail()
        {
            var signature = @"<div><div dir='ltr' data-smartmail='gmail_signature'><div dir='ltr'><div style='margin:0px;padding:0px 0px 20px;width:1102.03px;font-size:medium'><div style='margin:8px 0px 0px;padding:0px'><div dir='ltr'><div dir='ltr'><div dir='ltr'><div dir='ltr'><div dir='ltr'><div dir='ltr'><div dir='ltr'><div dir='ltr'><div dir='ltr'><div dir='ltr'><div dir='ltr'><div dir='ltr'><div dir='ltr'><div><font color='#0000ff' face='trebuchet ms, sans-serif'><b>ĐẶNG THANH TRÀ (Mr)</b></font></div><div style='font-family:Roboto,RobotoDraft,Helvetica,Arial,sans-serif;color:rgb(34,34,34)'><b><font face='trebuchet ms, sans-serif' color='#000000'>097.2222.767</font></b><b><font color='#38761d'><br></font></b></div><div style='font-family:Roboto,RobotoDraft,Helvetica,Arial,sans-serif;color:rgb(34,34,34)'><b><font face='trebuchet ms, sans-serif' color='#000000'><a href='mailto:tradangvns@gmail.com' target='_blank'>tradangvns@gmail.com</a></font></b></div><div style='font-family:Roboto,RobotoDraft,Helvetica,Arial,sans-serif;color:rgb(34,34,34)'><font face='trebuchet ms, sans-serif'><font color='#0000ff'><b>CÔNG TY C</b><b>Ổ PHẦN ÁNH D</b></font><b><font color='#0000ff'>ƯƠNG VIỆT NAM&nbsp;</font></b></font></div><div style='font-family:Roboto,RobotoDraft,Helvetica,Arial,sans-serif;color:rgb(34,34,34)'><b><font face='trebuchet ms, sans-serif' color='#00ff00'>VINASUN CORPORATION</font></b></div><div style='font-family:Roboto,RobotoDraft,Helvetica,Arial,sans-serif;color:rgb(34,34,34)'><b><font face='trebuchet ms, sans-serif' color='#00ff00'><img src='https://ci3.googleusercontent.com/proxy/ZFsGET1wIvgjiG__IeWrP5fjEmi4jJ0y1zyYyDqx_QM3If4tN-5GRcyHfGqAc3nxB7UprDmj0BUcL2myXTNdxNxkE1WwwSRpJGWJ6I7QEAdrllmC98LuHRABVRiyLtWmYnR5lM8PAt1OEeR2eg2aBM1j-hCFJaR63Yhwq1r8HVgQDU32DOuj5TvwUJ0tihMtlSgZZISQoZb3sXdQIw=s0-d-e1-ft#https://docs.google.com/uc?export=download&amp;id=19iMmMnOqbAC2PTxGhnGTYYBiyyVsoK4G&amp;revid=0BwFFskbIocABeUJabjZnVDc3bWR2ZUdpMm51ZkxJWWd2cUk4PQ' width='96' height='68' class='CToWUd'></font></b><br></div><div style='font-family:Roboto,RobotoDraft,Helvetica,Arial,sans-serif;color:rgb(34,34,34)'><b><font color='#0000ff' face='trebuchet ms, sans-serif'></font></b></div><div style='font-family:Roboto,RobotoDraft,Helvetica,Arial,sans-serif;color:rgb(34,34,34)'><font face='trebuchet ms, sans-serif'><font color='#0000ff'><b>Địa chỉ</b></font><font color='#000000'>:</font><font color='#0000ff'><b><i>Lầu 3, Tòa nhà Vinasun Tower, Số 648 Nguyễn Trãi, P. 11, Q. 5, TP.HCM</i></b></font></font></div><div dir='ltr' style='font-family:Roboto,RobotoDraft,Helvetica,Arial,sans-serif'><div style='color:rgb(34,34,34)'><a href='mailto:phulequang1971@gmail.com' target='_blank'></a></div></div></div></div></div></div></div></div></div></div></div></div></div></div></div><div style='font-family:Roboto,RobotoDraft,Helvetica,Arial,sans-serif;color:rgb(34,34,34)'></div><div style='font-family:Roboto,RobotoDraft,Helvetica,Arial,sans-serif;color:rgb(34,34,34)'></div></div><div style='font-family:Roboto,RobotoDraft,Helvetica,Arial,sans-serif;color:rgb(34,34,34)'><div style='width:1102.03px'></div><div>Vùng tệp đính kèm</div><div class='yj6qo'></div><div class='adL'></div><div class='adL'><div></div></div></div><div style='font-family:Roboto,RobotoDraft,Helvetica,Arial,sans-serif;color:rgb(34,34,34)' class='adL'></div></div></div></div></div>";
            return signature;
        }
        static void SendMail(SendEmailModel model)
        {
            var body = AlternateView.CreateAlternateViewFromString(model.Content + SignatureMail(), null, "text/html");
            model.Sender = "tradangvns@gmail.com";
            model.DisplayName = "VINASUN THÔNG BÁO";

            var msg = new MailMessage
            {
                From = new MailAddress(model.Sender, model.DisplayName),
                Subject = model.EmailSubject,
                AlternateViews = { body },
                IsBodyHtml = true,
                SubjectEncoding = Encoding.UTF8,
                BodyEncoding = Encoding.UTF8
            };

            if (model.ListToEmails != null)
            {
                foreach (var item in model.ListToEmails)
                {
                    msg.To.Add(new MailAddress(item));
                }
            }

            if (model.AttachmentFullPath.Any())
            {
                foreach (var path in model.AttachmentFullPath)
                {
                    if (File.Exists(path))
                    {
                        msg.Attachments.Add(new Attachment(path));
                    }
                }
            }

            var smtp = new SmtpClient
            {
                Host = "smtp.gmail.com",
                Port = 587,
                Credentials = new NetworkCredential(model.Sender, "Pp@198815"),
                EnableSsl = true
            };

            smtp.Send(msg);
            msg.Dispose();
        }
    }

    public class SendEmailModel
    {
        public SendEmailModel()
        {
            AttachmentFullPath = new List<string>();
            ListToEmails = new List<string>();
            ListCcEmails = new List<string>();
            ListBccEmails = new List<string>();
            EventCalendar = false;
        }

        public string Sender { get; set; }
        public string DisplayName { get; set; }
        public string EmailSubject { get; set; }

        public string Content { get; set; }
        public string ReplyTo { get; set; }
        public List<string> AttachmentFullPath { get; set; }
        public List<string> ListToEmails { get; set; }
        public List<string> ListCcEmails { get; set; }
        public List<string> ListBccEmails { get; set; }
        public bool DeleteAttachment { get; set; }
        public string ModuleName { get; set; }
        public string Priority { get; set; } = "normal"; // High, Normal, Low
        public bool EventCalendar { get; set; }
        public string CalendarStringBuilder { get; set; }
        public MemoryStream BytesCalendar { get; set; }
        public int CurrentUser { set; get; }

        public string Host { get; set; }
        public int Port { get; set; }
        public NetworkCredential Credentials { get; set; }
    }

}