/**
 * 參考：
 * NPOI 套件
 * SharpZipLib
 * MailKit  https://github.com/jstedfast/MailKit
 * MailKit轉寄信件 https://stackoverflow.com/questions/29414995/forward-email-using-mailkit-c
 * **/

using System;
using MailKit.Net.Smtp;
using MailKit;
using MimeKit;
using MailKit.Net.Imap;
using MailKit.Security;
using System.Linq;
using System.Text.RegularExpressions;
using System.IO;
using ICSharpCode.SharpZipLib.Zip;
using FolderZipper;
using MimeKit.Tnef;
using OpenMcdf;
using System.Collections.Generic;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Threading;

namespace TurnkeyMailChecker
{
    class Program
    {

        private static MailboxAddress sender = new MailboxAddress(Properties.Settings.Default.mailUser, String.Format("{0}@{1}", Properties.Settings.Default.mailUser, Properties.Settings.Default.mailServer));
        private static String mailServer = Properties.Settings.Default.mailServer;
        private static String mailUser = Properties.Settings.Default.mailUser;
        private static String mailUserPwd = Properties.Settings.Default.mailUserPwd;

        private static void timerC(object state)
        {
            Environment.Exit(0);
        }
        public static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("檢核開始");
                checkMail();
            }
            catch (Exception ex) {
                Console.Write(ex.Message);
                sendErrorMail(ex.Message);
            }

            Console.WriteLine("檢核結束,10秒後自動關閉程式");
            Timer t = new Timer(timerC, null, 10000, 10000);
            Console.ReadLine();
        }

        private static void checkMail() {


            if (!Directory.Exists(Properties.Settings.Default.inboxDir))
                Directory.CreateDirectory(Properties.Settings.Default.inboxDir);

            using (var IMAPclient = new ImapClient())
            {
                IMAPclient.Connect(mailServer, 143, SecureSocketOptions.None);
                IMAPclient.Authenticate(mailUser, mailUserPwd);

                // The Inbox folder is always available on all IMAP servers...
                var inbox = IMAPclient.Inbox;
                inbox.Open(FolderAccess.ReadWrite);

                Console.WriteLine("總信件數: {0}", inbox.Count);
                Console.WriteLine("未讀信件: {0}", inbox.Recent);

                foreach (var summary in inbox.Fetch(0, -1, MessageSummaryItems.UniqueId | MessageSummaryItems.BodyStructure))
                {
                    var message = inbox.GetMessage(summary.UniqueId);

                    var fromAddress = message.From.Mailboxes.First().Address;
                    var subject = message.Subject;
                    Console.WriteLine("寄件者: {0}", fromAddress);
                    Console.WriteLine("主旨: {0}", subject);

                    Regex regx = new Regex(String.Format("{0}{1}", Properties.Settings.Default.companyName, @"\d{4}-\d{2}-\d{2}歷史存證檢核表"));

                    //從發票平台寄出的歷史存證檢核表 才檢查
                    if (fromAddress == Properties.Settings.Default.turnkeyMail && regx.IsMatch(subject))
                    {
                        String filename = DateTime.Now.ToString("yyyyMMddHHmmssfff");
                        String zip_path = String.Format("./{0}/{1}.zip", Properties.Settings.Default.inboxDir, filename);
                        using (var file = new FileStream(zip_path, FileMode.OpenOrCreate, FileAccess.Write))
                        {
                            foreach (var attachement in message.Attachments)
                            {
                                if (attachement.GetType() == typeof(TnefPart)) //用outlook測試轉寄信turnkey時會是 tnef的格式
                                {
                                    TnefPart tnefPart = (TnefPart)attachement;

                                    MimeMessage tnefMessage = tnefPart.ConvertToMessage();
                                    var multipart = (IEnumerable<dynamic>)tnefMessage.Body;

                                    foreach (var m in multipart)
                                    {
                                        if (((MimeEntity)m).ContentType.MimeType == "application/zip")
                                        {
                                            ((MimeContent)m.Content).WriteTo(file);
                                        }
                                    }
                                }
                                else if (attachement.GetType() == typeof(MimePart)) //turnkey寄出的會是MimePart格式
                                {
                                    MimePart mimePart = (MimePart)attachement;
                                    mimePart.Content.DecodeTo(file);
                                }
                            }
                        }

                        //解壓附件
                        String unzip_path = String.Format("{0}/{1}" , Properties.Settings.Default.inboxDir, filename);
                        using (var file = new FileStream(zip_path, FileMode.Open, FileAccess.Read))
                        {
                            ZipUtil.UnZipFiles(file, unzip_path, "");
                        }

                        foreach (var tmpFilePath in Directory.GetFiles(unzip_path, "*.xls", SearchOption.AllDirectories))
                        {
                            HSSFWorkbook hSSFWorkbook;
                            using (FileStream fs = new FileStream(tmpFilePath, FileMode.Open, FileAccess.Read))
                            {
                                hSSFWorkbook = new HSSFWorkbook(fs);
                            }

                            //傳輸比對
                            ISheet sheet = hSSFWorkbook.GetSheetAt(0);
                            var sheetRow = sheet.GetRow(4);
                            var sendDiffCnt = sheetRow.GetCell(9).NumericCellValue;//傳輸差異數
                            var storageDiffCnt = sheetRow.GetCell(10).NumericCellValue;//存證差異數

                            //存證異常清單
                            int abnormal_cnt = 0;
                            sheet = hSSFWorkbook.GetSheetAt(2);
                            bool headerFound = false;
                            for (int row = 0; row <= sheet.LastRowNum; row++)
                            {
                                sheetRow = sheet.GetRow(row);
                                if (sheetRow == null) continue;
                                var cell = sheetRow.GetCell(0);

                                if (!headerFound && cell.StringCellValue == "送方統編")
                                {
                                    headerFound = true;
                                    continue;
                                }

                                if (headerFound)
                                    if (sheetRow != null) abnormal_cnt++;
                            }
                            if (!headerFound)
                                throw new Exception("存證異常清單表頭異常!");

                            if (sendDiffCnt > 0 || storageDiffCnt > 0 || abnormal_cnt > 0)
                            {
                                String chkResultMsg = String.Format("異常，請檢查！傳輸差異數:{0}, 存證差異數:{1},存證異常清單筆數:{2}", sendDiffCnt, storageDiffCnt, abnormal_cnt);

                                Console.WriteLine(chkResultMsg);

                                forwardMail(message, chkResultMsg);
                            }
                        }

                        Directory.Delete(unzip_path, true);
                        File.Delete(zip_path);
                    }

                    //處理完後就刪除郵件
                    inbox.Store(summary.UniqueId, new StoreFlagsRequest(StoreAction.Add, MessageFlags.Deleted) { Silent = true });
                    inbox.Expunge();

                }
                IMAPclient.Disconnect(true);
            }

        }


        public static void sendErrorMail(String bodyMsg)
        {
            var message = new MimeMessage();
            message.From.Add(sender);

            var misMails = Properties.Settings.Default.misMails.Split(';');
            foreach (var m in misMails)
            {
                var user = m.Split(',');
                message.To.Add(new MailboxAddress(user[0], user[1]));
                break;
            }

            message.Subject = "電子發票檢核歷史存證郵件時發生錯誤" + DateTime.Now.ToString("yyyyMMddHHmmss");

            // now to create our body...
            var builder = new BodyBuilder();
            builder.TextBody = bodyMsg;

            message.Body = builder.ToMessageBody();

            using (var client = new SmtpClient())
            {
                client.Connect(mailServer, 25, SecureSocketOptions.None);
                client.Authenticate(mailUser, mailUserPwd);

                client.Send(message);

                client.Disconnect(true);
            }
        }

        public static void forwardMail(MimeMessage messageToForward, String bodyMsg)
        {
            var message = new MimeMessage();
            message.From.Add(sender);

            var misMails = Properties.Settings.Default.misMails.Split(';');
            foreach (var m in misMails)
            {
                var user = m.Split(',');
                message.To.Add(new MailboxAddress(user[0], user[1]));
                break;
            }           

            message.Subject = "FWD: " + messageToForward.Subject;

            var builder = new BodyBuilder();
            builder.TextBody = bodyMsg;
            builder.Attachments.Add(new MessagePart { Message = messageToForward });

            message.Body = builder.ToMessageBody();

            using (var client = new SmtpClient())
            {
                client.Connect(mailServer, 25, SecureSocketOptions.None);
                client.Authenticate(mailUser, mailUserPwd);

                client.Send(message);

                client.Disconnect(true);
            }
        }
        
        
    }
}
