using System;
//using System.Collections.Generic;
//using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using LiugyCommon;

namespace liugyOfficeUtl
{
    //************************************************************************
    /// <summary>
    /// メールフォーマットチェック、送信などメールに関連するメソッド群である。
    /// </summary>
    //************************************************************************

    public class LiugyMail
    {

        //************************************************************************
        /// <summary>
        /// IsEmailメールフォーマットを検証する。
        /// </summary>
        /// <param name="txt">メールアドレス。</param>
        /// <returns>true:有効なメールアドレスフォーマット<br /> 
        /// 　　　 　false：メールアドレスフォーマットが不正</returns>
        //************************************************************************
        public static bool IsEmail(string txt)
        {
            bool ismatch = false;
            Regex emailregex = new Regex(@"\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*");
            try
            {
                ismatch = emailregex.IsMatch(txt);
                if (ismatch)
                {
                    txt = txt.Replace(".", "");
                    txt = txt.Replace("@", "");
                    txt = txt.Replace("_", "");
                    txt = txt.Replace("-", "");
                    ismatch = LiugyDataCheck.IsAlphanumericCharacter(txt);
                }
            }
            catch
            {
                return false;
            }
            return ismatch;
        }

        //************************************************************************
        /// <summary>
        /// SMTPサーバ経由でメールを送信する。
        /// </summary>
        /// <param name="senderMail">送信者のメールアドレス</param>
        /// <param name="recipientMail">宛先</param>
        /// <param name="subject">件名</param>
        /// <param name="body">本文</param>
        /// <param name="attachmentFile">添付ファイル名</param>
        /// 
        /// <returns>true:正常に送信した。<br /> 
        /// 　　　 　false：送信が失敗しました。</returns> 
        /// <example>
        /// string[] attatchmentFile = new string[1];<br /> 
        /// attatchmentFile[0] = "c:\\data.xls";<br /> 
        /// bool sendMail = Utility.sendMail("xxx@outlook.com",<br /> 
        ///                                  "yyy@outlook.com",<br /> 
        ///                                  "【件名】テストです。",<br /> 
        ///                                  "メール正文",<br /> 
        ///                                  attatchmentFile);<br /> 
        /// </example>
        /// <remarks>
        /// 以下のネームスペースが必要です。<br /> 
        /// using System.Net;<br /> 
        /// using System.Net.Mail;<br /> 
        /// using System.Net.Mime;
        /// </remarks>
        //************************************************************************

        public static bool sendMail(
            string senderMail,      //送信者
            string recipientMail,   //宛先
            string subject,         //件名
            string body,            //本文
            string[] attachmentFile //添付ファイル名
            )
        {
            bool result = false;
            try
            {
                System.Net.Mail.SmtpClient sc = new System.Net.Mail.SmtpClient();

                byte[] arrbJisSubject;      //iso-2022-jp に変換した件名
                string strEncSubject;       //B エンコードした件名

                //件名を iso-2022-jp に変換します。
                arrbJisSubject = System.Text.ASCIIEncoding.GetEncoding("iso-2022-jp").GetBytes(subject);

                //iso-2022-jp に変換した文字列を Base64 エンコードし、エンコードを表す文字列を追加します。
                strEncSubject = "=?iso-2022-jp?B?" + Convert.ToBase64String(arrbJisSubject) + "?=";

                MailMessage mailMessage = new MailMessage(senderMail, recipientMail, strEncSubject, body);
                mailMessage.BodyEncoding = Encoding.GetEncoding("iso-2022-jp");

                if (attachmentFile != null)
                {
                    foreach (string file in attachmentFile)
                    {
                        Attachment data = new Attachment(file, MediaTypeNames.Application.Octet);

                        // 添付されたファイル名が文字化けなので、以下の処理で処置する。
                        // 詳細はマクロソフトのＨＰをご参照ください。
                        //　 .NET Framework 2.0 ベースのアプリケーションで MailMessage を
                        // 　使ってメッセージを送信すると送受信者名、件名が文字化けする
                        // 　http://support.microsoft.com/kb/933866/ja
                        byte[] arrbJisAttachment;      //iso-2022-jp に変換した件名
                        string strEncAttachment;       //B エンコードした件名

                        //添付されたファイル件名を iso-2022-jp に変換します。
                        arrbJisAttachment = System.Text.ASCIIEncoding.GetEncoding("iso-2022-jp").GetBytes(data.Name);
                        strEncAttachment = "=?iso-2022-jp?B?" + Convert.ToBase64String(arrbJisAttachment) + "?=";
                        data.Name = strEncAttachment;

                        ContentDisposition disposition = data.ContentDisposition;
                        disposition.CreationDate = System.IO.File.GetCreationTime(file);
                        disposition.ModificationDate = System.IO.File.GetLastWriteTime(file);
                        disposition.ReadDate = System.IO.File.GetLastAccessTime(file);

                        mailMessage.Attachments.Add(data);
                    }
                }

                //SMTPサーバーを指定する
                sc.Host = "smtp.xxx.comm";

                //ユーザ名とパスワードを指定する。
                //SMTPサーバが認証を要求されていない場合、この処理が不要。
                sc.Credentials = new System.Net.NetworkCredential("xxx@outlook.com", "password");

                //一秒間（1000ミリ秒）を待ちます。
                //System.Threading.Thread.Sleep(1000);

                // アンチウィルスを動いている場合、メールがすぐに送信されない場合があるようです。
                // 以下のコードを追加する必要である。
                // 以下の資料にも記述されています。
                // http://dobon.net/vb/dotnet/internet/smtpclient.html#section2
                // http://social.msdn.microsoft.com/forums/en-US/netfxnetcom/thread/6ce868ba-220f-4ff1-b755-ad9eb2e2b13d/
                sc.ServicePoint.MaxIdleTime = 100;

                //メールを送信する
                sc.Send(mailMessage);

                result = true;

            }
            catch (Exception theException)
            {

                #region エラー処理
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, theException.Source);
                //MessageBox.Show(errorMessage, "Error");
                result = false;

                #endregion
            }
            return result;
        }

    }
}
