using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;

namespace liugyOfficeUtl
{
    class LiugMail
    {
        string Body;
        string Title;
        string CC;
        string TO;

        public void LiugyMail()
        {
        
        }

        public void Create()
        {

        }

        public void Create(string bodyTemplete)
        {

        }

        // copy from http://c-sharp-guide.com/?p=434
        /// <summary>
        /// メールを送信します
        /// </summary>
        public void Send()
        {
            try
            {
                MailMessage msg = new MailMessage();

                //送信者
                msg.From = new MailAddress("送信者のアドレス");
                //宛先（To）
                msg.To.Add(new MailAddress("宛先のアドレス1"));
                msg.To.Add(new MailAddress("宛先のアドレス2"));
                //宛先（Cc）
                msg.CC.Add(new MailAddress("CCのアドレス1"));
                msg.CC.Add(new MailAddress("CCのアドレス2"));
                //件名
                msg.Subject = "件名";
                //本文
                msg.Body = "本文";

                System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient();
                smtp.Host = "ホスト名";//SMTPサーバーを指定
                smtp.Port = 25;//（既定値は25）
                smtp.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                smtp.Send(msg);

                msg.Dispose();
                smtp.Dispose();//（.NET Framework 4.0以降）
            }
            catch (System.Exception e)
            {
                Console.Write(e.GetType().FullName + "の例外が発生しました。");
            }
        }

    }
}
