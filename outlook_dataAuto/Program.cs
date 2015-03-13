using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using  Microsoft.Office.Interop.Outlook;
using System.Data.SqlClient;
using System.Data.SqlTypes;

namespace EmailSorter
{
    /// <summary>
    /// Summary description for Loader.
    /// </summary>
    public class Loader
    {

        
        static void Main(string[] args)
        {
            //autoOutlook outlooka = new autoOutlook();
            ApplicationClass my = new ApplicationClass();
            NameSpace ns = my.Session;
             ns.Logon(@"chn\huangle","generics123456.",false,true);
             int mailcount = ns.GetDefaultFolder(OlDefaultFolders.olFolderInbox).Items.Count;
             int subfolder = ns.GetDefaultFolder(OlDefaultFolders.olFolderInbox).Folders.Count;
             Console.WriteLine("mail count:" + mailcount.ToString()+"  subfolder count:"+subfolder);


             SqlConnection myconn = new SqlConnection(@"data source = (local)\sqlexpress; Integrated Security = SSPI; Initial Catalog = Intelligence_China");
             myconn.Open();

             SqlCommand mycmd = myconn.CreateCommand();
             mycmd.CommandType = System.Data.CommandType.StoredProcedure;
             mycmd.CommandText = "dbo.senders_insert";
             mycmd.Parameters.Add("@sendername", System.Data.SqlDbType.NVarChar, 2000);
             mycmd.Parameters.Add("@senton", System.Data.SqlDbType.NVarChar, 20);
             mycmd.Parameters.Add("@subject_title", System.Data.SqlDbType.NVarChar, 2000);


             foreach (object em in ns.GetDefaultFolder(OlDefaultFolders.olFolderInbox).Items)
             {
                 try
                 {
                     string subject = (string)em.GetType().InvokeMember("Subject",
                         System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.GetProperty, null, em, null, System.Globalization.CultureInfo.CurrentCulture);
                     string senderem = (string)em.GetType().InvokeMember("SenderEmailAddress",
                         System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.GetProperty, null, em, null, System.Globalization.CultureInfo.CurrentCulture);

                     string senton = em.GetType().InvokeMember("SentOn",
                          System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.GetProperty, null, em, null, System.Globalization.CultureInfo.CurrentCulture).ToString();

                     Console.WriteLine("SUBJECT:" +subject);
                     //Console.WriteLine("creationtime:" + (string)em.GetType().InvokeMember("creationtime",
                     //    System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.GetProperty, null, em, null, System.Globalization.CultureInfo.CurrentCulture));
                     Console.WriteLine("SenderEmailAddress:" + senderem);
                     //MailItemClass email = (Microsoft.Office.Interop.Outlook.MailItemClass)em;
                     
                     //Console.WriteLine("subject:" + email.ToString());
                     Console.WriteLine("SentOn:" + senton);

                     mycmd.Parameters["@sendername"].Value = senderem;
                     mycmd.Parameters["@senton"].Value = senton;
                     mycmd.Parameters["@subject_title"].Value = subject;
                     mycmd.ExecuteNonQuery();
                 }
                 catch (System.Exception e)
                 {
                     Console.WriteLine(e.Message);
                 }


             }
             Console.WriteLine("done!");
             myconn.Close();
        }

     

    }

 
}

