
using dotnet_epha_api.Class;
using dotnet6_epha_api.Class;
using Microsoft.Exchange.WebServices.Data;
//using Microsoft.Office.Interop.Excel;
using Model;
using System.Data;
using System.Data.SqlClient;
using System.Security.Cryptography;
using System.Text;
using static Org.BouncyCastle.Math.EC.ECCurve;

namespace Class
{
    public class ClassEmail
    {
        //string server_url = "WebServer_ePHA_Index";// @"https://localhost:7096/hazop/Index?";
        string server_url = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("MailConfig")["WebServer_ePHA_Index"] ?? "";
        string server_url_home_task = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("MailConfig")["WebServer_ePHA_HomeTask"] ?? "";

        string sqlstr = "";
        string jsper = "";
        string s_subject = "";
        string s_body = "";

        ClassFunctions cls = new ClassFunctions();
        ClassJSON cls_json = new ClassJSON();
        ClassConnectionDb cls_conn = new ClassConnectionDb();

        private bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            Uri redirectionUri = new Uri(redirectionUrl);
            return redirectionUri.Scheme == "https";
        }
        public class sendEmailModel
        {
            public string mail_from { get; set; }
            public string mail_to { get; set; }
            public string mail_cc { get; set; }
            public string mail_subject { get; set; }
            public string mail_body { get; set; }
            public string mail_attachments { get; set; }

        }
        private static string EncryptDataWithAes(string plainText, string keyBase64, out string vectorBase64)
        {
            using (Aes aesAlgorithm = Aes.Create())
            {
                aesAlgorithm.Key = Convert.FromBase64String(keyBase64);
                aesAlgorithm.GenerateIV();
                Console.WriteLine($"Aes Cipher Mode : {aesAlgorithm.Mode}");
                Console.WriteLine($"Aes Padding Mode: {aesAlgorithm.Padding}");
                Console.WriteLine($"Aes Key Size : {aesAlgorithm.KeySize}");

                //set the parameters with out keyword
                vectorBase64 = Convert.ToBase64String(aesAlgorithm.IV);

                // Create encryptor object
                ICryptoTransform encryptor = aesAlgorithm.CreateEncryptor();

                byte[] encryptedData;

                //Encryption will be done in a memory stream through a CryptoStream object
                using (MemoryStream ms = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(ms, encryptor, CryptoStreamMode.Write))
                    {
                        using (StreamWriter sw = new StreamWriter(cs))
                        {
                            sw.Write(plainText);
                        }
                        encryptedData = ms.ToArray();
                    }
                }

                return Convert.ToBase64String(encryptedData);
            }
        }
        private static string DecryptDataWithAes(string cipherText, string keyBase64, string vectorBase64)
        {
            using (Aes aesAlgorithm = Aes.Create())
            {
                aesAlgorithm.Key = Convert.FromBase64String(keyBase64);
                aesAlgorithm.IV = Convert.FromBase64String(vectorBase64);

                Console.WriteLine($"Aes Cipher Mode : {aesAlgorithm.Mode}");
                Console.WriteLine($"Aes Padding Mode: {aesAlgorithm.Padding}");
                Console.WriteLine($"Aes Key Size : {aesAlgorithm.KeySize}");
                Console.WriteLine($"Aes Block Size : {aesAlgorithm.BlockSize}");


                // Create decryptor object
                ICryptoTransform decryptor = aesAlgorithm.CreateDecryptor();

                byte[] cipher = Convert.FromBase64String(cipherText);

                //Decryption will be done in a memory stream through a CryptoStream object
                using (MemoryStream ms = new MemoryStream(cipher))
                {
                    using (CryptoStream cs = new CryptoStream(ms, decryptor, CryptoStreamMode.Read))
                    {
                        using (StreamReader sr = new StreamReader(cs))
                        {
                            return sr.ReadToEnd();
                        }
                    }
                }
            }
        }
        private string server_url_by_action(string sub_software)
        {
            if (string.IsNullOrEmpty(sub_software))
            {
                throw new ArgumentException("Sub software cannot be null or empty", nameof(sub_software));
            }

            string xreplace_text = sub_software.ToLower();
            return server_url.Replace("dummy", xreplace_text);
        }

        public string sendMail(sendEmailModel value)
        {
            string s_mail_to = (value.mail_to ?? "").Trim();
            string s_mail_cc = (value.mail_cc ?? "").Trim();
            string s_subject = (value.mail_subject ?? "").Trim();
            string s_mail_body = (value.mail_body ?? "").Trim();
            string s_mail_attachments = (value.mail_attachments ?? "").Trim();

            string msg_mail = "";
            bool SendAndSaveCopy = false;

            // Load mail configuration from appsettings.json
            //IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build();
            string mail_from = ""; //config.GetSection("MailConfig")["MailFrom"] ?? "";
            string mail_test = ""; //config.GetSection("MailConfig")["MailTest"] ?? "";
            string mail_user = "";//config.GetSection("MailConfig")["MailUser"] ?? "";
            string mail_password = ""; //config.GetSection("MailConfig")["MailPassword"] ?? "";

            // Get test config from the database
            string sqlstr = @"SELECT DISTINCT LOWER(key_name) AS key_name, key_value 
                  FROM epha_m_config 
                  WHERE active_type = 1";

            DataTable dtConfig = new DataTable();
            dtConfig = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            if (dtConfig?.Rows.Count > 0)
            {
                // เก็บคีย์ที่ต้องการตรวจสอบใน array
                string[] keys = { "mailuser", "mailpassword", "mailfrom", "mailtest" };

                // วนลูปผ่านคีย์ต่างๆ และดึงค่าจาก DataTable
                foreach (var key in keys)
                {
                    DataRow[] dr = dtConfig.AsEnumerable().Where(row => row.Field<string>("key_name") == key).ToArray();
                    if (dr?.Length > 0)
                    {
                        string x = dr[0]["key_value"]?.ToString() ?? "";

                        // ตรวจสอบและกำหนดค่าตามคีย์
                        switch (key)
                        {
                            case "mailuser":
                                mail_user = x;
                                break;
                            case "mailpassword":
                                mail_password = x; // ปลดล็อกถ้าจำเป็น: mail_password = Decrypt(value);
                                break;
                            case "mailfrom":
                                mail_from = x;
                                break;
                            case "mailtest":
                                mail_test = x;
                                break;
                        }
                    }
                    else
                    {
                        return $"{key} not found in config.";
                    }
                }
            }
            else
            {
                return "No config mail.";
            }



            // Get test emails from the database
            sqlstr = @"SELECT DISTINCT EMAIL, EMAIL AS USER_EMAIL FROM EPHA_M_CONFIGMAIL WHERE ACTIVE_TYPE = 1";
            DataTable dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            if (dt?.Rows.Count > 0)
            {
                mail_test = string.Join(";", dt.AsEnumerable().Select(row => row["user_email"].ToString()));
            }
            if (!string.IsNullOrEmpty(mail_test))
            {
                s_mail_body += $"</br></br>ข้อมูล mail to: {s_mail_to}</br></br>ข้อมูล mail cc: {s_mail_cc}";
                s_mail_to = mail_test;
                s_mail_cc = mail_test;
            }

            string mail_font = "Cordia New";
            string mail_fontsize = "18";

            s_mail_body = $"<html><div style='font-size:{mail_fontsize}px; font-family:{mail_font};'>{s_mail_body}</div></html>";

            try
            {
                ExchangeService service = new ExchangeService();
                service.Credentials = new WebCredentials(mail_user, mail_password);
                service.TraceEnabled = true;
                service.AutodiscoverUrl(mail_user, RedirectionUrlValidationCallback);

                EmailMessage email = new EmailMessage(service)
                {
                    From = new EmailAddress("Mail Display ใส่ไม่มีผล", mail_from),
                    Subject = !string.IsNullOrEmpty(mail_test) ? "(DEV)" + s_subject : s_subject,
                    Body = new MessageBody(BodyType.HTML, s_mail_body)
                };

                foreach (string to in s_mail_to.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    email.ToRecipients.Add(to.Trim());
                }

                foreach (string cc in s_mail_cc.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    email.CcRecipients.Add(cc.Trim());
                }

                if (!string.IsNullOrEmpty(s_mail_attachments))
                {
                    foreach (string attachment in s_mail_attachments.Split(new[] { '|', '|' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        email.Attachments.AddFileAttachment(attachment.Trim());
                    }
                }

                if (SendAndSaveCopy)
                {
                    email.SendAndSaveCopy();
                }
                else
                {
                    email.Send();
                }

                //msg_mail = "Email sent successfully.";
                msg_mail = "";
            }
            catch (Exception ex)
            {
                msg_mail = ex.Message;
            }

            return msg_mail;
        }


        #region  mail create user
        public string MailToAdminRegisterAccount(string _user_displayname, string _user_email, string _user_password, string _user_password_confirm)
        {
            string user_displayname = _user_displayname;

            string urlAccept = "";
            string urlNotAccept = "";

            string to_displayname = "All";
            string s_mail_to = get_mail_admin_group();
            string s_mail_cc = "";
            string s_mail_from = "";

            string msg = "";

            // Generate URLs for Accept and Not Accept actions
            using (Aes aesAlgorithm = Aes.Create())
            {
                aesAlgorithm.KeySize = 256;
                aesAlgorithm.GenerateKey();
                string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                // URL for Accept
                string plainText = $"user_email={_user_email}&accept_status=1";
                string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);
                urlAccept = $"{server_url.Replace("dummy", "login").Replace("index", "RegisterAccount")}{cipherText}&{keyBase64}&{vectorBase64}";

                // URL for Not Accept
                plainText = $"user_email={_user_email}&accept_status=0";
                cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                urlNotAccept = $"{server_url.Replace("dummy", "login").Replace("index", "RegisterAccount")}{cipherText}&{keyBase64}&{vectorBase64}";
            }

            string s_subject = "ePHA Staff or Contractor register account.";

            string s_body = "<html><body><font face='tahoma' size='2'>";
            s_body += $"Dear {to_displayname},";
            s_body += $"<br/><br/>{user_displayname} register account.";
            s_body += $"<br/>Email address: {_user_email}";
            s_body += $"<br/>Password: {_user_password}";
            s_body += $"<br/>Confirm Password: {_user_password_confirm}";
            s_body += "<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please check your registration to use the system.";
            s_body += $"<br/><font color='blue'><a href='{urlAccept}'>Accept</a></font>";
            s_body += $",<font color='red'><a href='{urlNotAccept}'>Not Accept</a></font>";
            s_body += "<br/><br/>Best Regards,";
            s_body += "<br/>ePHA Online System ";
            s_body += "<br/><br/><br/>Note that this message was automatically sent by ePHA Online System.";
            s_body += "</font></body></html>";

            sendEmailModel data = new sendEmailModel
            {
                mail_subject = s_subject,
                mail_body = s_body,
                mail_to = s_mail_to,
                mail_cc = s_mail_cc,
                mail_from = s_mail_from
            };

            msg = sendMail(data);

            return msg;
        }
        public string MailToUserRegisterAccount(string _user_displayname, string _user_email, string _user_password, string _user_password_confirm, string _accept_status)
        {
            string user_displayname = _user_displayname;
            string url = "";
            //string step_text = "Register Account";
            string to_displayname = _user_displayname;
            string s_mail_to = _user_email;  // Send mail to the user
            string s_mail_cc = "";
            string s_mail_from = "";

            string msg = "";

            // Generate URL for user login
            using (Aes aesAlgorithm = Aes.Create())
            {
                aesAlgorithm.KeySize = 256;
                aesAlgorithm.GenerateKey();
                string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                // URL for user login
                string plainText = $"user_email={_user_email}";
                string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);
                url = $"{server_url_by_action("login")}{cipherText}&{keyBase64}&{vectorBase64}";
            }

            string s_subject = "ePHA Staff or Contractor Register Account.";

            string s_body = "<html><body><font face='tahoma' size='2'>";
            s_body += $"Dear {to_displayname},";
            s_body += "<br/><br/>Register account.";
            s_body += $"<br/><br/>Name: {user_displayname}";
            s_body += $"<br/>Email address: {_user_email}";
            s_body += $"<br/>Password: {_user_password}";
            s_body += $"<br/>Confirm Password: {_user_password_confirm}";
            s_body += $"<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Admin has {(_accept_status == "1" ? "accepted" : "not accepted")} your registration account.";
            s_body += $"<br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;System Administrator has {(_accept_status == "1" ? "confirmed" : "not confirmed")} your system registration.";
            if (_accept_status == "1")
            {
                s_body += $"<br/><font color='red'>You can now click <a href='{url}'>the link</a> to access the system.</font>";
            }
            s_body += "<br/><br/>Best Regards,";
            s_body += "<br/>ePHA Online System";
            s_body += "<br/><br/><br/>Note that this message was automatically sent by ePHA Online System.";
            s_body += "</font></body></html>";

            sendEmailModel data = new sendEmailModel
            {
                mail_subject = s_subject,
                mail_body = s_body,
                mail_to = s_mail_to,
                mail_cc = s_mail_cc,
                mail_from = s_mail_from
            };

            msg = sendMail(data);

            return msg;
        }
        #endregion  mail create user

        #region mail workflow
        public string get_mail_admin_group()
        {
            StringBuilder emailList = new StringBuilder();

            ClassLogin cls_login = new ClassLogin();
            DataTable dt = cls_login.dataEmployeeRole("");

            if (dt?.Rows.Count > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    string user_email = row["user_email"]?.ToString() ?? "";
                    string role_type = row["role_type"]?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(user_email)! && string.IsNullOrEmpty(role_type))
                    {
                        if (role_type == "admin")
                        {
                            if (emailList.Length > 0)
                            {
                                emailList.Append(";");
                            }
                            emailList.Append(user_email);
                        }
                    }
                }
            }

            return emailList.ToString();
        }

        public string QueryActionOwner(string seq, string responder_user_name, string sub_software, ref List<SqlParameter> parameters)
        {
            cls = new ClassFunctions();
            StringBuilder sqlstr = new StringBuilder();

            sqlstr.Append(@"select h.pha_status, h.pha_sub_software, h.pha_no, g.pha_request_name as pha_name, empre.user_email as request_email
                    , a.responder_user_name, emp.user_displayname, emp.user_email
                    , count(1) as total
                    , count(case when lower(a.action_status) = 'open' then 1 else null end) as 'open'
                    , count(case when lower(a.action_status) = 'closed' then 1 else null end) as 'closed' 
                    , g.reference_moc
                    from epha_t_header h
                    inner join EPHA_T_GENERAL g on lower(h.id) = lower(g.id_pha) ");

            if (sub_software == "hazop")
            {
                sqlstr.Append("left join EPHA_T_NODE_WORKSHEET a on lower(h.id) = lower(a.id_pha) ");
            }
            else if (sub_software == "jsea")
            {
                sqlstr.Append("left join EPHA_T_LIST_WORKSHEET a on lower(h.id) = lower(a.id_pha) ");
            }
            else if (sub_software == "whatif")
            {
                sqlstr.Append("left join EPHA_T_TASKS_WORKSHEET a on h.id = a.id_pha ");
            }

            sqlstr.Append(@"left join VW_EPHA_PERSON_DETAILS emp on lower(a.responder_user_name) = lower(emp.user_name)  
                    left join VW_EPHA_PERSON_DETAILS empre on lower(h.pha_request_by) = lower(empre.user_name)  
                    where a.responder_user_name is not null and h.id = @seq ");

            parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar) { Value = seq });

            if (!string.IsNullOrEmpty(responder_user_name))
            {
                sqlstr.Append(" and lower(a.responder_user_name) = lower(@responder_user_name) ");
                parameters.Add(new SqlParameter("@responder_user_name", SqlDbType.VarChar, 50) { Value = responder_user_name });
            }

            sqlstr.Append(@"group by h.pha_status, h.pha_sub_software, h.pha_no, g.pha_request_name, empre.user_email, a.responder_user_name, emp.user_displayname, emp.user_email, g.reference_moc");

            return sqlstr.ToString();
        }


        public string QueryActionOwnerUpperTA2(string seq, string approver_user_name, string sub_software, ref List<SqlParameter> parameters)
        {
            cls = new ClassFunctions();
            StringBuilder sqlstr = new StringBuilder();

            sqlstr.Append(@"
                            select h.pha_status, h.pha_sub_software, h.pha_no, g.pha_request_name as pha_name, empre.user_email as request_email,
                            nw.responder_user_name, emp.user_displayname, emp.user_email,
                            count(1) as total,
                            count(case when lower(nw.action_status) = 'open' then 1 else null end) as 'open',
                            count(case when lower(nw.action_status) = 'closed' then 1 else null end) as 'closed',
                            g.reference_moc 
                            from epha_t_header h
                            inner join EPHA_T_GENERAL g on lower(h.id) = lower(g.id_pha) 
                            inner join EPHA_T_SESSION s on lower(h.id) = lower(s.id_pha)  
                            inner join (select max(id) as id_session, id_pha from EPHA_T_SESSION group by id_pha) s1 on h.id = s1.id_pha and s.id = s1.id_session and s.id_pha = s1.id_pha  
                            inner join EPHA_T_APPROVER ta2 on lower(h.id) = lower(ta2.id_pha) and s1.id_session = ta2.id_session and ta2.action_review not in (2) 
                            left join EPHA_T_NODE_WORKSHEET nw on lower(h.id) = lower(nw.id_pha) 
                            left join VW_EPHA_PERSON_DETAILS emp on lower(nw.responder_user_name) = lower(emp.user_name)  
                            left join VW_EPHA_PERSON_DETAILS empre on lower(h.pha_request_by) = lower(empre.user_name)  
                            where nw.responder_user_name is not null and h.seq = @seq ");

            parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar) { Value = seq });

            if (!string.IsNullOrEmpty(approver_user_name))
            {
                sqlstr.Append(" and lower(ta2.user_name) = lower(@approver_user_name) ");
                parameters.Add(new SqlParameter("@approver_user_name", SqlDbType.VarChar, 50) { Value = approver_user_name });
            }

            sqlstr.Append(@"
                 group by h.pha_status, h.pha_sub_software, h.pha_no, g.pha_request_name, empre.user_email, nw.responder_user_name, emp.user_displayname, emp.user_email,
                 nw.action_status, g.reference_moc");

            string query = sqlstr.ToString();

            // เนื่องจากโครงสร้าง field ที่นำมาใช้งานเหมือนกัน
            if (sub_software == "jsea")
            {
                query = query.Replace("EPHA_T_NODE_WORKSHEET", "EPHA_T_TASKS_WORKSHEET");
            }
            else if (sub_software == "whatif")
            {
                query = query.Replace("EPHA_T_NODE_WORKSHEET", "EPHA_T_LIST_WORKSHEET");
            }

            return query;
        }

        public string MailToActionOwner(string seq, string sub_software)
        {

            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }

            #region call function export full report pdf 
            string _ram_type = sub_software switch
            {
                "hazop" => "5",
                "jsea" => "5",
                "whatif" => "5",
                _ => ""
            };
            ReportByWorksheetModel param = new ReportByWorksheetModel
            {
                seq = seq,
                export_type = "pdf",
                sub_software = sub_software,
                user_name = "",
                ram_type = _ram_type
            };
            ClassExcel clsExcel = new ClassExcel();
            string file_ResponseSheet = clsExcel.export_report_recommendation(param, false, true, true);

            if (string.IsNullOrEmpty(file_ResponseSheet))
            {
                return "Invalid file.";
            }
            else
            {
                // ตรวจสอบว่าไฟล์ที่ได้มามีอยู่จริงหรือไม่
                if (!File.Exists(file_ResponseSheet))
                {
                    return $"The specified file does not exist.{file_ResponseSheet}";
                }
            }
            #endregion call function export full report pdf 


            string doc_no = "";
            string doc_name = "";
            string reference_moc = "";

            string url = "";
            string step_text = "PHA Follow up Item";

            string to_displayname = "";
            string s_mail_to = "";
            string s_mail_cc = "";
            string s_mail_from = "";

            //string meeting_date = "";

            DataTable dt = new DataTable();
            var parameters = new List<SqlParameter>();

            sqlstr = QueryActionOwner(seq, "", sub_software, ref parameters);
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);

            #region url  
            using (Aes aesAlgorithm = Aes.Create())
            {
                aesAlgorithm.KeySize = 256;
                aesAlgorithm.GenerateKey();
                string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                string plainText = $"seq={seq}&pha_no={doc_no}&step=3";
                string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);

                url = $"{server_url_by_action(sub_software)}{cipherText}&{keyBase64}&{vectorBase64}";
            }
            #endregion url 

            #region mail to
            string msg = "";
            if (dt?.Rows.Count > 0)
            {
                string xbefor = "";
                string xafter = "";
                for (int i = 0; i < dt?.Rows.Count; i++)
                {
                    xbefor = dt.Rows[i]["user_displayname"]?.ToString() ?? "";
                    if (xbefor != xafter)
                    {
                        xafter = xbefor;
                    }
                    else
                    {
                        if (i != dt?.Rows.Count - 1) { continue; }
                    }

                    s_mail_cc = dt.Rows[i]["request_email"]?.ToString() ?? "";
                    s_mail_to = dt.Rows[i]["user_email"]?.ToString() ?? "";

                    doc_no = dt.Rows[0]["pha_no"]?.ToString() ?? "";
                    doc_name = dt.Rows[0]["pha_name"]?.ToString() ?? "";
                    reference_moc = dt.Rows[0]["reference_moc"]?.ToString() ?? "";
                    to_displayname = dt.Rows[i]["user_displayname"]?.ToString() ?? "";

                    int iTotal = Convert.ToInt32(dt.Rows[i]["total"].ToString());
                    int iOpen = Convert.ToInt32(dt.Rows[i]["open"].ToString());
                    int iClosed = Convert.ToInt32(dt.Rows[i]["closed"].ToString());

                    string s_subject = $"EPHA {doc_no}{(doc_name == "" ? "" : "")}, Please follow up item and update action.";

                    string s_body = $@"
                                       <html>
                                       <body>
                                           <font face='tahoma' size='2'>
                                               Dear {to_displayname},<br/><br/>
                                               <b>Step</b> : {step_text}<br/>
                                               {(reference_moc != "" ? $"<b>Reference MOC</b> : {reference_moc}<br/>" : "")}
                                               <b>Project Name</b> : {doc_name}<br/><br/>
                                               Items Status Total: {iTotal}, Open: {iOpen}, Closed: {iClosed}<br/><br/>
                                               &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please review data of PHA No.{doc_no}<br/>
                                               To see the detailed information, <font color='red'> please click <a href='{url}'>here</a></font><br/><br/>
                                               Best Regards,<br/>
                                               ePHA Online System<br/><br/><br/>
                                               Note that this message was automatically sent by ePHA Online System.
                                           </font>
                                       </body>
                                       </html>";

                    sendEmailModel data = new sendEmailModel
                    {
                        mail_subject = s_subject,
                        mail_body = s_body,
                        mail_to = s_mail_to,
                        mail_cc = s_mail_cc,
                        mail_from = s_mail_from
                    };

                    if (!string.IsNullOrEmpty(file_ResponseSheet) && File.Exists(file_ResponseSheet))
                    {
                        data.mail_attachments = file_ResponseSheet;
                    }

                    msg = sendMail(data);
                    if (!string.IsNullOrEmpty(msg))
                    {
                        // Handle email sending error if needed
                    }
                }
            }
            #endregion mail to

            return msg;
        }

        public string MailToAdminCaseStudy(string seq, string sub_software)
        {
            string doc_no = "";
            string doc_name = "";
            string reference_moc = "";
            string url = "";
            string step_text = "Original Closed PHA.";

            string to_displayname = "All";
            string s_mail_to = "";
            string s_mail_cc = "";
            string s_mail_from = "";

            string pha_sub_software = "";
            string meeting_date = "";
            string meeting_time = "";
            DataTable dt = new DataTable();

            StringBuilder sqlstr = new StringBuilder();
            if (sub_software == "hazop")
            {
                sqlstr.Append(@"
            select distinct h.pha_status, h.pha_no as pha_no, g.pha_request_name as pha_name, empre.user_email as request_email,
                   b.no, format(a.meeting_date, 'dd MMM yyyy') as meeting_date,
                   replace(a.meeting_start_time, '1/1/1970 ', '') + ' - ' + replace(a.meeting_end_time, '1/1/1970 ', '') as meeting_time,
                   emp.user_displayname, emp.user_email, g.reference_moc,
                   lower(h.pha_sub_software) as pha_sub_software
            from epha_t_header h
            inner join EPHA_T_GENERAL g on lower(h.id) = lower(g.id_pha)
            inner join EPHA_T_SESSION a on lower(h.id) = lower(a.id_pha)
            inner join EPHA_T_MEMBER_TEAM b on lower(a.id) = lower(b.id_session)
            inner join VW_EPHA_PERSON_DETAILS emp on lower(b.user_name) = lower(emp.user_name)
            inner join VW_EPHA_PERSON_DETAILS empre on lower(h.pha_request_by) = lower(empre.user_name)
            where h.id = @seq
            order by b.no");
            }
            else if (sub_software == "whatif")
            {
                sqlstr.Append(@"
            select distinct h.pha_status, h.pha_no as pha_no, g.pha_request_name as pha_name, empre.user_email as request_email,
                   b.no, format(a.meeting_date, 'dd MMM yyyy') as meeting_date,
                   replace(a.meeting_start_time, '1/1/1970 ', '') + ' - ' + replace(a.meeting_end_time, '1/1/1970 ', '') as meeting_time,
                   emp.user_displayname, emp.user_email, g.reference_moc,
                   lower(h.pha_sub_software) as pha_sub_software
            from epha_t_header h
            inner join EPHA_T_GENERAL g on lower(h.id) = lower(g.id_pha)
            inner join EPHA_T_SESSION a on lower(h.id) = lower(a.id_pha)
            inner join EPHA_T_MEMBER_TEAM b on lower(a.id) = lower(b.id_session)
            inner join VW_EPHA_PERSON_DETAILS emp on lower(b.user_name) = lower(emp.user_name)
            inner join VW_EPHA_PERSON_DETAILS empre on lower(h.pha_request_by) = lower(empre.user_name)
            where h.id = @seq
            order by b.no");
            }

            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "" });

            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr.ToString(), parameters);


            #region mail to
            if (dt?.Rows.Count > 0)
            {
                doc_no = dt.Rows[0]["pha_no"]?.ToString() ?? "";
                doc_name = dt.Rows[0]["pha_name"]?.ToString() ?? "";
                reference_moc = dt.Rows[0]["reference_moc"]?.ToString() ?? "";
                pha_sub_software = dt.Rows[0]["pha_sub_software"]?.ToString() ?? "";
                pha_sub_software = (pha_sub_software == "hazop" || pha_sub_software == "jsea" || pha_sub_software == "hra") ? pha_sub_software.ToUpper() : pha_sub_software;
                meeting_date = dt.Rows[0]["meeting_date"]?.ToString() ?? "";
                meeting_time = dt.Rows[0]["meeting_time"]?.ToString() ?? "";

                for (int i = 0; i < dt?.Rows.Count; i++)
                {
                    if (i > 0) { s_mail_to += ";"; }
                    s_mail_to += dt.Rows[i]["user_email"]?.ToString() ?? "";
                }
            }
            #endregion mail to

            #region mail cc 
            if (dt?.Rows.Count > 0)
            {
                s_mail_cc = dt.Rows[0]["request_email"]?.ToString() ?? "";
            }
            #endregion mail cc

            #region url  
            using (Aes aesAlgorithm = Aes.Create())
            {
                aesAlgorithm.KeySize = 256;
                aesAlgorithm.GenerateKey();
                string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                string plainText = $"seq={seq}&pha_no={doc_no}&step=2";
                string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);

                url = $"{server_url_by_action(sub_software)}{cipherText}&{keyBase64}&{vectorBase64}";
            }
            #endregion url 

            string s_subject = $"EPHA : {pha_sub_software.ToUpper()}, Please Review data.";

            string s_body = $@"
        <html>
        <body>
            <font face='tahoma' size='2'>
                Dear All,<br/><br/>
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                Original Closed the {pha_sub_software.ToUpper()}, PHA No.{doc_no} as details mentioned below,<br/><br/>
                <b>Project Name</b> : {doc_name}<br/>
                <b>Date</b> : {meeting_date}<br/>
                <b>Time</b> : {meeting_time}<br/><br/>
                <b>Step</b> : {step_text}<br/>
                {(reference_moc != "" ? $"<b>Reference MOC</b> : {reference_moc}<br/>" : "")}
                <b>Project Name</b> : {doc_name}<br/><br/>
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please review data of PHA No.{doc_no}<br/>
                To see the detailed information,<font color='red'> please click <a href='{url}'>here</a></font><br/><br/>
                More details of the study, <font color='red'; text-decoration: underline;><a href='{url}'> please click here</a></font><br/><br/>
                DISCLAIMER:<br/>
                This email message (including any attachment) is intended only for the personal use of the recipient(s) named above. It is confidential and may be legally privileged. If you are not an intended recipient, any use of this information is prohibited. If you have received this communication in error, please notify us immediately by email and delete the original message. In addition, we shall not be liable or responsible for any contents, including damages resulting from any virus transmitted by this email. Any information, comment, opinion, or statement contained in this email, including any attachments (if any), is that of the author only. Furthermore, this email (including any attachment) does not create any legally binding rights or obligations whatsoever, which may only be engaged and obliged by the exchange of hard copy documents signed by duly authorized representative(s of Thai Oil Public Company Limited.
            </font>
        </body>
        </html>";

            sendEmailModel data = new sendEmailModel
            {
                mail_subject = s_subject,
                mail_body = s_body,
                mail_to = s_mail_to,
                mail_cc = s_mail_cc,
                mail_from = s_mail_from
            };

            return sendMail(data);
        }


        #endregion mail workflow

        #region mail workflow last version 
        public string MailNotificationWorkshopInvitation(string seq, string sub_software)
        {
            string doc_no = "";
            string doc_name = "";
            string reference_moc = "";
            string url = "";
            string s_mail_to = "";
            string s_mail_cc = "";
            string s_mail_from = "";

            string pha_sub_software = sub_software;
            string meeting_date = "";
            string meeting_time = "";
            DataTable dt = new DataTable();

            StringBuilder sqlstr = new StringBuilder();

            if (sub_software == "hazop")
            {
                sqlstr.Append(@"
            select distinct h.pha_status, h.pha_no as pha_no, g.pha_request_name as pha_name, empre.user_email as request_email,
                   b.no, format(a.meeting_date, 'dd MMM yyyy') as meeting_date,
                   replace(a.meeting_start_time, '1/1/1970 ', '') + ' - ' + replace(a.meeting_end_time, '1/1/1970 ', '') as meeting_time,
                   emp.user_displayname, emp.user_email, g.reference_moc,
                   lower(h.pha_sub_software) as pha_sub_software
            from epha_t_header h
            inner join EPHA_T_GENERAL g on lower(h.id) = lower(g.id_pha)
            inner join EPHA_T_SESSION a on lower(h.id) = lower(a.id_pha)
            inner join EPHA_T_MEMBER_TEAM b on lower(h.id) = lower(b.id_pha) and lower(a.id) = lower(b.id_session)
            inner join VW_EPHA_PERSON_DETAILS emp on lower(b.user_name) = lower(emp.user_name)
            inner join VW_EPHA_PERSON_DETAILS empre on lower(h.pha_request_by) = lower(empre.user_name)
            inner join (select max(id) as id_session, id_pha from EPHA_T_SESSION group by id_pha) s1 
                on h.id = s1.id_pha and a.id = s1.id_session and a.id_pha = s1.id_pha
            where h.id = @seq
            order by b.no");
            }
            else if (sub_software == "whatif")
            {
                sqlstr.Append(@"
            select distinct h.pha_status, h.pha_no as pha_no, g.pha_request_name as pha_name, empre.user_email as request_email,
                   b.no, format(a.meeting_date, 'dd MMM yyyy') as meeting_date,
                   replace(a.meeting_start_time, '1/1/1970 ', '') + ' - ' + replace(a.meeting_end_time, '1/1/1970 ', '') as meeting_time,
                   emp.user_displayname, emp.user_email, g.reference_moc,
                   lower(h.pha_sub_software) as pha_sub_software
            from epha_t_header h
            inner join EPHA_T_GENERAL g on lower(h.id) = lower(g.id_pha)
            inner join EPHA_T_SESSION a on lower(h.id) = lower(a.id_pha)
            inner join EPHA_T_MEMBER_TEAM b on lower(h.id) = lower(b.id_pha) and lower(a.id) = lower(b.id_session)
            inner join VW_EPHA_PERSON_DETAILS emp on lower(b.user_name) = lower(emp.user_name)
            inner join VW_EPHA_PERSON_DETAILS empre on lower(h.pha_request_by) = lower(empre.user_name)
            inner join (select max(id) as id_session, id_pha from EPHA_T_SESSION group by id_pha) s1 
                on h.id = s1.id_pha and a.id = s1.id_session and a.id_pha = s1.id_pha
            where h.id = @seq
            order by b.no");
            }

            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "" });

            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr.ToString(), parameters);

            #region mail to
            if (dt?.Rows.Count > 0)
            {
                doc_no = dt.Rows[0]["pha_no"]?.ToString() ?? "";
                doc_name = dt.Rows[0]["pha_name"]?.ToString() ?? "";
                reference_moc = dt.Rows[0]["reference_moc"]?.ToString() ?? "";
                pha_sub_software = dt.Rows[0]["pha_sub_software"]?.ToString() ?? "";
                pha_sub_software = (pha_sub_software.ToLower() == "hazop" || pha_sub_software.ToLower() == "jsea" || pha_sub_software.ToLower() == "hra") ? pha_sub_software.ToUpper() : "What's If";
                meeting_date = dt.Rows[0]["meeting_date"]?.ToString() ?? "";
                meeting_time = dt.Rows[0]["meeting_time"]?.ToString() ?? "";

                for (int i = 0; i < dt?.Rows.Count; i++)
                {
                    if (i > 0) { s_mail_to += ";"; }
                    s_mail_to += dt.Rows[i]["user_email"]?.ToString() ?? "";
                }
            }
            #endregion mail to

            #region mail cc 
            if (dt?.Rows.Count > 0)
            {
                s_mail_cc = dt.Rows[0]["request_email"]?.ToString() ?? "";
            }
            #endregion mail cc

            #region url  
            using (Aes aesAlgorithm = Aes.Create())
            {
                aesAlgorithm.KeySize = 256;
                aesAlgorithm.GenerateKey();
                string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                string plainText = $"seq={seq}&pha_no={doc_no}&step=2";
                string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);

                url = $"{server_url_by_action(sub_software)}{cipherText}&{keyBase64}&{vectorBase64}";
            }
            #endregion url 

            string s_subject = $"EPHA : {pha_sub_software.ToUpper()} Workshop Invitation";

            string s_body = $@"
        <html>
        <body>
            <font face='tahoma' size='2'>
                Dear All,<br/><br/>
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                All are invited to participate in the {pha_sub_software.ToUpper()} Workshop, PHA No.{doc_no} as details mentioned below,<br/><br/>
                <b>Project Name</b> : {doc_name}<br/>
                <b>Date</b> : {meeting_date}<br/>
                <b>Time</b> : {meeting_time}<br/>
                {(reference_moc != "" ? $"<br/><b>Reference MOC</b> : {reference_moc}<br/>" : "")}
                <br/><br/>
                More details of the study, <font color='red'; text-decoration: underline;><a href='{url}'> please click here</a></font><br/><br/>
                DISCLAIMER:<br/>
                This email message (including any attachment) is intended only for the personal use of the recipient(s) named above. It is confidential and may be legally privileged. If you are not an intended recipient, any use of this information is prohibited. If you have received this communication in error, please notify us immediately by email and delete the original message. In addition, we shall not be liable or responsible for any contents, including damages resulting from any virus transmitted by this email. Any information, comment, opinion, or statement contained in this email, including any attachments (if any), is that of the author only. Furthermore, this email (including any attachment) does not create any legally binding rights or obligations whatsoever, which may only be engaged and obliged by the exchange of hard copy documents signed by duly authorized representative(s of Thai Oil Public Company Limited.
            </font>
        </body>
        </html>";

            sendEmailModel data = new sendEmailModel
            {
                mail_subject = s_subject,
                mail_body = s_body,
                mail_to = s_mail_to,
                mail_cc = s_mail_cc,
                mail_from = s_mail_from
            };

            return sendMail(data);
        }

        public string MailNotificationMemberReview(string seq, string sub_software)
        {
            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }

            #region call function export full report pdf 
            string _ram_type = sub_software switch
            {
                "hazop" => "5",
                "jsea" => "5",
                "whatif" => "5",
                _ => ""
            };
            ReportModel param = new ReportModel
            {
                seq = seq,
                export_type = "pdf",
                sub_software = sub_software,
                user_name = "",
                ram_type = _ram_type
            };
            ClassExcel clsExcel = new ClassExcel();
            string file_ResponseSheet = clsExcel.export_full_report(param, true);

            if (string.IsNullOrEmpty(file_ResponseSheet))
            {
                return "Invalid file.";
            }
            else
            {
                // ตรวจสอบว่าไฟล์ที่ได้มามีอยู่จริงหรือไม่
                if (!File.Exists(file_ResponseSheet))
                {
                    return $"The specified file does not exist.{file_ResponseSheet}";
                }
            }
            #endregion call function export full report pdf 



            string url = "";
            string url_home_task = "";
            string step_text = "Outstanding Action Notification";

            string to_displayname = "All";
            string s_mail_to = "";
            string s_mail_cc = "";
            string s_mail_from = "";

            string date_now = DateTime.Now.ToString("dd/MMM/yyyy");

            DataTable dt = new DataTable();
            DataTable dtUser = new DataTable();
            string mail_admin_group = get_mail_admin_group();


            ClassNoti classNoti = new ClassNoti();
            dtUser = classNoti.DataDailyByActionRequired_TeammMember("", seq, sub_software, true);
            dt = classNoti.DataDailyByActionRequired_TeammMember("", seq, sub_software, false);


            #region mail to 
            s_mail_cc = mail_admin_group;

            string msg = "";
            if (dt?.Rows.Count > 0)
            {
                for (int i = 0; i < dtUser?.Rows.Count; i++)
                {
                    if (i > 0) { s_mail_to += ";"; }
                    s_mail_to += dtUser.Rows[i]["user_email"]?.ToString() ?? "";
                }

                if (true)
                {
                    if (string.IsNullOrEmpty(url_home_task))
                    {
                        using (Aes aesAlgorithm = Aes.Create())
                        {
                            aesAlgorithm.KeySize = 256;
                            aesAlgorithm.GenerateKey();
                            string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                            string pha_no = "";

                            string plainText = $"seq={seq}&pha_no={pha_no}";
                            string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);
                            url_home_task = $"{server_url_home_task}{cipherText}&{keyBase64}&{vectorBase64}";
                        }
                    }

                    string s_subject = $"EPHA {step_text.ToUpper()}_{date_now}";

                    StringBuilder s_body = new StringBuilder();
                    s_body.Append("<html><body><font face='tahoma' size='2'>");
                    s_body.AppendFormat("Dear {0},", to_displayname);
                    s_body.Append(@"<br/><br/>You have the following document(s) for action. Could you please proceed promptly.");
                    s_body.Append(@"<br/><br/><small style='color:red'>Note : For review action, ""Reviewer"" please response by reply this email within five working days prior auto proceed to next step.</small>");

                    s_body.Append(@"<br/>
                            <table style ='zoom: 100%;border-collapse: collapse;font-family: tahoma, geneva, sans-serif;background-color: #215289;color: #ffffff;font-weight: bold;font-size: 13px;border: 1px solid #54585d;'>   <thead>    
                                <tr>
                                    <td style ='padding: 15px;' rowspan='1'>Task</td>
                                    <td style ='padding: 15px;' rowspan='1'>PHA Type</td>
                                    <td style ='padding: 15px;' rowspan='1'>Action Required</td>
                                    <td style ='padding: 15px;' rowspan='1'>Document Number</td>
                                    <td style ='padding: 15px;' rowspan='1'>Document Title</td>
                                    <td style ='padding: 15px;' rowspan='1'>Rev.</td>
                                    <td style ='padding: 15px;' rowspan='1'>Originator</td>
                                    <td style ='padding: 15px;' rowspan='1'>Received</td>
                                    <td style ='padding: 15px;' rowspan='1'>Due Date</td>
                                    <td style ='padding: 15px;' rowspan='1'>Remaining</td> 
                                    <td style ='padding: 15px;' rowspan='1'>Consolidator</td> 
                                </tr>
                            </thead>
                            <tbody style='color: #636363;background-color: #ffffff;border: 1px solid #dddfe1;'> ");

                    int iNo = 1;
                    DataRow[] dr = dt.Select();
                    for (int a = 0; a < dr.Length; a++)
                    {
                        string doc_no = dr[a]["document_number"]?.ToString() ?? "";

                        string background_color = "white";
                        string font_color = "black";
                        int iRemaining = 0;
                        bool action_status_close = (dr[a]["remaining"]?.ToString() ?? "").ToLower() == "closed";

                        try
                        {
                            iRemaining = Convert.ToInt32(dr[a]["remaining"]?.ToString() ?? "0");
                            if (iRemaining > 3)
                            {
                                background_color = "green"; font_color = "red";
                            }
                            else if ((iRemaining > 0 && iRemaining < 3) && !action_status_close)
                            {
                                background_color = "yellow";
                            }
                            else if (iRemaining <= 0 && !action_status_close)
                            {
                                background_color = "red"; font_color = "white";
                            }
                        }
                        catch { }

                        #region url  
                        url = "";
                        string url_def = "";
                        string url_approver = "";
                        string url_reject = "";
                        string url_reject_comment = "";
                        using (Aes aesAlgorithm = Aes.Create())
                        {
                            aesAlgorithm.KeySize = 256;
                            aesAlgorithm.GenerateKey();
                            string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                            string plainText = $"seq={seq}&pha_no={doc_no}&step=3";
                            string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);
                            if (dr[a]["pha_status"]?.ToString() == "13")
                            {
                                url = $"{server_url_by_action(sub_software)}{cipherText}&{keyBase64}&{vectorBase64}";
                            }
                            else if (dr[a]["pha_status"]?.ToString() == "21")
                            {
                                plainText = $"seq={seq}&pha_no={doc_no}&step=4&approver_type=required";
                                cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                                url = $"{server_url_by_action(sub_software)}{cipherText}&{keyBase64}&{vectorBase64}";

                                plainText = $"seq={seq}&pha_no={doc_no}&step=4&approver_type=reject";
                                cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                                url_reject = $"{server_url_by_action(sub_software)}{cipherText}&{keyBase64}&{vectorBase64}";

                                plainText = $"seq={seq}&pha_no={doc_no}&step=4&approver_type=reject_no_comment";
                                cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                                url_reject_comment = $"{server_url_by_action(sub_software)}{cipherText}&{keyBase64}&{vectorBase64}";

                                plainText = $"seq={seq}&pha_no={doc_no}&step=4&approver_type=approve";
                                cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                                url_approver = $"{server_url_by_action(sub_software)}{cipherText}&{keyBase64}&{vectorBase64}";
                            }
                        }
                        #endregion url 

                        s_body.Append("<tr>");
                        s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", iNo);
                        s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", dr[a]["pha_type"]);
                        if ((dr[a]["action_required"]?.ToString() ?? "").ToLower() == "recommendation closing"
                            || (dr[a]["action_required"]?.ToString() ?? "").ToLower() == "review"
                            || (dr[a]["action_required"]?.ToString() ?? "").ToLower() == "review & approve")
                        {
                            url_def = url;
                        }
                        else if ((dr[a]["action_required"]?.ToString() ?? "").ToLower() == "approve")
                        {
                            url_def = url;
                        }
                        s_body.AppendFormat("<td style ='padding: 15px;'><a href='{0}'>{1}</a></td>", url_def, dr[a]["action_required"]);
                        s_body.AppendFormat("<td style ='padding: 15px;'><a href='{0}'>{1}</a></td>", url_def, dr[a]["document_number"]);
                        s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", dr[a]["document_title"]);
                        s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", dr[a]["rev"]);
                        s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", dr[a]["originator"]);
                        s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", dr[a]["receivesd"]);
                        s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", dr[a]["due_date"]);
                        s_body.AppendFormat("<td style ='padding: 15px; background-color:{0};color:{1}; '>{2}</td>", background_color, font_color, dr[a]["remaining"]);
                        s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", dr[a]["consolidator"]);
                        s_body.Append("</tr>");
                        iNo += 1;
                    }

                    s_body.Append("</tbody>");
                    s_body.Append("</table>");

                    s_body.Append("<br/><br/>The message of color assignment is as follow:");

                    s_body.Append("<table>");
                    s_body.Append("<tbody>");
                    s_body.Append("<tr>");
                    s_body.Append("<td style='width: 120px;padding:4px;background-color:green; color:red'><label>Green Color</label></td>");
                    s_body.Append("<td> : &gt; 3 days; this document has more than 3 days to complete your task</td>");
                    s_body.Append("</tr>");
                    s_body.Append("<tr>");
                    s_body.Append("<td style='width: 120px;padding:4px;background-color:yellow;'><label>Yellow Color</label></td>");
                    s_body.Append("<td> : &lt; 3 days; this document has less than 3 days to complete your task</td>");
                    s_body.Append("</tr>");
                    s_body.Append("<tr>");
                    s_body.Append("<td style='width: 120px;padding:4px;background-color:Red; color : white'><label>Red Color</label></td>");
                    s_body.Append("<td> : &lt;= 0 days; this document <label style='color:red'>is overdue, please urgent action</label></td>");
                    s_body.Append("</tr>");
                    s_body.Append("</tbody>");
                    s_body.Append("</table>");

                    s_body.AppendFormat("<br/><br/><a href='{0}'>Click here to access your Overall Tasks Window</a>", url_home_task);

                    s_body.Append("<br/><br/>DISCLAIMER:");
                    s_body.Append("<br/>");
                    s_body.Append(@"This email message (including any attachment) is intended only for the personal use of the recipient(s) named above. It is confidential and may be legally privileged. If you are not an intended recipient, any use of this information is prohibited. If you have received this communication in error, please notify us immediately by email and delete the original message. In addition, we shall not be liable or responsible for any contents, including damages resulting from any virus transmitted by this email. Any information, comment, opinion, or statement contained in this email, including any attachments (if any), is that of the author only. Furthermore, this email (including any attachment) does not create any legally binding rights or obligations whatsoever, which may only be engaged and obliged by the exchange of hard copy documents signed by duly authorized representative(s of Thai Oil Public Company Limited.");
                    s_body.Append("</font></body></html>");

                    sendEmailModel data = new sendEmailModel
                    {
                        mail_subject = s_subject,
                        mail_body = s_body.ToString(),
                        mail_to = s_mail_to,
                        mail_cc = s_mail_cc,
                        mail_from = s_mail_from
                    };

                    if (!string.IsNullOrEmpty(file_ResponseSheet) && File.Exists(file_ResponseSheet))
                    {
                        data.mail_attachments = file_ResponseSheet;
                    }

                    msg = sendMail(data);
                    if (msg != "")
                    {
                        data.mail_attachments = null;
                        msg = sendMail(data);
                    }
                }
            }
            #endregion mail to

            return msg;
        }
        public string MailNotificationApproverTA2eMOC(string seq, string sub_software)
        {
            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }

            #region call function export full report pdf 
            string _ram_type = sub_software switch
            {
                "hazop" => "5",
                "jsea" => "5",
                "whatif" => "5",
                _ => ""
            };
            ReportModel param = new ReportModel
            {
                seq = seq,
                export_type = "pdf",
                sub_software = sub_software,
                user_name = "",
                ram_type = _ram_type
            };
            ClassExcel clsExcel = new ClassExcel();
            string file_ResponseSheet = clsExcel.export_full_report(param, true);

            if (string.IsNullOrEmpty(file_ResponseSheet))
            {
                return "Invalid file.";
            }
            else
            {
                // ตรวจสอบว่าไฟล์ที่ได้มามีอยู่จริงหรือไม่
                if (!File.Exists(file_ResponseSheet))
                {
                    return $"The specified file does not exist.{file_ResponseSheet}";
                }
            }
            #endregion call function export full report pdf 


            string url = "";
            string url_home_task = "";
            string step_text = "TA2 Review and Approve for MOC";

            string to_displayname = "All";
            string s_mail_to = "";
            string s_mail_cc = "";
            string s_mail_from = "";

            string date_now = DateTime.Now.ToString("dd/MMM/yyyy");

            DataTable dt;
            DataTable dtUser;
            string mail_admin_group = get_mail_admin_group();


            ClassNoti classNoti = new ClassNoti();
            dtUser = classNoti.DataDailyByActionRequired_TeammMember_TA2eMOC("", seq, sub_software, true);
            dt = classNoti.DataDailyByActionRequired_TeammMember_TA2eMOC("", seq, sub_software, false);

            #region mail to 
            s_mail_cc = mail_admin_group;

            string msg = "";
            if (dt?.Rows.Count > 0)
            {
                for (int i = 0; i < dtUser?.Rows.Count; i++)
                {
                    if (i > 0) { s_mail_to += ";"; }
                    s_mail_to += dtUser.Rows[i]["user_email"]?.ToString() ?? "";

                    if (string.IsNullOrEmpty(url_home_task))
                    {
                        using (Aes aesAlgorithm = Aes.Create())
                        {
                            aesAlgorithm.KeySize = 256;
                            aesAlgorithm.GenerateKey();
                            string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                            string pha_no = ""; // (dtUser.Rows[i]["document_number"]?.ToString() ?? "");

                            string plainText = $"seq={seq}&pha_no={pha_no}";
                            string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);
                            url_home_task = $"{server_url_home_task}{cipherText}&{keyBase64}&{vectorBase64}";
                        }
                    }
                }

                if (true)
                {
                    string s_subject = $"EPHA {step_text.ToUpper()}_{date_now}";

                    StringBuilder s_body = new StringBuilder();
                    s_body.Append("<html><body><font face='tahoma' size='2'>");
                    s_body.AppendFormat("Dear {0},", to_displayname);
                    s_body.Append(@"<br/><br/>You have the following document(s) for action. Could you please proceed promptly.");
                    s_body.Append(@"<br/><br/><small style='color:red'>Note : For review action, ""Reviewer"" please response by reply this email within five working days prior auto proceed to next step.</small>");

                    s_body.Append(@"<br/>
                            <table style ='zoom: 100%;border-collapse: collapse;font-family: tahoma, geneva, sans-serif;background-color: #215289;color: #ffffff;font-weight: bold;font-size: 13px;border: 1px solid #54585d;'>   <thead>    
                                <tr>
                                    <td style ='padding: 15px;' rowspan='1'>Task</td>
                                    <td style ='padding: 15px;' rowspan='1'>PHA Type</td>
                                    <td style ='padding: 15px;' rowspan='1'>Action Required</td>
                                    <td style ='padding: 15px;' rowspan='1'>Document Number</td>
                                    <td style ='padding: 15px;' rowspan='1'>Document Title</td>
                                    <td style ='padding: 15px;' rowspan='1'>Rev.</td>
                                    <td style ='padding: 15px;' rowspan='1'>Originator</td>
                                    <td style ='padding: 15px;' rowspan='1'>Received</td>
                                    <td style ='padding: 15px;' rowspan='1'>Due Date</td>
                                    <td style ='padding: 15px;' rowspan='1'>Remaining</td> 
                                    <td style ='padding: 15px;' rowspan='1'>Consolidator</td> 
                                </tr>
                            </thead>
                            <tbody style='color: #636363;background-color: #ffffff;border: 1px solid #dddfe1;'> ");

                    int iNo = 1;
                    DataRow[] dr = dt.Select();
                    for (int a = 0; a < dr.Length; a++)
                    {
                        string doc_no = dr[a]["document_number"]?.ToString() ?? "";

                        string background_color = "white";
                        string font_color = "black";
                        int iRemaining = 0;
                        bool action_status_close = (dr[a]["remaining"]?.ToString() ?? "").ToLower() == "closed";

                        try
                        {
                            iRemaining = Convert.ToInt32(dr[a]["remaining"]?.ToString() ?? "0");
                            if (iRemaining > 3)
                            {
                                background_color = "green"; font_color = "red";
                            }
                            else if ((iRemaining > 0 && iRemaining < 3) && !action_status_close)
                            {
                                background_color = "yellow";
                            }
                            else if (iRemaining <= 0 && !action_status_close)
                            {
                                background_color = "red"; font_color = "white";
                            }
                        }
                        catch { }

                        #region url  
                        url = "";
                        string url_def = "";
                        string url_approver = "";
                        string url_reject = "";
                        string url_reject_comment = "";
                        using (Aes aesAlgorithm = Aes.Create())
                        {
                            aesAlgorithm.KeySize = 256;
                            aesAlgorithm.GenerateKey();
                            string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                            string plainText = $"seq={seq}&pha_no={doc_no}&step=3";
                            string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);
                            if (dr[a]["pha_status"]?.ToString() == "13")
                            {
                                url = $"{server_url_by_action(sub_software)}{cipherText}&{keyBase64}&{vectorBase64}";
                            }
                            else if (dr[a]["pha_status"]?.ToString() == "21")
                            {
                                plainText = $"seq={seq}&pha_no={doc_no}&step=4&approver_type=required";
                                cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                                url = $"{server_url_by_action(sub_software)}{cipherText}&{keyBase64}&{vectorBase64}";

                                plainText = $"seq={seq}&pha_no={doc_no}&step=4&approver_type=reject";
                                cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                                url_reject = $"{server_url_by_action(sub_software)}{cipherText}&{keyBase64}&{vectorBase64}";

                                plainText = $"seq={seq}&pha_no={doc_no}&step=4&approver_type=reject_no_comment";
                                cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                                url_reject_comment = $"{server_url_by_action(sub_software)}{cipherText}&{keyBase64}&{vectorBase64}";

                                plainText = $"seq={seq}&pha_no={doc_no}&step=4&approver_type=approve";
                                cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                                url_approver = $"{server_url_by_action(sub_software)}{cipherText}&{keyBase64}&{vectorBase64}";
                            }
                        }
                        #endregion url 

                        s_body.Append("<tr>");
                        s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", iNo);
                        s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", dr[a]["pha_type"]);//hazop
                        if ((dr[a]["action_required"]?.ToString() ?? "").ToLower() == "recommendation closing"
                            || (dr[a]["action_required"]?.ToString() ?? "").ToLower() == "review"
                            || (dr[a]["action_required"]?.ToString() ?? "").ToLower() == "review & approve")
                        {
                            url_def = url;
                        }
                        else if ((dr[a]["action_required"]?.ToString() ?? "").ToLower() == "approve")
                        {
                            url_def = url;
                        }
                        s_body.AppendFormat("<td style ='padding: 15px;'><a href='{0}'>{1}</a></td>", url_def, step_text);

                        s_body.AppendFormat("<td style ='padding: 15px;'><a href='{0}'>{1}</a></td>", url_def, dr[a]["document_number"]);//hazop-2023-0000023
                        s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", dr[a]["document_title"]);//xxmoc0003

                        s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", dr[a]["rev"]);
                        s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", dr[a]["originator"]);
                        s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", dr[a]["receivesd"]);
                        s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", dr[a]["due_date"]);
                        s_body.AppendFormat("<td style ='padding: 15px; background-color:{0};color:{1}; '>{2}</td>", background_color, font_color, dr[a]["remaining"]);
                        s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", dr[a]["consolidator"]);
                        s_body.Append("</tr>");
                        iNo += 1;
                    }

                    s_body.Append("</tbody>");
                    s_body.Append("</table>");

                    s_body.Append("<br/><br/>The message of color assignment is as follow:");

                    s_body.Append("<table>");
                    s_body.Append("<tbody>");
                    s_body.Append("<tr>");
                    s_body.Append("<td style='width: 120px;padding:4px;background-color:green; color:red'><label>Green Color</label></td>");
                    s_body.Append("<td> : &gt; 3 days; this document has more than 3 days to complete your task</td>");
                    s_body.Append("</tr>");
                    s_body.Append("<tr>");
                    s_body.Append("<td style='width: 120px;padding:4px;background-color:yellow;'><label>Yellow Color</label></td>");
                    s_body.Append("<td> : &lt; 3 days; this document has less than 3 days to complete your task</td>");
                    s_body.Append("</tr>");
                    s_body.Append("<tr>");
                    s_body.Append("<td style='width: 120px;padding:4px;background-color:Red; color : white'><label>Red Color</label></td>");
                    s_body.Append("<td> : &lt;= 0 days; this document <label style='color:red'>is overdue, please urgent action</label></td>");
                    s_body.Append("</tr>");
                    s_body.Append("</tbody>");
                    s_body.Append("</table>");

                    s_body.AppendFormat("<br/><br/><a href='{0}'>Click here to access your Overall Tasks Window</a>", url_home_task);

                    s_body.Append("<br/><br/>DISCLAIMER:");
                    s_body.Append("<br/>");
                    s_body.Append(@"This email message (including any attachment) is intended only for the personal use of the recipient(s) named above. It is confidential and may be legally privileged. If you are not an intended recipient, any use of this information is prohibited. If you have received this communication in error, please notify us immediately by email and delete the original message. In addition, we shall not be liable or responsible for any contents, including damages resulting from any virus transmitted by this email. Any information, comment, opinion, or statement contained in this email, including any attachments (if any), is that of the author only. Furthermore, this email (including any attachment) does not create any legally binding rights or obligations whatsoever, which may only be engaged and obliged by the exchange of hard copy documents signed by duly authorized representative(s of Thai Oil Public Company Limited.");
                    s_body.Append("</font></body></html>");

                    sendEmailModel data = new sendEmailModel
                    {
                        mail_subject = s_subject,
                        mail_body = s_body.ToString(),
                        mail_to = s_mail_to,
                        mail_cc = s_mail_cc,
                        mail_from = s_mail_from
                    };

                    if (!string.IsNullOrEmpty(file_ResponseSheet) && File.Exists(file_ResponseSheet))
                    {
                        data.mail_attachments = file_ResponseSheet;
                    }

                    msg = sendMail(data);
                    if (msg != "")
                    {
                        data.mail_attachments = null;
                        msg = sendMail(data);
                    }
                }
            }
            #endregion mail to

            return msg;
        }

        public string MailNotificationApproverQMTSReviewer(string seq, string sub_software, string pha_status_def)
        {
            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }

            #region call function export full report pdf 
            string _ram_type = sub_software switch
            {
                "hazop" => "5",
                "jsea" => "5",
                "whatif" => "5",
                _ => ""
            };
            ReportModel param = new ReportModel
            {
                seq = seq,
                export_type = "pdf",
                sub_software = sub_software,
                user_name = "",
                ram_type = _ram_type
            };
            ClassExcel clsExcel = new ClassExcel();
            string file_ResponseSheet = clsExcel.export_full_report(param, true);

            if (string.IsNullOrEmpty(file_ResponseSheet))
            {
                return "Invalid file.";
            }
            else
            {
                // ตรวจสอบว่าไฟล์ที่ได้มามีอยู่จริงหรือไม่
                if (!File.Exists(file_ResponseSheet))
                {
                    return $"The specified file does not exist.{file_ResponseSheet}";
                }
            }
            #endregion call function export full report pdf 


            string doc_no = "";
            string doc_name = "";
            string reference_moc = "";

            string url_home_task = "";
            string url = "";
            string url_approver = "";
            string url_reject_no_comment = "";
            string url_reject_comment = "";
            //string step_text = "Approver QMTS Review.";

            string pha_sub_software = "";
            string meeting_date = "";
            string meeting_time = "";

            string to_displayname = "All";
            string s_mail_to = "";
            string s_mail_cc = "";
            string s_mail_from = "";

            string note_to_approver = "";

            DataTable dt = new DataTable();

            StringBuilder sqlstr = new StringBuilder();
            sqlstr.Append(@"select distinct h.pha_status, h.pha_no as pha_no, g.pha_request_name as pha_name, empre.user_email as request_email,
                         ta2.no, format(s.meeting_date, 'dd MMM yyyy') as meeting_date,
                         replace(s.meeting_start_time,'1/1/1970 ','') +' - '+ replace(s.meeting_end_time,'1/1/1970 ','') as meeting_time,
                         emp.user_displayname, emp.user_email, g.reference_moc, lower(h.pha_sub_software) as pha_sub_software,
                         h.approver_user_name, h.approve_action_type, h.approve_status, h.approve_comment, s.note_to_approver
                         from epha_t_header h
                         inner join EPHA_T_GENERAL g on lower(h.id) = lower(g.id_pha)
                         inner join EPHA_T_SESSION s on lower(h.id) = lower(s.id_pha)
                         inner join EPHA_T_APPROVER ta2 on lower(h.id) = lower(ta2.id_pha) and s.seq = ta2.id_session
                         inner join VW_EPHA_PERSON_DETAILS emp on lower(ta2.user_name) = lower(emp.user_name)
                         inner join VW_EPHA_PERSON_DETAILS empre on lower(h.pha_request_by) = lower(empre.user_name)
                         inner join (select max(id) as id_session, id_pha from EPHA_T_SESSION group by id_pha) s1 on h.id = s1.id_pha and s.id = s1.id_session and s.id_pha = s1.id_pha
                         where h.request_approver = 1 and h.id = @seq and isnull(ta2.approver_action_type, 0) < 2
                         order by convert(int, ta2.no)");

            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "" });
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr.ToString(), parameters);

            #region mail to
            if (dt?.Rows.Count > 0)
            {
                to_displayname = dt.Rows[0]["user_displayname"]?.ToString() ?? "";

                doc_no = dt.Rows[0]["pha_no"]?.ToString() ?? "";
                doc_name = dt.Rows[0]["pha_name"]?.ToString() ?? "";
                reference_moc = dt.Rows[0]["reference_moc"]?.ToString() ?? "";
                pha_sub_software = dt.Rows[0]["pha_sub_software"]?.ToString() ?? "";
                pha_sub_software = (pha_sub_software == "hazop" || pha_sub_software == "jsea" || pha_sub_software == "hra") ? pha_sub_software.ToUpper() : pha_sub_software;

                meeting_date = dt.Rows[0]["meeting_date"]?.ToString() ?? "";
                meeting_time = dt.Rows[0]["meeting_time"]?.ToString() ?? "";

                for (int i = 0; i < dt?.Rows.Count; i++)
                {
                    if (i > 0) { s_mail_to += ";"; }
                    s_mail_to += dt.Rows[i]["user_email"]?.ToString() ?? "";
                }

                note_to_approver = dt.Rows[0]["note_to_approver"]?.ToString() ?? "";
            }
            #endregion mail to

            #region mail cc 
            if (dt?.Rows.Count > 0)
            {
                s_mail_cc = dt.Rows[0]["request_email"]?.ToString() ?? "";
            }
            #endregion mail cc

            if (true)
            {
                using (Aes aesAlgorithm = Aes.Create())
                {
                    aesAlgorithm.KeySize = 256;
                    aesAlgorithm.GenerateKey();
                    string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                    string pha_no = doc_no;
                    string plainText = $"seq={seq}&pha_no={pha_no}";
                    string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);
                    url_home_task = $"{server_url_home_task}{cipherText}&{keyBase64}&{vectorBase64}";
                }
            }
            #region url  
            if (true)
            {
                using (Aes aesAlgorithm = Aes.Create())
                {
                    string approver_url = server_url.Replace("hazop", "approve");

                    aesAlgorithm.KeySize = 256;
                    aesAlgorithm.GenerateKey();
                    string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                    string plainText = $"seq={seq}&pha_no={doc_no}&step=4&approver_type=reject";
                    string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);

                    url = $"{server_url_by_action(sub_software)}{cipherText}&{keyBase64}&{vectorBase64}";

                    //reject 
                    url_reject_comment = url;

                    //reject no comment
                    plainText = $"seq={seq}&pha_no={doc_no}&step=4&approver_type=reject_no_comment";
                    cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                    url_reject_no_comment = $"{server_url_by_action(sub_software)}{cipherText}&{keyBase64}&{vectorBase64}";

                    //approve
                    plainText = $"seq={seq}&pha_no={doc_no}&step=4&approver_type=approve";
                    cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                    url_approver = $"{server_url_by_action(sub_software)}{cipherText}&{keyBase64}&{vectorBase64}";
                }
            }
            #endregion url 

            StringBuilder s_body = new StringBuilder();
            s_body.Append("<html><body><font face='tahoma' size='2'>");
            s_body.AppendFormat("Dear {0},", to_displayname);

            s_body.Append("<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;");
            s_body.AppendFormat("Please approve the {0}, PHA No.{1} as details mentioned below,", pha_sub_software.ToUpper(), doc_no);

            s_body.AppendFormat("<br/><br/><b>Project Name</b> : {0}", doc_name);
            s_body.AppendFormat("<br/><b>Date</b> : {0}", meeting_date);
            s_body.AppendFormat("<br/><b>Time</b> : {0}", meeting_time);

            if (!string.IsNullOrEmpty(reference_moc)) { s_body.AppendFormat("<br/><b>Reference MOC</b> : {0}", reference_moc); }
            if (!string.IsNullOrEmpty(note_to_approver)) { s_body.AppendFormat("<br/><br/><b>Note to Approver</b> : {0}", note_to_approver); }

            s_body.Append("<br/><br/>");
            s_body.AppendFormat("More details of the study, <font color='red'; text-decoration: underline;><a href='{0}'> please click here</a></font>", url);

            s_body.Append("<br/><br/><b>Reply :</b>");
            s_body.AppendFormat("<br/><a style='border: none;background-color: #25b003; padding: 14px 28px;font-size: 14px;cursor: pointer;display: inline-block; ' href='{0}'>Approve</a>", url_approver);
            s_body.AppendFormat("<a style='border: none;background-color: #d90476; padding: 14px 28px;font-size: 14px;cursor: pointer;display: inline-block; margin-left: 30px;'  href='{0}'>Send back No Comment</a>", url_reject_no_comment);
            s_body.AppendFormat("<a style='border: none;background-color: #f64a8a; padding: 14px 28px;font-size: 14px;cursor: pointer;display: inline-block; margin-left: 30px;'  href='{0}'>Send back with Comment</a>", url_reject_comment);

            s_body.Append("<br/><br/>DISCLAIMER:");
            s_body.Append("<br/>");
            s_body.Append(@"This email message (including any attachment) is intended only for the personal use of the recipient(s) named above. It is confidential and may be legally privileged. If you are not an intended recipient, any use of this information is prohibited. If you have received this communication in error, please notify us immediately by email and delete the original message. In addition, we shall not be liable or responsible for any contents, including damages resulting from any virus transmitted by this email. Any information, comment, opinion, or statement contained in this email, including any attachments (if any), is that of the author only. Furthermore, this email (including any attachment) does not create any legally binding rights or obligations whatsoever, which may only be engaged and obliged by the exchange of hard copy documents signed by duly authorized representative(s of Thai Oil Public Company Limited.");
            s_body.Append("</font></body></html>");

            sendEmailModel data = new sendEmailModel
            {
                mail_subject = $"EPHA : {pha_sub_software.ToUpper()} Waiting QMTS Review",
                mail_body = s_body.ToString(),
                mail_to = s_mail_to,
                mail_cc = s_mail_cc,
                mail_from = s_mail_from
            };

            if (!string.IsNullOrEmpty(file_ResponseSheet) && File.Exists(file_ResponseSheet))
            {
                data.mail_attachments = file_ResponseSheet;
            }

            return sendMail(data);
        }
        public string MailNotificationApproverSafetyReviewer(string seq, string sub_software, string pha_status_def)
        {
            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }

            #region call function export full report pdf 
            string _ram_type = sub_software switch
            {
                "hazop" => "5",
                "jsea" => "5",
                "whatif" => "5",
                _ => ""
            };
            ReportModel param = new ReportModel
            {
                seq = seq,
                export_type = "pdf",
                sub_software = sub_software,
                user_name = "",
                ram_type = _ram_type
            };
            ClassExcel clsExcel = new ClassExcel();
            string file_ResponseSheet = clsExcel.export_full_report(param, true);

            if (string.IsNullOrEmpty(file_ResponseSheet))
            {
                return "Invalid file.";
            }
            else
            {
                // ตรวจสอบว่าไฟล์ที่ได้มามีอยู่จริงหรือไม่
                if (!File.Exists(file_ResponseSheet))
                {
                    return $"The specified file does not exist.{file_ResponseSheet}";
                }
            }
            #endregion call function export full report pdf 


            string doc_no = "";
            string doc_name = "";
            string reference_moc = "";

            string url_home_task = "";
            string url = "";
            string url_approver = "";
            string url_reject_no_comment = "";
            string url_reject_comment = "";
            string pha_sub_software = "";
            string meeting_date = "";
            string meeting_time = "";

            string to_displayname = "All";
            string s_mail_to = "";
            string s_mail_cc = "";
            string s_mail_from = "";

            string note_to_approver = "";

            DataTable dt = new DataTable();
            StringBuilder sqlstr = new StringBuilder();
            sqlstr.Append(@"select distinct h.pha_status, h.pha_no as pha_no,g.pha_request_name as pha_name,empre.user_email as request_email,
                         ta2.no, format(s.meeting_date, 'dd MMM yyyy') as meeting_date,
                         replace(s.meeting_start_time,'1/1/1970 ','') +' - '+ replace(s.meeting_end_time,'1/1/1970 ','') as meeting_time,
                         emp.user_displayname, emp.user_email, g.reference_moc, lower(h.pha_sub_software) as pha_sub_software,
                         h.approver_user_name, h.approve_action_type, h.approve_status, h.approve_comment, s.note_to_approver
                         from epha_t_header h
                         inner join EPHA_T_GENERAL g on lower(h.id) = lower(g.id_pha)
                         inner join EPHA_T_SESSION s on lower(h.id) = lower(s.id_pha)
                         inner join EPHA_T_APPROVER ta2 on lower(h.id) = lower(ta2.id_pha) and s.seq = ta2.id_session
                         inner join VW_EPHA_PERSON_DETAILS emp on lower(ta2.user_name) = lower(emp.user_name)
                         inner join VW_EPHA_PERSON_DETAILS empre on lower(h.pha_request_by) = lower(empre.user_name)
                         inner join (select max(id) as id_session, id_pha from EPHA_T_SESSION group by id_pha) s1 on h.id = s1.id_pha and s.id = s1.id_session and s.id_pha = s1.id_pha
                         where h.request_approver = 1 and h.id = @seq and isnull(ta2.approver_action_type, 0) < 2
                         order by convert(int, ta2.no)");


            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "" });
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr.ToString(), parameters);


            #region mail to
            if (dt?.Rows.Count > 0)
            {
                to_displayname = dt.Rows[0]["user_displayname"]?.ToString() ?? "";

                doc_no = dt.Rows[0]["pha_no"]?.ToString() ?? "";
                doc_name = dt.Rows[0]["pha_name"]?.ToString() ?? "";
                reference_moc = dt.Rows[0]["reference_moc"]?.ToString() ?? "";
                pha_sub_software = dt.Rows[0]["pha_sub_software"]?.ToString() ?? "";
                pha_sub_software = (pha_sub_software == "hazop" || pha_sub_software == "jsea" || pha_sub_software == "hra") ? pha_sub_software.ToUpper() : pha_sub_software;

                meeting_date = dt.Rows[0]["meeting_date"]?.ToString() ?? "";
                meeting_time = dt.Rows[0]["meeting_time"]?.ToString() ?? "";

                for (int i = 0; i < dt?.Rows.Count; i++)
                {
                    if (i > 0) { s_mail_to += ";"; }
                    s_mail_to += dt.Rows[i]["user_email"]?.ToString() ?? "";
                }

                note_to_approver = dt.Rows[0]["note_to_approver"]?.ToString() ?? "";
            }
            #endregion mail to

            #region mail cc 
            if (dt?.Rows.Count > 0)
            {
                s_mail_cc = dt.Rows[0]["request_email"]?.ToString() ?? "";
            }
            #endregion mail cc

            if (true)
            {
                using (Aes aesAlgorithm = Aes.Create())
                {
                    aesAlgorithm.KeySize = 256;
                    aesAlgorithm.GenerateKey();
                    string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                    string pha_no = doc_no;
                    string plainText = $"seq={seq}&pha_no={pha_no}";
                    string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);
                    url_home_task = $"{server_url_home_task}{cipherText}&{keyBase64}&{vectorBase64}";
                }
            }
            #region url  
            if (true)
            {
                using (Aes aesAlgorithm = Aes.Create())
                {
                    string approver_url = server_url.Replace("hazop", "approve");

                    aesAlgorithm.KeySize = 256;
                    aesAlgorithm.GenerateKey();
                    string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                    string plainText = $"seq={seq}&pha_no={doc_no}&step=4&approver_type=reject";
                    string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);

                    url = $"{server_url_by_action(sub_software)}{cipherText}&{keyBase64}&{vectorBase64}";

                    //reject 
                    url_reject_comment = url;

                    //reject no comment
                    plainText = $"seq={seq}&pha_no={doc_no}&step=4&approver_type=reject_no_comment";
                    cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                    url_reject_no_comment = $"{server_url_by_action(sub_software)}{cipherText}&{keyBase64}&{vectorBase64}";

                    //approve
                    plainText = $"seq={seq}&pha_no={doc_no}&step=4&approver_type=approve";
                    cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                    url_approver = $"{server_url_by_action(sub_software)}{cipherText}&{keyBase64}&{vectorBase64}";
                }
            }
            #endregion url 

            StringBuilder s_body = new StringBuilder();
            s_body.Append("<html><body><font face='tahoma' size='2'>");
            s_body.AppendFormat("Dear {0},", to_displayname);

            s_body.Append("<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;");
            s_body.AppendFormat("Please approve the {0}, PHA No.{1} as details mentioned below,", pha_sub_software.ToUpper(), doc_no);

            s_body.AppendFormat("<br/><br/><b>Project Name</b> : {0}", doc_name);
            s_body.AppendFormat("<br/><b>Date</b> : {0}", meeting_date);
            s_body.AppendFormat("<br/><b>Time</b> : {0}", meeting_time);

            if (!string.IsNullOrEmpty(reference_moc)) { s_body.AppendFormat("<br/><b>Reference MOC</b> : {0}", reference_moc); }
            if (!string.IsNullOrEmpty(note_to_approver)) { s_body.AppendFormat("<br/><br/><b>Note to Approver</b> : {0}", note_to_approver); }

            s_body.Append("<br/><br/>");
            s_body.AppendFormat("More details of the study, <font color='red'; text-decoration: underline;><a href='{0}'> please click here</a></font>", url);

            s_body.Append("<br/><br/><b>Reply :</b>");
            s_body.AppendFormat("<br/><a style='border: none;background-color: #25b003; padding: 14px 28px;font-size: 14px;cursor: pointer;display: inline-block; ' href='{0}'>Approve</a>", url_approver);
            s_body.AppendFormat("<a style='border: none;background-color: #d90476; padding: 14px 28px;font-size: 14px;cursor: pointer;display: inline-block; margin-left: 30px;'  href='{0}'>Send back No Comment</a>", url_reject_no_comment);
            s_body.AppendFormat("<a style='border: none;background-color: #f64a8a; padding: 14px 28px;font-size: 14px;cursor: pointer;display: inline-block; margin-left: 30px;'  href='{0}'>Send back with Comment</a>", url_reject_comment);

            s_body.Append("<br/><br/>DISCLAIMER:");
            s_body.Append("<br/>");
            s_body.Append(@"This email message (including any attachment) is intended only for the personal use of the recipient(s) named above. It is confidential and may be legally privileged. If you are not an intended recipient, any use of this information is prohibited. If you have received this communication in error, please notify us immediately by email and delete the original message. In addition, we shall not be liable or responsible for any contents, including damages resulting from any virus transmitted by this email. Any information, comment, opinion, or statement contained in this email, including any attachments (if any), is that of the author only. Furthermore, this email (including any attachment) does not create any legally binding rights or obligations whatsoever, which may only be engaged and obliged by the exchange of hard copy documents signed by duly authorized representative(s of Thai Oil Public Company Limited.");
            s_body.Append("</font></body></html>");

            sendEmailModel data = new sendEmailModel
            {
                mail_subject = $"EPHA : {pha_sub_software.ToUpper()} Waiting Safety Reviewer Review",
                mail_body = s_body.ToString(),
                mail_to = s_mail_to,
                mail_cc = s_mail_cc,
                mail_from = s_mail_from
            };

            if (!string.IsNullOrEmpty(file_ResponseSheet) && File.Exists(file_ResponseSheet))
            {
                data.mail_attachments = file_ResponseSheet;
            }

            return sendMail(data);
        }

        public string MailNotificationApproverTA2(string seq, string sub_software, string pha_status_def, string approver_user_name)
        {
            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }

            #region call function export full report pdf 
            string _ram_type = sub_software switch
            {
                "hazop" => "5",
                "jsea" => "5",
                "whatif" => "5",
                _ => ""
            };
            ReportModel param = new ReportModel
            {
                seq = seq,
                export_type = "pdf",
                sub_software = sub_software,
                user_name = approver_user_name,
                ram_type = _ram_type
            };
            ClassExcel clsExcel = new ClassExcel();
            string file_ResponseSheet = clsExcel.export_full_report(param, true);

            if (string.IsNullOrEmpty(file_ResponseSheet))
            {
                return "Invalid file.";
            }
            else
            {
                // ตรวจสอบว่าไฟล์ที่ได้มามีอยู่จริงหรือไม่
                if (!File.Exists(file_ResponseSheet))
                {
                    return $"The specified file does not exist.{file_ResponseSheet}";
                }
            }
            #endregion call function export full report pdf 

            string doc_no = "";
            string doc_name = "";
            string reference_moc = "";

            string url_home_task = "";
            string url = "";
            string url_approver = "";
            string url_reject_no_comment = "";
            string url_reject_comment = "";
            string step_text = "Approver TA2 Review.";

            string pha_sub_software = "";
            string meeting_date = "";
            string meeting_time = "";

            string to_displayname = "All";
            string s_mail_to = "";
            string s_mail_cc = "";
            string s_mail_from = "";

            string note_to_approver = "";

            DataTable dtOwner = new DataTable();
            var parameters = new List<SqlParameter>();
            string sqlstr = QueryActionOwnerUpperTA2(seq, approver_user_name, sub_software, ref parameters);

            dtOwner = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);

            StringBuilder queryBuilder = new StringBuilder();
            queryBuilder.Append(@"select distinct h.pha_status, h.pha_no as pha_no, g.pha_request_name as pha_name, empre.user_email as request_email, ta2.no, 
                          format(s.meeting_date, 'dd MMM yyyy') as meeting_date, 
                          replace(s.meeting_start_time,'1/1/1970 ','') +' - '+ replace(s.meeting_end_time,'1/1/1970 ','') as meeting_time, 
                          emp.user_displayname, emp.user_email, g.reference_moc, lower(h.pha_sub_software) as pha_sub_software, 
                          h.approver_user_name, h.approve_action_type, h.approve_status, h.approve_comment, s.note_to_approver 
                          from epha_t_header h 
                          inner join EPHA_T_GENERAL g on lower(h.id) = lower(g.id_pha) 
                          inner join EPHA_T_SESSION s on lower(h.id) = lower(s.id_pha) 
                          inner join EPHA_T_APPROVER ta2 on lower(h.id) = lower(ta2.id_pha) and s.seq = ta2.id_session 
                          inner join VW_EPHA_PERSON_DETAILS emp on lower(ta2.user_name) = lower(emp.user_name) 
                          inner join VW_EPHA_PERSON_DETAILS empre on lower(h.pha_request_by) = lower(empre.user_name) 
                          inner join (select max(id) as id_session, id_pha from EPHA_T_SESSION group by id_pha) s1 
                          on h.id = s1.id_pha and s.id = s1.id_session and s.id_pha = s1.id_pha 
                          where h.request_approver = 1 and h.id = @seq and isnull(ta2.approver_action_type, 0) < 2");

            parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "" });
            if (!string.IsNullOrEmpty(sub_software))
            {
                queryBuilder.Append(" and lower(h.pha_sub_software) = lower(@sub_software) ");
                parameters.Add(new SqlParameter("@sub_software", SqlDbType.VarChar, 50) { Value = sub_software ?? "" });
            }
            if (!string.IsNullOrEmpty(approver_user_name))
            {
                queryBuilder.Append(" and lower(ta2.user_name) = lower(@approver_user_name) ");
                parameters.Add(new SqlParameter("@approver_user_name", SqlDbType.VarChar, 50) { Value = approver_user_name ?? "" });
            }
            queryBuilder.Append(" order by convert(int,ta2.no) ");

            DataTable dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(queryBuilder.ToString(), parameters);

            if (dt?.Rows.Count == 0) { return ""; }

            #region mail to
            if (dt?.Rows.Count > 0)
            {
                to_displayname = dt.Rows[0]["user_displayname"]?.ToString() ?? "";

                doc_no = dt.Rows[0]["pha_no"]?.ToString() ?? "";
                doc_name = dt.Rows[0]["pha_name"]?.ToString() ?? "";
                reference_moc = dt.Rows[0]["reference_moc"]?.ToString() ?? "";
                pha_sub_software = dt.Rows[0]["pha_sub_software"]?.ToString() ?? "";
                pha_sub_software = (pha_sub_software == "hazop" || pha_sub_software == "jsea" || pha_sub_software == "hra") ? pha_sub_software.ToUpper() : pha_sub_software;
                meeting_date = dt.Rows[0]["meeting_date"]?.ToString() ?? "";
                meeting_time = dt.Rows[0]["meeting_time"]?.ToString() ?? "";

                for (int i = 0; i < dt?.Rows.Count; i++)
                {
                    if (i > 0) { s_mail_to += ";"; }
                    s_mail_to += dt.Rows[i]["user_email"]?.ToString() ?? "";
                }

                note_to_approver = dt.Rows[0]["note_to_approver"]?.ToString() ?? "";
            }
            #endregion mail to

            #region mail cc 
            if (dt?.Rows.Count > 0)
            {
                s_mail_cc = dt.Rows[0]["request_email"]?.ToString() ?? "";
            }
            #endregion mail cc

            if (true)
            {
                using (Aes aesAlgorithm = Aes.Create())
                {
                    aesAlgorithm.KeySize = 256;
                    aesAlgorithm.GenerateKey();
                    string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                    string pha_no = doc_no;
                    string plainText = $"seq={seq}&pha_no={pha_no}";
                    string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);
                    url_home_task = $"{server_url_home_task}{cipherText}&{keyBase64}&{vectorBase64}";
                }
            }
            #region url  
            if (true)
            {
                using (Aes aesAlgorithm = Aes.Create())
                {
                    string approver_url = server_url_by_action("approve");

                    aesAlgorithm.KeySize = 256;
                    aesAlgorithm.GenerateKey();
                    string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                    string plainText = $"seq={seq}&pha_no={doc_no}&step=4&approver_type=reject";
                    string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);

                    url = $"{server_url_by_action(sub_software)}{cipherText}&{keyBase64}&{vectorBase64}";

                    url_reject_comment = url;

                    plainText = $"seq={seq}&pha_no={doc_no}&step=4&approver_type=reject_no_comment";
                    cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                    url_reject_no_comment = $"{server_url_by_action(sub_software)}{cipherText}&{keyBase64}&{vectorBase64}";

                    plainText = $"seq={seq}&pha_no={doc_no}&step=4&approver_type=approve";
                    cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                    url_approver = $"{server_url_by_action(sub_software)}{cipherText}&{keyBase64}&{vectorBase64}";
                }
            }
            #endregion url 

            StringBuilder s_body = new StringBuilder();
            s_body.Append("<html><body><font face='tahoma' size='2'>");
            s_body.AppendFormat("Dear {0},", to_displayname);

            s_body.Append("<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;");
            s_body.AppendFormat("Please approve the {0}, PHA No.{1} as details mentioned below,", pha_sub_software.ToUpper(), doc_no);

            s_body.AppendFormat("<br/><br/><b>Project Name</b> : {0}", doc_name);
            s_body.AppendFormat("<br/><b>Date</b> : {0}", meeting_date);
            s_body.AppendFormat("<br/><b>Time</b> : {0}", meeting_time);

            if (!string.IsNullOrEmpty(reference_moc)) { s_body.AppendFormat("<br/><b>Reference MOC</b> : {0}", reference_moc); }
            s_body.AppendFormat("<br/><br/><b>Note to Approver</b> : {0}", note_to_approver);

            s_body.Append("<br/><br/>");
            s_body.AppendFormat("More details of the study, <font color='red'; text-decoration: underline;><a href='{0}'> please click here</a></font>", url);

            if (dtOwner?.Rows.Count > 0)
            {
                s_body.Append(@"<br/><br/>
                <table style ='border-collapse: collapse;font-family: Tahoma, Geneva, sans-serif;background-color: #215289;color: #ffffff;font-weight: bold;font-size: 13px;border: 1px solid #54585d;'>
                    <thead>
                     <tr>
	                     <td style ='padding: 15px;' rowspan='2'>SUB-SOFTWARE</td>
	                     <td style ='padding: 15px;' rowspan='2'>PHA NO.</td>
	                     <td style ='padding: 15px;' rowspan='2'>RESPONDER</td>
	                     <td style ='padding: 15px; text-align: center;' colspan='3'>ITEMS STATUS</td> 
                     </tr>
                        <tr>
                            <td style ='padding: 15px;'>TOTAL</td>
                            <td style ='padding: 15px;'>OPEN</td>
                            <td style ='padding: 15px;'>CLOSE</td>		
                        </tr>
                    </thead> ");

                s_body.Append("<tbody style='color: #636363;background-color: #ffffff;border: 1px solid #dddfe1;'>");
                for (int o = 0; o < dtOwner?.Rows.Count; o++)
                {
                    s_body.Append("<tr>");
                    s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", sub_software.ToUpper());
                    s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", dtOwner.Rows[o]["pha_no"]);
                    s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", dtOwner.Rows[o]["user_displayname"]);
                    s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", dtOwner.Rows[o]["total"]);
                    s_body.AppendFormat("<td style ='padding: 15px; color: red'>{0}</td>", dtOwner.Rows[o]["open"]);
                    s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", dtOwner.Rows[o]["closed"]);
                    s_body.Append("</tr>");
                }
                s_body.Append("</tbody>");
                s_body.Append("</table>");
            }

            s_body.Append("<br/><br/><b>Reply :</b>");
            s_body.AppendFormat("<br/><a style='border: none;background-color: #25b003; padding: 14px 28px;font-size: 14px;cursor: pointer;display: inline-block;' href='{0}'>Approve</a>", url_approver);
            s_body.AppendFormat("<a style='border: none;background-color: #d90476; padding: 14px 28px;font-size: 14px;cursor: pointer;display: inline-block; margin-left: 30px;' href='{0}'>Send back No Comment</a>", url_reject_no_comment);
            s_body.AppendFormat("<a style='border: none;background-color: #f64a8a; padding: 14px 28px;font-size: 14px;cursor: pointer;display: inline-block; margin-left: 30px;' href='{0}'>Send back with Comment</a>", url_reject_comment);

            s_body.Append("<br/><br/>DISCLAIMER:");
            s_body.Append("<br/>");
            s_body.Append(@"This email message (including any attachment) is intended only for the personal use of the recipient(s) named above. It is confidential and may be legally privileged. If you are not an intended recipient, any use of this information is prohibited. If you have received this communication in error, please notify us immediately by email and delete the original message. In addition, we shall not be liable or responsible for any contents, including damages resulting from any virus transmitted by this email. Any information, comment, opinion, or statement contained in this email, including any attachments (if any), is that of the author only. Furthermore, this email (including any attachment) does not create any legally binding rights or obligations whatsoever, which may only be engaged and obliged by the exchange of hard copy documents signed by duly authorized representative(s of Thai Oil Public Company Limited.");
            s_body.Append("</font></body></html>");

            sendEmailModel data = new sendEmailModel
            {
                mail_subject = s_subject,
                mail_body = s_body.ToString(),
                mail_to = s_mail_to,
                mail_cc = s_mail_cc,
                mail_from = s_mail_from
            };

            if (!string.IsNullOrEmpty(file_ResponseSheet) && File.Exists(file_ResponseSheet))
            {
                data.mail_attachments = file_ResponseSheet;
            }

            return sendMail(data);
        }

        public string MailApprovByApprover(string seq, string sub_software, string user_approver_active)
        {
            string file_ResponseSheet = "";
            get_document_file_approver(ref file_ResponseSheet, seq, user_approver_active);

            string role_type = "";
            ClassHazop clshazop = new ClassHazop();
            clshazop.check_role_user_active(user_approver_active, ref role_type);

            string doc_no = "";
            string doc_name = "";
            string pha_sub_software = "";
            string meeting_date = "";
            string meeting_time = "";
            string reference_moc = "";
            string comment = "";
            string approve_status = "";
            string approver_displayname = "XXXXX (TOP-XX)";

            string url = "";

            string to_displayname = "All";
            string s_mail_to = "";
            string s_mail_cc = "";
            string s_mail_from = "";

            #region mail to
            string mail_admin_group = get_mail_admin_group();
            s_mail_to = mail_admin_group;
            #endregion mail to

            cls = new ClassFunctions();
            DataTable dt = new DataTable();
            var parameters = new List<SqlParameter>
    {
        new SqlParameter("@seq", SqlDbType.Int) { Value = seq }
    };

            StringBuilder sqlBuilder = new StringBuilder();
            sqlBuilder.Append(@" select distinct h.pha_status, h.pha_no as pha_no, g.pha_request_name as pha_name, empre.user_email as request_email,
                        ta2.no, format(a.meeting_date, 'dd MMM yyyy') as meeting_date, 
                        replace(a.meeting_start_time,'1/1/1970 ','') + ' - ' + replace(a.meeting_end_time,'1/1/1970 ','') as meeting_time,
                        emp.user_displayname, emp.user_email, g.reference_moc, lower(h.pha_sub_software) as pha_sub_software,
                        lower(ta2.user_name) as approver_user_name, ta2.approver_action_type, ta2.action_review as approve_action_review, 
                        ta2.comment as approve_comment
                        from epha_t_header h
                        inner join EPHA_T_GENERAL g on lower(h.id) = lower(g.id_pha)
                        inner join EPHA_T_SESSION a on lower(h.id) = lower(a.id_pha)
                        inner join EPHA_T_APPROVER ta2 on lower(h.id) = lower(ta2.id_pha) and a.seq = ta2.id_session
                        inner join VW_EPHA_PERSON_DETAILS emp on lower(ta2.user_name) = lower(emp.user_name)
                        inner join VW_EPHA_PERSON_DETAILS empre on lower(h.pha_request_by) = lower(empre.user_name)
                        inner join (select max(id) as id_session, id_pha from EPHA_T_SESSION group by id_pha) s 
                        on h.id = s.id_pha and a.id = s.id_session and a.id_pha = s.id_pha
                        where h.request_approver = 1 and h.id = @seq");

            if (role_type != "admin")
            {
                sqlBuilder.Append(" and lower(ta2.user_name) like lower(@user_approver_active)");
                parameters.Add(new SqlParameter("@user_approver_active", SqlDbType.VarChar, 100) { Value = user_approver_active });
            }

            sqlBuilder.Append(" order by convert(int,ta2.no)");

            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlBuilder.ToString(), parameters);

            if (dt?.Rows.Count > 0)
            {
                doc_no = dt.Rows[0]["pha_no"]?.ToString() ?? "";
                doc_name = dt.Rows[0]["pha_name"]?.ToString() ?? "";
                reference_moc = dt.Rows[0]["reference_moc"]?.ToString() ?? "";
                comment = dt.Rows[0]["approve_comment"]?.ToString() ?? "";
                approve_status = dt.Rows[0]["approve_action_review"]?.ToString() ?? "";
                approver_displayname = dt.Rows[0]["user_displayname"]?.ToString() ?? "";

                pha_sub_software = dt.Rows[0]["pha_sub_software"]?.ToString() ?? "";
                meeting_date = dt.Rows[0]["meeting_date"]?.ToString() ?? "";
                meeting_time = dt.Rows[0]["meeting_time"]?.ToString() ?? "";

                s_mail_cc += dt.Rows[0]["user_email"]?.ToString() ?? "";
                if (!string.IsNullOrEmpty(dt.Rows[0]["request_email"]?.ToString()))
                {
                    s_mail_cc += ";" + dt.Rows[0]["request_email"];
                }
            }

            #region url  
            using (Aes aesAlgorithm = Aes.Create())
            {
                aesAlgorithm.KeySize = 256;
                aesAlgorithm.GenerateKey();
                string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                string plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=2";
                string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);

                url = server_url_by_action(sub_software) + cipherText + "&" + keyBase64 + "&" + vectorBase64;
            }
            #endregion url 

            StringBuilder s_body = new StringBuilder();
            s_body.Append("<html><body><font face='tahoma' size='2'>");
            s_body.Append("Dear All,");

            s_body.Append("<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;");
            s_body.AppendFormat("{0} has approve the {1}, PHA No.{2} as details mentioned below,", approver_displayname, pha_sub_software.ToUpper(), doc_no);

            s_body.AppendFormat("<br/><br/><b>Project Name</b> : {0}", doc_name);
            s_body.AppendFormat("<br/><b>Date</b> : {0}", meeting_date);
            s_body.AppendFormat("<br/><b>Time</b> : {0}", meeting_time);

            if (!string.IsNullOrEmpty(reference_moc)) { s_body.AppendFormat("<br/><br/><b>Reference MOC</b> : {0}", reference_moc); }
            if (!string.IsNullOrEmpty(comment))
            {
                s_body.AppendFormat("<br/><br/><b>Comment :</b>{0}", comment);
            }

            s_body.Append("<br/><br/>");
            s_body.AppendFormat("More details of the study, <font color='red'; text-decoration: underline;><a href='{0}'> please click here</a></font>", url);

            s_body.Append("<br/><br/>DISCLAIMER:");
            s_body.Append("<br/>");
            s_body.Append(@"This email message (including any attachment) is intended only for the personal use of the recipient(s) named above. It is confidential and may be legally privileged. If you are not an intended recipient, any use of this information is prohibited. If you have received this communication in error, please notify us immediately by email and delete the original message. In addition, we shall not be liable or responsible for any contents, including damages resulting from any virus transmitted by this email. Any information, comment, opinion, or statement contained in this email, including any attachments (if any), is that of the author only. Furthermore, this email (including any attachment) does not create any legally binding rights or obligations whatsoever, which may only be engaged and obliged by the exchange of hard copy documents signed by duly authorized representative(s of Thai Oil Public Company Limited.");
            s_body.Append("</font></body></html>");

            sendEmailModel data = new sendEmailModel
            {
                mail_subject = "EPHA : " + pha_sub_software.ToUpper() + ", Approver Approve.",
                mail_body = s_body.ToString(),
                mail_to = s_mail_to,
                mail_cc = s_mail_cc,
                mail_from = s_mail_from
            };

            if (!string.IsNullOrEmpty(file_ResponseSheet))
            {
                data.mail_attachments = file_ResponseSheet;
            }

            return sendMail(data);
        }

        public string MailRejectByApprover(string seq, string sub_software, string user_approver_active)
        {
            string file_ResponseSheet = "";
            get_document_file_approver(ref file_ResponseSheet, seq, user_approver_active);

            string role_type = "";
            ClassHazop clshazop = new ClassHazop();
            clshazop.check_role_user_active(user_approver_active, ref role_type);

            string doc_no = "";
            string doc_name = "";
            string pha_sub_software = "";
            string meeting_date = "";
            string meeting_time = "";
            string reference_moc = "";
            string comment = "";
            string approve_status = "";
            string approver_displayname = "XXXXX (TOP-XX)";

            string url = "";

            string to_displayname = "All";
            string s_mail_to = "";
            string s_mail_cc = "";
            string s_mail_from = "";

            #region mail to
            string mail_admin_group = get_mail_admin_group();
            s_mail_to = mail_admin_group;
            #endregion mail to

            cls = new ClassFunctions();
            DataTable dt = new DataTable();

            StringBuilder sqlBuilder = new StringBuilder();
            var parameters = new List<SqlParameter>
    {
        new SqlParameter("@seq", SqlDbType.Int) { Value = seq }
    };

            sqlBuilder.Append(@" select distinct h.pha_status, h.pha_no as pha_no, g.pha_request_name as pha_name, 
                        empre.user_email as request_email, ta2.no, 
                        format(a.meeting_date, 'dd MMM yyyy') as meeting_date, 
                        replace(a.meeting_start_time,'1/1/1970 ','') + ' - ' + replace(a.meeting_end_time,'1/1/1970 ','') as meeting_time, 
                        emp.user_displayname, emp.user_email, g.reference_moc, 
                        lower(h.pha_sub_software) as pha_sub_software, lower(ta2.user_name) as approver_user_name, 
                        ta2.approver_action_type, ta2.action_review as approve_action_review, ta2.comment as approve_comment 
                        from epha_t_header h
                        inner join EPHA_T_GENERAL g on lower(h.id) = lower(g.id_pha) 
                        inner join EPHA_T_SESSION a on lower(h.id) = lower(a.id_pha)  
                        inner join EPHA_T_APPROVER ta2 on lower(h.id) = lower(ta2.id_pha) and a.seq = ta2.id_session  
                        inner join VW_EPHA_PERSON_DETAILS emp on lower(ta2.user_name) = lower(emp.user_name) 
                        inner join VW_EPHA_PERSON_DETAILS empre on lower(h.pha_request_by) = lower(empre.user_name) 
                        inner join (select max(id) as id_session, id_pha from EPHA_T_SESSION group by id_pha) s 
                        on h.id = s.id_pha and a.id = s.id_session and a.id_pha = s.id_pha  
                        where h.request_approver = 1 and h.id = @seq");

            if (role_type != "admin")
            {
                sqlBuilder.Append(" and lower(ta2.user_name) like lower(@user_approver_active)");
                parameters.Add(new SqlParameter("@user_approver_active", SqlDbType.VarChar, 100) { Value = user_approver_active });
            }

            sqlBuilder.Append(" order by convert(int, ta2.no)");

            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlBuilder.ToString(), parameters);

            if (dt?.Rows.Count > 0)
            {
                doc_no = dt.Rows[0]["pha_no"]?.ToString() ?? "";
                doc_name = dt.Rows[0]["pha_name"]?.ToString() ?? "";
                reference_moc = dt.Rows[0]["reference_moc"]?.ToString() ?? "";
                comment = dt.Rows[0]["approve_comment"]?.ToString() ?? "";
                approve_status = dt.Rows[0]["approve_action_review"]?.ToString() ?? "";
                approver_displayname = dt.Rows[0]["user_displayname"]?.ToString() ?? "";

                pha_sub_software = dt.Rows[0]["pha_sub_software"]?.ToString() ?? "";
                meeting_date = dt.Rows[0]["meeting_date"]?.ToString() ?? "";
                meeting_time = dt.Rows[0]["meeting_time"]?.ToString() ?? "";

                s_mail_cc += dt.Rows[0]["user_email"]?.ToString() ?? "";
                if (!string.IsNullOrEmpty(dt.Rows[0]["request_email"]?.ToString()))
                {
                    s_mail_cc += ";" + dt.Rows[0]["request_email"];
                }
            }

            #region url  
            using (Aes aesAlgorithm = Aes.Create())
            {
                aesAlgorithm.KeySize = 256;
                aesAlgorithm.GenerateKey();
                string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                string plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=2";
                string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);

                url = server_url_by_action(sub_software) + cipherText + "&" + keyBase64 + "&" + vectorBase64;
            }
            #endregion url 

            StringBuilder s_body = new StringBuilder();
            s_body.Append("<html><body><font face='tahoma' size='2'>");
            s_body.Append("Dear All,");

            s_body.Append("<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;");
            s_body.AppendFormat("{0} has rejected the {1}, PHA No.{2} as details mentioned below,", approver_displayname, pha_sub_software.ToUpper(), doc_no);

            s_body.AppendFormat("<br/><br/><b>Project Name</b> : {0}", doc_name);
            s_body.AppendFormat("<br/><b>Date</b> : {0}", meeting_date);
            s_body.AppendFormat("<br/><b>Time</b> : {0}", meeting_time);

            if (!string.IsNullOrEmpty(reference_moc)) { s_body.AppendFormat("<br/><b>Reference MOC</b> : {0}", reference_moc); }

            s_body.Append("<br/><br/><b>Send back with comment :</b>");
            s_body.AppendFormat("<br/><b>{0}</b>", comment);

            s_body.Append("<br/><br/>");
            s_body.AppendFormat("More details of the study, <font color='red'; text-decoration: underline;><a href='{0}'> please click here</a></font>", url);

            s_body.Append("<br/><br/>DISCLAIMER:");
            s_body.Append("<br/>");
            s_body.Append(@"This email message (including any attachment) is intended only for the personal use of the recipient(s) named above. It is confidential and may be legally privileged. If you are not an intended recipient, any use of this information is prohibited. If you have received this communication in error, please notify us immediately by email and delete the original message. In addition, we shall not be liable or responsible for any contents, including damages resulting from any virus transmitted by this email. Any information, comment, opinion, or statement contained in this email, including any attachments (if any), is that of the author only. Furthermore, this email (including any attachment) does not create any legally binding rights or obligations whatsoever, which may only be engaged and obliged by the exchange of hard copy documents signed by duly authorized representative(s of Thai Oil Public Company Limited.");
            s_body.Append("</font></body></html>");

            sendEmailModel data = new sendEmailModel
            {
                mail_subject = "EPHA : " + pha_sub_software.ToUpper() + ", Approver Send back with Comment.",
                mail_body = s_body.ToString(),
                mail_to = s_mail_to,
                mail_cc = s_mail_cc,
                mail_from = s_mail_from
            };

            if (!string.IsNullOrEmpty(file_ResponseSheet))
            {
                data.mail_attachments = file_ResponseSheet;
            }

            return sendMail(data);
        }

        public string MailNotificationOutstandingAction(string user_name, string seq, string sub_software)
        {
            string url = "";
            string url_home_task = "";
            //string step_text = "Outstanding Action Notification";

            string to_displayname = "All";
            string s_mail_to = "";
            string s_mail_cc = "";
            string s_mail_from = "";

            string date_now = DateTime.Now.ToString("dd/MMM/yyyy");
            string file_ResponseSheet = "";

            if (sub_software == "hra")
            {
                try
                {
                    ReportModel param = new ReportModel
                    {
                        export_type = "pdf",
                        sub_software = sub_software,
                        seq = seq
                    };
                    ClassExcel cls = new ClassExcel();
                    file_ResponseSheet = cls.export_potential_health_checklist_template(param, true);
                }
                catch { }
            }

            DataTable dt = new DataTable();
            string mail_admin_group = get_mail_admin_group();

            cls_conn = new ClassConnectionDb();
            DataTable dtOwner = new DataTable();
            ClassNoti classNoti = new ClassNoti();

            dtOwner = classNoti.DataDailyByActionRequired(user_name, seq, sub_software, true, false);
            dt = classNoti.DataDailyByActionRequired(user_name, seq, sub_software, false, false);

            #region mail to
            s_mail_cc = mail_admin_group;

            string msg = "";
            if (dt?.Rows.Count > 0)
            {
                foreach (DataRow ownerRow in dtOwner.Rows)
                {
                    to_displayname = ownerRow["user_displayname"].ToString();
                    s_mail_to = ownerRow["user_email"].ToString();
                    string responder_user_name = ownerRow["user_name"].ToString();

                    using (Aes aesAlgorithm = Aes.Create())
                    {
                        aesAlgorithm.KeySize = 256;
                        aesAlgorithm.GenerateKey();
                        string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                        string pha_no = dt.Select($"user_name='{responder_user_name}'").FirstOrDefault()?["document_number"].ToString() ?? "";
                        string plainText = $"seq={seq}&pha_no={pha_no}";
                        string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);
                        url_home_task = server_url_home_task + cipherText + "&" + keyBase64 + "&" + vectorBase64;
                    }

                    StringBuilder s_body = new StringBuilder();
                    s_body.Append("<html><body><font face='tahoma' size='2'>");
                    s_body.AppendFormat("Dear {0},", to_displayname);
                    s_body.Append(@"<br/><br/>You have the following document(s) for action. Could you please proceed promptly.");
                    s_body.Append(@"<br/><br/><small style='color:red'>Note : For review action, ""Reviewer"" please response by reply this email within five working days prior auto proceed to next step.</small>");
                    s_body.Append(@"<br/>
                            <table style ='zoom: 100%;border-collapse: collapse;font-family: tahoma, geneva, sans-serif;background-color: #215289;color: #ffffff;font-weight: bold;font-size: 13px;border: 1px solid #54585d;'>   
                            <thead>    
                                <tr>
                                    <td style ='padding: 15px;' rowspan='1'>Task</td>
                                    <td style ='padding: 15px;' rowspan='1'>PHA Type</td>
                                    <td style ='padding: 15px;' rowspan='1'>Action Required</td>
                                    <td style ='padding: 15px;' rowspan='1'>Document Number</td>
                                    <td style ='padding: 15px;' rowspan='1'>Document Title</td>
                                    <td style ='padding: 15px;' rowspan='1'>Rev.</td>
                                    <td style ='padding: 15px;' rowspan='1'>Originator</td>
                                    <td style ='padding: 15px;' rowspan='1'>Received</td>
                                    <td style ='padding: 15px;' rowspan='1'>Due Date</td>
                                    <td style ='padding: 15px;' rowspan='1'>Remaining</td> 
                                    <td style ='padding: 15px;' rowspan='1'>Consolidator</td> 
                                </tr>
                            </thead>
                            <tbody style='color: #636363;background-color: #ffffff;border: 1px solid #dddfe1;'>");

                    int iNo = 1;
                    foreach (DataRow dr in dt.Select($"user_name='{responder_user_name}'"))
                    {
                        string url_reject_comment = "";
                        string url_reject = "";
                        string url_approver = "";

                        string doc_no = dr["document_number"]?.ToString() ?? "";

                        string background_color = "white";
                        string font_color = "black";
                        int iRemaining = 0;
                        bool action_status_close = (dr["remaining"]?.ToString() ?? "").ToLower() == "closed";

                        try
                        {
                            iRemaining = Convert.ToInt32(dr["remaining"].ToString());
                            if (iRemaining > 3)
                            {
                                background_color = "green"; font_color = "red";
                            }
                            else if (iRemaining > 0 && iRemaining < 3 && !action_status_close)
                            {
                                background_color = "yellow";
                            }
                            else if (iRemaining <= 0 && !action_status_close)
                            {
                                background_color = "red"; font_color = "white";
                            }
                        }
                        catch { }

                        using (Aes aesAlgorithm = Aes.Create())
                        {
                            aesAlgorithm.KeySize = 256;
                            aesAlgorithm.GenerateKey();
                            string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                            string plainText = $"seq={seq}&pha_no={doc_no}&step=3";
                            string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);
                            if (dr["pha_status"].ToString() == "13")
                            {
                                plainText = $"seq={seq}&pha_no={doc_no}&step=3";
                                cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                                url = server_url_by_action(sub_software) + cipherText + "&" + keyBase64 + "&" + vectorBase64;
                            }
                            else if (dr["pha_status"].ToString() == "21")
                            {
                                plainText = $"seq={seq}&pha_no={doc_no}&step=4&approver_type=required";
                                cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                                url = server_url_by_action(sub_software) + cipherText + "&" + keyBase64 + "&" + vectorBase64;

                                plainText = $"seq={seq}&pha_no={doc_no}&step=4&approver_type=reject";
                                cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                                url_reject = server_url_by_action(sub_software) + cipherText + "&" + keyBase64 + "&" + vectorBase64;

                                plainText = $"seq={seq}&pha_no={doc_no}&step=4&approver_type=reject_no_comment";
                                cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                                url_reject_comment = server_url_by_action(sub_software) + cipherText + "&" + keyBase64 + "&" + vectorBase64;

                                plainText = $"seq={seq}&pha_no={doc_no}&step=4&approver_type=approve";
                                cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                                url_approver = server_url_by_action(sub_software) + cipherText + "&" + keyBase64 + "&" + vectorBase64;
                            }
                        }

                        s_body.Append("<tr>");
                        s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", iNo);
                        s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", dr["pha_type"]);
                        s_body.AppendFormat("<td style ='padding: 15px;'><a href='{0}'>{1}</a></td>", url, dr["action_required"]);
                        s_body.AppendFormat("<td style ='padding: 15px;'><a href='{0}'>{1}</a></td>", url, dr["document_number"]);
                        s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", dr["document_title"]);
                        s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", dr["rev"]);
                        s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", dr["originator"]);
                        s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", dr["receivesd"]);
                        s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", dr["due_date"]);
                        s_body.AppendFormat("<td style ='padding: 15px; background-color:{0};color:{1};'>{2}</td>", background_color, font_color, dr["remaining"]);
                        s_body.AppendFormat("<td style ='padding: 15px;'>{0}</td>", dr["consolidator"]);
                        s_body.Append("</tr>");
                        iNo++;
                    }

                    s_body.Append("</tbody></table>");
                    s_body.Append("<br/><br/>The message of color assignment is as follow:");
                    s_body.Append("<table><tbody>");
                    s_body.Append("<tr><td style='width: 120px;padding:4px;background-color:green; color:red'><label>Green Color</label></td><td> : &gt; 3 days; this document has more than 3 days to complete your task</td></tr>");
                    s_body.Append("<tr><td style='width: 120px;padding:4px;background-color:yellow;'><label>Yellow Color</label></td><td> : &lt; 3 days; this document has less than 3 days to complete your task</td></tr>");
                    s_body.Append("<tr><td style='width: 120px;padding:4px;background-color:Red; color : white'><label>Red Color</label></td><td> : &lt;= 0 days; this document <label style='color:red'>is overdue, please urgent action</label></td></tr>");
                    s_body.Append("</tbody></table>");
                    s_body.AppendFormat("<br/><br/><a href='{0}'>Click here to access your Overall Tasks Window</a>", url_home_task);

                    s_body.Append("<br/><br/>DISCLAIMER:<br/>");
                    s_body.Append(@"This email message (including any attachment) is intended only for the personal use of the recipient(s) named above. It is confidential and may be legally privileged. If you are not an intended recipient, any use of this information is prohibited. If you have received this communication in error, please notify us immediately by email and delete the original message. In addition, we shall not be liable or responsible for any contents, including damages resulting from any virus transmitted by this email. Any information, comment, opinion, or statement contained in this email, including any attachments (if any), is that of the author only. Furthermore, this email (including any attachment) does not create any legally binding rights or obligations whatsoever, which may only be engaged and obliged by the exchange of hard copy documents signed by duly authorized representative(s of Thai Oil Public Company Limited.");
                    s_body.Append("</font></body></html>");

                    sendEmailModel data = new sendEmailModel
                    {
                        mail_subject = $"EPHA OUTSTANDING ACTION NOTIFICATION_{to_displayname}_{date_now}",
                        mail_body = s_body.ToString(),
                        mail_to = s_mail_to,
                        mail_cc = s_mail_cc,
                        mail_from = s_mail_from
                    };

                    if (!string.IsNullOrEmpty(file_ResponseSheet) && File.Exists(file_ResponseSheet))
                    {
                        data.mail_attachments = file_ResponseSheet;
                    }

                    msg = sendMail(data);
                    if (!string.IsNullOrEmpty(msg))
                    {
                        // Handle error
                    }
                }
            }
            #endregion mail to

            return msg;
        }

        public string MailNotificationReviewerReviewFollowup(string seq, string responder_user_name, string sub_software, Boolean responder_close_all)
        {
            string url = "";
            string url_home_task = "";
            string step_text = "Outstanding Action Notification";

            string to_displayname = "All";
            string s_mail_to = "";
            string s_mail_cc = "";
            string s_mail_from = "";


            string date_now = DateTime.Now.ToString("dd/MMM/yyyy");

            DataTable dt = new DataTable();
            string mail_admin_group = get_mail_admin_group();

            cls_conn = new ClassConnectionDb();
            ClassNoti classNoti = new ClassNoti();
            dt = classNoti.DataDailyByActionRequired_ReviewApprove(seq, responder_user_name, sub_software, false, responder_close_all);

            #region mail to
            s_mail_to = mail_admin_group;
            s_mail_cc = "";

            string msg = "";
            if (dt?.Rows.Count > 0)
            {
                for (int i = 0; i < dt?.Rows.Count; i++)
                {
                    s_mail_cc += (dt.Rows[i]["user_email"] + ";");

                    if (url_home_task == "")
                    {
                        using (Aes aesAlgorithm = Aes.Create())
                        {
                            aesAlgorithm.KeySize = 256;
                            aesAlgorithm.GenerateKey();
                            string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                            string pha_no = (dt.Rows[i]["document_number"] + "");

                            //insert keyBase64 to db 
                            string plainText = "seq=" + seq + "&pha_no=" + pha_no;
                            string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);
                            url_home_task = server_url_home_task + cipherText + "&" + keyBase64 + "&" + vectorBase64;
                        }
                    }
                }

                s_subject = "EPHA " + ("Outstanding Action Notification").ToString().ToUpper() + "_" + date_now;

                s_body = "<html><body><font face='tahoma' size='2'>";
                s_body += "Dear " + to_displayname + ",";

                s_body += @"<br/><br/>You have the following document(s) for action. Could you please proceed promptly.";
                s_body += @"<br/><br/><small style='color:red'>Note : For review action, ""Reviewer"" please response by reply this email within five working days prior auto proceed to next step.</small>";


                s_body += @"<br/>
                                <table style ='zoom: 100%;border-collapse: collapse;font-family: tahoma, geneva, sans-serif;background-color: #215289;color: #ffffff;font-weight: bold;font-size: 13px;border: 1px solid #54585d;'>   <thead>    
                                    <tr>
                                        <td style ='padding: 15px;' rowspan='1'>Task</td>
                                        <td style ='padding: 15px;' rowspan='1'>PHA Type</td>
                                        <td style ='padding: 15px;' rowspan='1'>Action Required</td>
                                        <td style ='padding: 15px;' rowspan='1'>Document Number</td>
                                        <td style ='padding: 15px;' rowspan='1'>Document Title</td>
                                        <td style ='padding: 15px;' rowspan='1'>Rev.</td>
                                        <td style ='padding: 15px;' rowspan='1'>Originator</td>
                                        <td style ='padding: 15px;' rowspan='1'>Received</td>
                                        <td style ='padding: 15px;' rowspan='1'>Due Date</td>
                                        <td style ='padding: 15px;' rowspan='1'>Remaining</td> 
                                        <td style ='padding: 15px;' rowspan='1'>Consolidator</td> 
                                    </tr>
                                </thead>
                                <tbody style='color: #636363;background-color: #ffffff;border: 1px solid #dddfe1;'> ";

                int iNo = 1;
                DataRow[] dr = dt.Select();
                for (int a = 0; a < dr.Length; a++)
                {
                    string doc_no = (dr[a]["document_number"] + "");

                    string background_color = "white";
                    string font_color = "black";
                    int iRemaining = 0;
                    Boolean action_status_close = (dr[a]["remaining"] + "").ToLower() == "closed";

                    try
                    {
                        iRemaining = Convert.ToInt32(dr[a]["remaining"] + "");
                        if (iRemaining > 3)
                        {
                            background_color = "green"; font_color = "red";
                        }
                        else if ((iRemaining > 0 && iRemaining < 3) && action_status_close == false)
                        {
                            background_color = "yellow";
                        }
                        else if (iRemaining <= 0 && action_status_close == false)
                        { background_color = "red"; font_color = "white"; }
                    }
                    catch { }

                    #region url  
                    url = "";
                    string url_def = "";
                    string url_approver = "";
                    string url_reject = "";
                    string url_reject_comment = "";
                    using (Aes aesAlgorithm = Aes.Create())
                    {
                        aesAlgorithm.KeySize = 256;
                        aesAlgorithm.GenerateKey();
                        string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                        //insert keyBase64 to db 
                        string plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=3";
                        string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);
                        if ((dr[a]["pha_status"] + "") == "13")
                        {
                            plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=3";
                            cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                            url = server_url_by_action(sub_software) + cipherText + "&" + keyBase64 + "&" + vectorBase64;
                        }
                        else if ((dr[a]["pha_status"] + "") == "21")
                        {
                            plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=4" + "&approver_type=required";
                            cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                            url = server_url_by_action(sub_software) + cipherText + "&" + keyBase64 + "&" + vectorBase64;

                            //reject 
                            plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=4" + "&approver_type=reject";
                            cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                            url_reject = server_url_by_action(sub_software) + cipherText + "&" + keyBase64 + "&" + vectorBase64;


                            //reject no comment
                            plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=4" + "&approver_type=reject_no_comment";
                            cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                            url_reject_comment = server_url_by_action(sub_software) + cipherText + "&" + keyBase64 + "&" + vectorBase64;

                            //approve
                            plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=4" + "&approver_type=approve";
                            cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                            url_approver = server_url_by_action(sub_software) + cipherText + "&" + keyBase64 + "&" + vectorBase64;
                        }

                    }
                    #endregion url 

                    s_body += "<tr>";
                    s_body += "<td style ='padding: 15px;'>" + (iNo) + "</td>";
                    s_body += "<td style ='padding: 15px;'>" + dr[a]["pha_type"] + "</td>";//hazop
                    if ((dr[a]["action_required"] + "").ToLower() == "recommendation closing"
                        || (dr[a]["action_required"] + "").ToLower() == "review"
                        || (dr[a]["action_required"] + "").ToLower() == "review & approve")
                    {
                        url_def = url;
                    }
                    else if ((dr[a]["action_required"] + "").ToLower() == "approve")
                    {
                        url_def = url;
                    }
                    s_body += "<td style ='padding: 15px;'><a href='" + url_def + "'>" + dr[a]["action_required"] + "</a></td>";//Recommendation Closing, Review, Approve

                    s_body += "<td style ='padding: 15px;'><a href='" + url_def + "'>" + dr[a]["document_number"] + "</a></td>";//hazop-2023-0000023
                    s_body += "<td style ='padding: 15px;'>" + dr[a]["document_title"] + "</td>";//xxmoc0003

                    s_body += "<td style ='padding: 15px;'>" + dr[a]["rev"] + "</td>";
                    s_body += "<td style ='padding: 15px;'>" + dr[a]["originator"] + "</td>";
                    s_body += "<td style ='padding: 15px;'>" + dr[a]["receivesd"] + "</td>";
                    s_body += "<td style ='padding: 15px;'>" + dr[a]["due_date"] + "</td>";
                    s_body += "<td style ='padding: 15px; background-color:" + background_color + ";color:" + font_color + "; '>" + dr[a]["remaining"] + "</td>";
                    s_body += "<td style ='padding: 15px;'>" + dr[a]["consolidator"] + "</td>";
                    s_body += "</tr>";
                    iNo += 1;

                }

                s_body += "</tbody>";
                s_body += "</table>";

                s_body += "<br/><br/>The message of color assignment is as follow:";

                s_body += "<table>";
                s_body += "<tbody>";
                s_body += "<tr>";
                s_body += "<td style='width: 120px;padding:4px;background-color:green; color:red'><label>Green Color</label></td>";
                s_body += "<td> : &gt; 3 days; this document has more than 3 days to complete your task</td>";
                s_body += "</tr>";
                s_body += "<tr>";
                s_body += "<td style='width: 120px;padding:4px;background-color:yellow;'><label>Yellow Color</label></td>";
                s_body += "<td> : &lt; 3 days; this document has less than 3 days to complete your task</td>";
                s_body += "</tr>";
                s_body += "<tr>";
                s_body += "<td style='width: 120px;padding:4px;background-color:Red; color : white'><label>Red Color</label></td>";
                s_body += "<td> : &lt;= 0 days; this document <label style='color:red'>is overdue, please urgent action</label></td>";
                s_body += "</tr>";
                s_body += "</tbody>";
                s_body += "</table>";
                //s_body += "<br/><label style='width: 120px;padding:4px;background-color:green; color:red'>Green Color</label> : &gt; 3 days; this document has more than 3 days to complete your task";
                //s_body += "<br/><label style='width: 120px;padding:4px;background-color:yellow;'>Yellow Color</label> : &lt; 3 days; this document has less than 3 days to complete your task";
                //s_body += "<br/><label style='width: 130px;padding:4px;background-color:Red; color : white'>Red Color</label> : &lt;= 0 days; this document <label style='color:red'>is overdue, please urgent action</label>";

                s_body += "<br/><br/><a href='" + url_home_task + "'>Click here to access your Overall Tasks Window</a>";

                s_body += "<br/><br/>DISCLAIMER:";
                s_body += "<br/>";
                s_body += @"This email message (including any attachment) is intended only for the personal use of the recipient(s) named above. It is confidential and may be legally privileged. If you are not an intended recipient, any use of this information is prohibited. If you have received this communication in error, please notify us immediately by email and delete the original message. In addition, we shall not be liable or responsible for any contents, including damages resulting from any virus transmitted by this email. Any information, comment, opinion, or statement contained in this email, including any attachments (if any), is that of the author only. Furthermore, this email (including any attachment) does not create any legally binding rights or obligations whatsoever, which may only be engaged and obliged by the exchange of hard copy documents signed by duly authorized representative(s of Thai Oil Public Company Limited.";
                s_body += "</font></body></html>";


                sendEmailModel data = new sendEmailModel();
                data.mail_subject = s_subject;
                data.mail_body = s_body;
                data.mail_to = s_mail_to;
                data.mail_cc = s_mail_cc;
                data.mail_from = s_mail_from;

                msg = sendMail(data);
                if (msg != "") { }

            }
            #endregion mail to

            return msg;


        }
        public string MailNotificationReviewerClosedAll(string seq, string sub_software)
        {
            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }

            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }

            #region call function export full report pdf 
            string _ram_type = sub_software switch
            {
                "hazop" => "5",
                "jsea" => "5",
                "whatif" => "5",
                _ => ""
            };
            ReportByWorksheetModel param = new ReportByWorksheetModel
            {
                seq = seq,
                export_type = "pdf",
                sub_software = sub_software,
                user_name = "",
                ram_type = _ram_type
            };
            ClassExcel clsExcel = new ClassExcel();
            string file_ResponseSheet = clsExcel.export_report_recommendation(param, false, true, true);

            if (string.IsNullOrEmpty(file_ResponseSheet))
            {
                return "Invalid file.";
            }
            else
            {
                // ตรวจสอบว่าไฟล์ที่ได้มามีอยู่จริงหรือไม่
                if (!File.Exists(file_ResponseSheet))
                {
                    return $"The specified file does not exist.{file_ResponseSheet}";
                }
            }
            #endregion call function export full report pdf 

            DataTable dt = new DataTable();
            string mail_admin_group = get_mail_admin_group();

            string url = "";
            string url_home_task = "";
            string step_text = "Outstanding Action Notification";

            string to_displayname = "All";
            string s_mail_to = "";
            string s_mail_cc = "";
            string s_mail_from = "";

            string date_now = DateTime.Now.ToString("dd/MMM/yyyy");

            cls_conn = new ClassConnectionDb();
            ClassNoti classNoti = new ClassNoti();
            dt = classNoti.DataDailyByActionRequired_Closed(seq, sub_software, false, false);

            #region mail to
            s_mail_to = mail_admin_group;
            s_mail_cc = "";

            string msg = "";
            if (dt?.Rows.Count > 0)
            {
                for (int i = 0; i < dt?.Rows.Count; i++)
                {
                    s_mail_cc += (dt.Rows[i]["user_email"] + ";");

                    if (url_home_task == "")
                    {
                        using (Aes aesAlgorithm = Aes.Create())
                        {
                            aesAlgorithm.KeySize = 256;
                            aesAlgorithm.GenerateKey();
                            string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                            string pha_no = (dt.Rows[i]["document_number"] + "");

                            //insert keyBase64 to db 
                            string plainText = "seq=" + seq + "&pha_no=" + pha_no;
                            string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);
                            url_home_task = server_url_home_task + cipherText + "&" + keyBase64 + "&" + vectorBase64;
                        }
                    }
                }

                s_subject = "EPHA " + ("Outstanding Action Notification").ToString().ToUpper() + "_" + date_now;

                s_body = "<html><body><font face='tahoma' size='2'>";
                s_body += "Dear " + to_displayname + ",";

                s_body += @"<br/><br/>You have the following document(s) for action. Could you please proceed promptly.";
                s_body += @"<br/><br/><small style='color:red'>Note : For review action, ""Reviewer"" please response by reply this email within five working days prior auto proceed to next step.</small>";


                s_body += @"<br/>
                                <table style ='zoom: 100%;border-collapse: collapse;font-family: tahoma, geneva, sans-serif;background-color: #215289;color: #ffffff;font-weight: bold;font-size: 13px;border: 1px solid #54585d;'>   <thead>    
                                    <tr>
                                        <td style ='padding: 15px;' rowspan='1'>Task</td>
                                        <td style ='padding: 15px;' rowspan='1'>PHA Type</td>
                                        <td style ='padding: 15px;' rowspan='1'>Action Required</td>
                                        <td style ='padding: 15px;' rowspan='1'>Document Number</td>
                                        <td style ='padding: 15px;' rowspan='1'>Document Title</td>
                                        <td style ='padding: 15px;' rowspan='1'>Rev.</td>
                                        <td style ='padding: 15px;' rowspan='1'>Originator</td>
                                        <td style ='padding: 15px;' rowspan='1'>Received</td>
                                        <td style ='padding: 15px;' rowspan='1'>Due Date</td>
                                        <td style ='padding: 15px;' rowspan='1'>Remaining</td> 
                                        <td style ='padding: 15px;' rowspan='1'>Consolidator</td> 
                                    </tr>
                                </thead>
                                <tbody style='color: #636363;background-color: #ffffff;border: 1px solid #dddfe1;'> ";

                int iNo = 1;
                DataRow[] dr = dt.Select();
                for (int a = 0; a < dr.Length; a++)
                {
                    string doc_no = (dr[a]["document_number"] + "");

                    string background_color = "white";
                    string font_color = "black";
                    int iRemaining = 0;
                    Boolean action_status_close = (dr[a]["remaining"] + "").ToLower() == "closed";


                    try
                    {
                        iRemaining = Convert.ToInt32(dr[a]["remaining"] + "");
                        if (iRemaining > 3)
                        {
                            background_color = "green"; font_color = "red";
                        }
                        else if ((iRemaining > 0 && iRemaining < 3) && action_status_close == false)
                        {
                            background_color = "yellow";
                        }
                        else if (iRemaining <= 0 && action_status_close == false)
                        { background_color = "red"; font_color = "white"; }
                    }
                    catch { }

                    #region url  
                    url = "";
                    string url_def = "";
                    string url_approver = "";
                    string url_reject = "";
                    string url_reject_comment = "";
                    using (Aes aesAlgorithm = Aes.Create())
                    {
                        aesAlgorithm.KeySize = 256;
                        aesAlgorithm.GenerateKey();
                        string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                        //insert keyBase64 to db 
                        string plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=3";
                        string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);
                        if ((dr[a]["pha_status"] + "") == "13")
                        {
                            plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=3";
                            cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                            url = server_url_by_action(sub_software) + cipherText + "&" + keyBase64 + "&" + vectorBase64;
                        }
                        else if ((dr[a]["pha_status"] + "") == "21")
                        {
                            plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=4" + "&approver_type=required";
                            cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                            url = server_url_by_action(sub_software) + cipherText + "&" + keyBase64 + "&" + vectorBase64;

                            //reject 
                            plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=4" + "&approver_type=reject";
                            cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                            url_reject = server_url_by_action(sub_software) + cipherText + "&" + keyBase64 + "&" + vectorBase64;


                            //reject no comment
                            plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=4" + "&approver_type=reject_no_comment";
                            cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                            url_reject_comment = server_url_by_action(sub_software) + cipherText + "&" + keyBase64 + "&" + vectorBase64;

                            //approve
                            plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=4" + "&approver_type=approve";
                            cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                            url_approver = server_url_by_action(sub_software) + cipherText + "&" + keyBase64 + "&" + vectorBase64;
                        }

                    }
                    #endregion url 

                    s_body += "<tr>";
                    s_body += "<td style ='padding: 15px;'>" + (iNo) + "</td>";
                    s_body += "<td style ='padding: 15px;'>" + dr[a]["pha_type"] + "</td>";//hazop

                    if ((dr[a]["action_required"] + "").ToLower() == "recommendation closing"
                        || (dr[a]["action_required"] + "").ToLower() == "review"
                        || (dr[a]["action_required"] + "").ToLower() == "review & approve")
                    {
                        url_def = url;
                    }
                    else if ((dr[a]["action_required"] + "").ToLower() == "approve")
                    {
                        url_def = url;
                    }
                    s_body += "<td style ='padding: 15px;'><a href='" + url_def + "'>" + dr[a]["action_required"] + "</a></td>";//Recommendation Closing, Review, Approve

                    s_body += "<td style ='padding: 15px;'><a href='" + url_def + "'>" + dr[a]["document_number"] + "</a></td>";//hazop-2023-0000023
                    s_body += "<td style ='padding: 15px;'>" + dr[a]["document_title"] + "</td>";//xxmoc0003

                    s_body += "<td style ='padding: 15px;'>" + dr[a]["rev"] + "</td>";
                    s_body += "<td style ='padding: 15px;'>" + dr[a]["originator"] + "</td>";
                    s_body += "<td style ='padding: 15px;'>" + dr[a]["receivesd"] + "</td>";
                    s_body += "<td style ='padding: 15px;'>" + dr[a]["due_date"] + "</td>";
                    s_body += "<td style ='padding: 15px; background-color:" + background_color + ";color:" + font_color + "; '>" + dr[a]["remaining"] + "</td>";
                    s_body += "<td style ='padding: 15px;'>" + dr[a]["consolidator"] + "</td>";
                    s_body += "</tr>";
                    iNo += 1;

                }

                s_body += "</tbody>";
                s_body += "</table>";

                s_body += "<br/><br/>The message of color assignment is as follow:";
                s_body += "<br/><label style='width: 120px;padding:4px;background-color:green; color:red'>Green Color</label> : &gt; 3 days; this document has more than 3 days to complete your task";
                s_body += "<br/><label style='width: 120px;padding:4px;background-color:yellow;'>Yellow Color</label> : &lt; 3 days; this document has less than 3 days to complete your task";
                s_body += "<br/><label style='width: 130px;padding:4px;background-color:Red; color : white'>Red Color</label> : &lt;= 0 days; this document <label style='color:red'>is overdue, please urgent action</label>";

                s_body += "<br/><br/><a href='" + url_home_task + "'>Click here to access your Overall Tasks Window</a>";

                s_body += "<br/><br/>DISCLAIMER:";
                s_body += "<br/>";
                s_body += @"This email message (including any attachment) is intended only for the personal use of the recipient(s) named above. It is confidential and may be legally privileged. If you are not an intended recipient, any use of this information is prohibited. If you have received this communication in error, please notify us immediately by email and delete the original message. In addition, we shall not be liable or responsible for any contents, including damages resulting from any virus transmitted by this email. Any information, comment, opinion, or statement contained in this email, including any attachments (if any), is that of the author only. Furthermore, this email (including any attachment) does not create any legally binding rights or obligations whatsoever, which may only be engaged and obliged by the exchange of hard copy documents signed by duly authorized representative(s of Thai Oil Public Company Limited.";
                s_body += "</font></body></html>";


                sendEmailModel data = new sendEmailModel();
                data.mail_subject = s_subject;
                data.mail_body = s_body;
                data.mail_to = s_mail_to;
                data.mail_cc = s_mail_cc;
                data.mail_from = s_mail_from;

                if (file_ResponseSheet != "")
                {
                    if (File.Exists(file_ResponseSheet))
                    {
                        data.mail_attachments = file_ResponseSheet;
                    }
                }
                msg = sendMail(data);
                if (msg != "")
                {
                    data.mail_attachments = null;
                    msg = sendMail(data);
                }

            }
            #endregion mail to

            return msg;
        }

        public string MailNotificationApproverTA3(string seq, string sub_software, string seq_approver, string user_approver_active)
        {
            string file_ResponseSheet = "";
            try
            {
                get_document_file_approver(ref file_ResponseSheet, seq_approver, user_approver_active);
            }
            catch { }

            string doc_no = "";
            string doc_name = "";
            string pha_sub_software = "";
            string meeting_date = "";
            string meeting_time = "";
            string reference_moc = "";
            string comment = "";
            string approve_status = "";
            string approver_displayname = "XXXXX (TOP-XX)";

            string url = "";

            string to_displayname = "All";
            string s_mail_to = "";
            string s_mail_cc = "";
            string s_mail_from = "";

            string mail_admin_group = get_mail_admin_group();

            ClassFunctions cls = new ClassFunctions();
            DataTable dt = new DataTable();
            List<SqlParameter> parameters = new List<SqlParameter>();

            if (sub_software == "hazop" || sub_software == "whatif")
            {
                // Approver ที่ Active
                sqlstr = @" select distinct h.pha_status, h.pha_no as pha_no,g.pha_request_name as pha_name,empre.user_email as request_email
                         , ta3.no 
                         , format(a.meeting_date, 'dd MMM yyyy') as meeting_date
                         , replace(a.meeting_start_time,'1/1/1970 ','') +' - '+ replace(a.meeting_end_time,'1/1/1970 ','') as meeting_time
                         , emp.user_displayname, emp.user_email, g.reference_moc
                         , lower(h.pha_sub_software) as pha_sub_software
                         , lower(ta3.user_name) as approver_user_name, ta3.approver_action_type 
                         from epha_t_header h
                         inner join EPHA_T_GENERAL g on lower(h.id) = lower(g.id_pha) 
                         inner join EPHA_T_SESSION a on lower(h.id) = lower(a.id_pha)  
                         inner join EPHA_T_APPROVER ta2 on lower(h.id) = lower(ta2.id_pha) and a.seq = ta2.id_session  
                         inner join EPHA_T_APPROVER_TA3 ta3 on lower(h.id) = lower(ta3.id_pha) and a.seq = ta3.id_session and ta2.id = ta3.id_approver    
                         inner join VW_EPHA_PERSON_DETAILS emp on lower(ta3.user_name)  = lower(emp.user_name) 
                         inner join VW_EPHA_PERSON_DETAILS empre on lower(h.pha_request_by) = lower(empre.user_name) 
                         inner join (select max(id) as id_session, id_pha from EPHA_T_SESSION group by id_pha ) s on h.id = s.id_pha and a.id = s.id_session and a.id_pha = s.id_pha  
                         where coalesce(ta3.approver_action_type,0) = 0 and h.id = @seq";
                sqlstr += @" and lower(ta2.user_name) like lower(@user_approver_active)";
                if (!string.IsNullOrEmpty(seq))
                {
                    parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar)
                    { Value = cls.ChkSqlNum(seq, "N")?.ToString() });
                }
                if (!string.IsNullOrEmpty(user_approver_active))
                {
                    parameters.Add(new SqlParameter("@user_approver_active", SqlDbType.VarChar)
                    { Value = cls.ChkSqlStr(user_approver_active ?? "", 100) });
                }
                sqlstr += @" order by convert(int,ta3.no) ";
            }

            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);

            if (dt?.Rows.Count > 0)
            {
                doc_no = (dt.Rows[0]["pha_no"] + "");
                doc_name = (dt.Rows[0]["pha_name"] + "");
                reference_moc = (dt.Rows[0]["reference_moc"] + "");

                approver_displayname = (dt.Rows[0]["user_displayname"] + "");

                pha_sub_software = (dt.Rows[0]["pha_sub_software"] + "");
                meeting_date = (dt.Rows[0]["meeting_date"] + "");
                meeting_time = (dt.Rows[0]["meeting_time"] + "");

                s_mail_to = (dt.Rows[0]["user_email"] + "");

                s_mail_cc = mail_admin_group;
                //cc originator
                if ((dt.Rows[0]["request_email"] + "") != "")
                {
                    s_mail_cc += ";" + (dt.Rows[0]["request_email"] + "");
                }
            }
            else { return "data is null"; }

            #region url  
            using (Aes aesAlgorithm = Aes.Create())
            {
                aesAlgorithm.KeySize = 256;
                aesAlgorithm.GenerateKey();
                string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                //insert keyBase64 to db 
                string plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=3";
                string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);

                url = server_url_by_action(sub_software) + cipherText + "&" + keyBase64 + "&" + vectorBase64;

            }
            #endregion url 

            s_subject = "EPHA : " + pha_sub_software.ToUpper() + ",Approval Responsibility Assigned to TA3.";

            s_body = "<html><body><font face='tahoma' size='2'>";
            s_body += "Dear " + approver_displayname + ",";

            s_body += "<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp";
            s_body += " Approval Responsibility Assigned to TA3";
            s_body += ", PHA No." + doc_no + " as details mentioned below,";

            s_body += "<br/><br/><b>Project Name</b> : " + doc_name;
            s_body += "<br/><b>Date</b> : " + meeting_date;
            s_body += "<br/><b>Time</b> : " + meeting_time;

            if (reference_moc != "") { s_body += "<br/><b>Reference MOC</b> : " + reference_moc; }


            s_body += "<br/><br/>";
            s_body += "More details of the study, <font color='red'; text-decoration: underline;><a href='" + url + "'> please click here</a></font>";

            s_body += "<br/><br/>DISCLAIMER:";
            s_body += "<br/>";
            s_body += @"This email message (including any attachment) is intended only for the personal use of the recipient(s) named above. It is confidential and may be legally privileged. If you are not an intended recipient, any use of this information is prohibited. If you have received this communication in error, please notify us immediately by email and delete the original message. In addition, we shall not be liable or responsible for any contents, including damages resulting from any virus transmitted by this email. Any information, comment, opinion, or statement contained in this email, including any attachments (if any), is that of the author only. Furthermore, this email (including any attachment) does not create any legally binding rights or obligations whatsoever, which may only be engaged and obliged by the exchange of hard copy documents signed by duly authorized representative(s of Thai Oil Public Company Limited.";
            s_body += "</font></body></html>";

            sendEmailModel data = new sendEmailModel();
            data.mail_subject = s_subject;
            data.mail_body = s_body;
            data.mail_to = s_mail_to;
            data.mail_cc = s_mail_cc;
            data.mail_from = s_mail_from;

            if (file_ResponseSheet != "")
            {
                //check path มาแล้ว 
                if (true)
                {
                    data.mail_attachments = file_ResponseSheet;
                }
            }


            return sendMail(data);

        }
        #endregion mail workflow last version


        #region file step approver

        private void get_document_file_approver(ref string file_ResponseSheet, string seq, string user_approver_active)
        {
            if (string.IsNullOrEmpty(seq)) { return; }
            if (string.IsNullOrEmpty(user_approver_active)) { return; }

            try
            {
                file_ResponseSheet = "";

                List<SqlParameter> parameters = new List<SqlParameter>();

                #region call function  export excel 
                sqlstr = @" select distinct da.document_file_path, lower(h.pha_sub_software) as sub_software
                         from epha_t_header h
                         inner join EPHA_T_GENERAL g on lower(h.id) = lower(g.id_pha) 
                         inner join EPHA_T_SESSION a on lower(h.id) = lower(a.id_pha)  
                         inner join EPHA_T_APPROVER ta2 on lower(h.id) = lower(ta2.id_pha) and a.seq = ta2.id_session  
                         inner join (select max(id) as id_session, id_pha from EPHA_T_SESSION group by id_pha ) s on h.id = s.id_pha and a.id = s.id_session and a.id_pha = s.id_pha  
                         inner join EPHA_T_DRAWING_APPROVER da on lower(h.id) = lower(da.id_pha) and ta2.id_session = da.id_session and ta2.seq = da.id_approver 
						 where h.request_approver = 1 and isnull(da.document_file_name,'') <>''
                         and h.id = @seq and lower(ta2.user_name) like lower(@user_approver_active) ";

                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                parameters.Add(new SqlParameter("@user_approver_active", SqlDbType.VarChar, 100) { Value = user_approver_active ?? "" });

                DataTable dtDrawing = new DataTable();
                dtDrawing = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
                for (int i = 0; i < dtDrawing?.Rows.Count; i++)
                {
                    //https://localhost:7098/AttachedFileTemp/hazop/HAZOP-2024-0000004-DRAWING-202401102132.PDF
                    //to
                    //D:\dotnet6-epha-api\dotnet6-epha-api/wwwroot/AttachedFileTemp/Hazop/HAZOP Report 202311281602.pdf

                    string sub_software = dtDrawing.Rows[i]["sub_software"]?.ToString() ?? "";
                    string fileTemp = dtDrawing.Rows[i]["document_file_path"]?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(sub_software) || !string.IsNullOrEmpty(fileTemp))
                    {
                        string file_fullpath_name = "";
                        string msg = ClassFile.check_file_other(sub_software, ref file_fullpath_name, fileTemp);
                        if (string.IsNullOrEmpty(msg) && !string.IsNullOrEmpty(file_fullpath_name))
                        {
                            if (file_ResponseSheet != "") { file_ResponseSheet += "|"; }
                            file_ResponseSheet += file_fullpath_name;
                        }
                    }
                }
                #endregion call function  export excel 
            }
            catch { }

        }
        #endregion file step approver
        public string MailNotificationChangeActionOwner(string seq, string sub_software, string seq_worksheet_list)
        {
            string url = "";
            string url_home_task = "";
            string step_text = "Outstanding Action Notification (Change Action Owner) ";

            string to_displayname = "All";
            string s_mail_to = "";
            string s_mail_cc = "";
            string s_mail_from = "";


            string date_now = DateTime.Now.ToString("dd/MMM/yyyy");

            DataTable dt = new DataTable();
            string mail_admin_group = get_mail_admin_group();

            cls_conn = new ClassConnectionDb();
            ClassNoti classNoti = new ClassNoti();
            dt = classNoti.DataDailyByActionRequired_Responder(seq, sub_software, false, seq_worksheet_list);

            #region mail to
            s_mail_to = "";
            s_mail_cc = mail_admin_group;

            string msg = "";
            if (dt?.Rows.Count > 0)
            {
                for (int i = 0; i < dt?.Rows.Count; i++)
                {
                    string user_name = (dt.Rows[i]["user_name"] + ";");
                    s_mail_to = (dt.Rows[i]["user_email"] + ";");

                    to_displayname = (dt.Rows[i]["user_displayname"] + "");

                    if (url_home_task == "")
                    {
                        using (Aes aesAlgorithm = Aes.Create())
                        {
                            aesAlgorithm.KeySize = 256;
                            aesAlgorithm.GenerateKey();
                            string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                            string pha_no = (dt.Rows[i]["document_number"] + "");

                            //insert keyBase64 to db 
                            string plainText = "seq=" + seq + "&pha_no=" + pha_no;
                            string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);
                            url_home_task = server_url_home_task + cipherText + "&" + keyBase64 + "&" + vectorBase64;
                        }
                    }

                    s_subject = "EPHA " + (step_text).ToString().ToUpper() + "_" + date_now;

                    s_body = "<html><body><font face='tahoma' size='2'>";
                    s_body += "Dear " + to_displayname + ",";

                    s_body += @"<br/><br/>You have the following document(s) for action. Could you please proceed promptly.";
                    s_body += @"<br/><br/><small style='color:red'>Note : For review action, ""Reviewer"" please response by reply this email within five working days prior auto proceed to next step.</small>";


                    s_body += @"<br/>
                                <table style ='zoom: 100%;border-collapse: collapse;font-family: tahoma, geneva, sans-serif;background-color: #215289;color: #ffffff;font-weight: bold;font-size: 13px;border: 1px solid #54585d;'>   <thead>    
                                    <tr>
                                        <td style ='padding: 15px;' rowspan='1'>Task</td>
                                        <td style ='padding: 15px;' rowspan='1'>PHA Type</td>
                                        <td style ='padding: 15px;' rowspan='1'>Action Required</td>
                                        <td style ='padding: 15px;' rowspan='1'>Document Number</td>
                                        <td style ='padding: 15px;' rowspan='1'>Document Title</td>
                                        <td style ='padding: 15px;' rowspan='1'>Rev.</td>
                                        <td style ='padding: 15px;' rowspan='1'>Originator</td>
                                        <td style ='padding: 15px;' rowspan='1'>Received</td>
                                        <td style ='padding: 15px;' rowspan='1'>Due Date</td>
                                        <td style ='padding: 15px;' rowspan='1'>Remaining</td> 
                                        <td style ='padding: 15px;' rowspan='1'>Consolidator</td> 
                                    </tr>
                                </thead>
                                <tbody style='color: #636363;background-color: #ffffff;border: 1px solid #dddfe1;'> ";

                    DataRow[] dr = dt.Select("user_name='" + user_name + "'");
                    for (int a = 0; a < dr.Length; a++)
                    {
                        string doc_no = (dr[a]["document_number"] + "");

                        string background_color = "white";
                        string font_color = "black";
                        int iRemaining = 0;
                        int iNo = 1;
                        Boolean action_status_close = (dr[a]["remaining"] + "").ToLower() == "closed";

                        try
                        {
                            iRemaining = Convert.ToInt32(dr[a]["remaining"] + "");
                            if (iRemaining > 3)
                            {
                                background_color = "green"; font_color = "red";
                            }
                            else if ((iRemaining > 0 && iRemaining < 3) && action_status_close == false)
                            {
                                background_color = "yellow";
                            }
                            else if (iRemaining <= 0 && action_status_close == false)
                            { background_color = "red"; font_color = "white"; }
                        }
                        catch { }

                        #region url  
                        url = "";
                        string url_def = "";
                        string url_approver = "";
                        string url_reject = "";
                        string url_reject_comment = "";
                        using (Aes aesAlgorithm = Aes.Create())
                        {
                            aesAlgorithm.KeySize = 256;
                            aesAlgorithm.GenerateKey();
                            string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                            //insert keyBase64 to db 
                            string plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=3";
                            string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);
                            if ((dr[a]["pha_status"] + "") == "13")
                            {
                                plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=3";
                                cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                                url = server_url_by_action(sub_software) + cipherText + "&" + keyBase64 + "&" + vectorBase64;
                            }
                            else if ((dr[a]["pha_status"] + "") == "21")
                            {
                                plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=4" + "&approver_type=required";
                                cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                                url = server_url_by_action(sub_software) + cipherText + "&" + keyBase64 + "&" + vectorBase64;

                                //reject 
                                plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=4" + "&approver_type=reject";
                                cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                                url_reject = server_url_by_action(sub_software) + cipherText + "&" + keyBase64 + "&" + vectorBase64;


                                //reject no comment
                                plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=4" + "&approver_type=reject_no_comment";
                                cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                                url_reject_comment = server_url_by_action(sub_software) + cipherText + "&" + keyBase64 + "&" + vectorBase64;

                                //approve
                                plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=4" + "&approver_type=approve";
                                cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                                url_approver = server_url_by_action(sub_software) + cipherText + "&" + keyBase64 + "&" + vectorBase64;
                            }

                        }
                        #endregion url 

                        s_body += "<tr>";
                        s_body += "<td style ='padding: 15px;'>" + (iNo) + "</td>";
                        s_body += "<td style ='padding: 15px;'>" + dr[a]["pha_type"] + "</td>";//hazop
                        if ((dr[a]["action_required"] + "").ToLower() == "recommendation closing"
                            || (dr[a]["action_required"] + "").ToLower() == "review"
                            || (dr[a]["action_required"] + "").ToLower() == "review & approve")
                        {
                            url_def = url;
                        }
                        else if ((dr[a]["action_required"] + "").ToLower() == "approve")
                        {
                            url_def = url;
                        }
                        s_body += "<td style ='padding: 15px;'><a href='" + url_def + "'>" + dr[a]["action_required"] + "</a></td>";//Recommendation Closing, Review, Approve

                        s_body += "<td style ='padding: 15px;'><a href='" + url_def + "'>" + dr[a]["document_number"] + "</a></td>";//hazop-2023-0000023
                        s_body += "<td style ='padding: 15px;'>" + dr[a]["document_title"] + "</td>";//xxmoc0003

                        s_body += "<td style ='padding: 15px;'>" + dr[a]["rev"] + "</td>";
                        s_body += "<td style ='padding: 15px;'>" + dr[a]["originator"] + "</td>";
                        s_body += "<td style ='padding: 15px;'>" + dr[a]["receivesd"] + "</td>";
                        s_body += "<td style ='padding: 15px;'>" + dr[a]["due_date"] + "</td>";
                        s_body += "<td style ='padding: 15px; background-color:" + background_color + ";color:" + font_color + "; '>" + dr[a]["remaining"] + "</td>";
                        s_body += "<td style ='padding: 15px;'>" + dr[a]["consolidator"] + "</td>";
                        s_body += "</tr>";
                        iNo += 1;

                    }

                    s_body += "</tbody>";
                    s_body += "</table>";

                    s_body += "<br/><br/>The message of color assignment is as follow:";

                    s_body += "<table>";
                    s_body += "<tbody>";
                    s_body += "<tr>";
                    s_body += "<td style='width: 120px;padding:4px;background-color:green; color:red'><label>Green Color</label></td>";
                    s_body += "<td> : &gt; 3 days; this document has more than 3 days to complete your task</td>";
                    s_body += "</tr>";
                    s_body += "<tr>";
                    s_body += "<td style='width: 120px;padding:4px;background-color:yellow;'><label>Yellow Color</label></td>";
                    s_body += "<td> : &lt; 3 days; this document has less than 3 days to complete your task</td>";
                    s_body += "</tr>";
                    s_body += "<tr>";
                    s_body += "<td style='width: 120px;padding:4px;background-color:Red; color : white'><label>Red Color</label></td>";
                    s_body += "<td> : &lt;= 0 days; this document <label style='color:red'>is overdue, please urgent action</label></td>";
                    s_body += "</tr>";
                    s_body += "</tbody>";
                    s_body += "</table>";
                    //s_body += "<br/><label style='width: 120px;padding:4px;background-color:green; color:red'>Green Color</label> : &gt; 3 days; this document has more than 3 days to complete your task";
                    //s_body += "<br/><label style='width: 120px;padding:4px;background-color:yellow;'>Yellow Color</label> : &lt; 3 days; this document has less than 3 days to complete your task";
                    //s_body += "<br/><label style='width: 130px;padding:4px;background-color:Red; color : white'>Red Color</label> : &lt;= 0 days; this document <label style='color:red'>is overdue, please urgent action</label>";

                    s_body += "<br/><br/><a href='" + url_home_task + "'>Click here to access your Overall Tasks Window</a>";

                    s_body += "<br/><br/>DISCLAIMER:";
                    s_body += "<br/>";
                    s_body += @"This email message (including any attachment) is intended only for the personal use of the recipient(s) named above. It is confidential and may be legally privileged. If you are not an intended recipient, any use of this information is prohibited. If you have received this communication in error, please notify us immediately by email and delete the original message. In addition, we shall not be liable or responsible for any contents, including damages resulting from any virus transmitted by this email. Any information, comment, opinion, or statement contained in this email, including any attachments (if any), is that of the author only. Furthermore, this email (including any attachment) does not create any legally binding rights or obligations whatsoever, which may only be engaged and obliged by the exchange of hard copy documents signed by duly authorized representative(s of Thai Oil Public Company Limited.";
                    s_body += "</font></body></html>";

                    sendEmailModel data = new sendEmailModel();
                    data.mail_subject = s_subject;
                    data.mail_body = s_body;
                    data.mail_to = s_mail_to;
                    data.mail_cc = s_mail_cc;
                    data.mail_from = s_mail_from;

                    msg = sendMail(data);
                    if (msg != "") { }
                }
            }
            #endregion mail to

            return msg;


        }

    }
}
