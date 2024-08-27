using Class;
using Microsoft.AspNetCore.Mvc;
using Model;
using Newtonsoft.Json;
using System.Data;
using System.Data.SqlClient;
using System.Transactions;
using dotnet6_epha_api.Class;
using dotnet_epha_api.Class; 

namespace Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    [IgnoreAntiforgeryToken] // ข้ามการตรวจสอบ CSRF
    public class FlowController : ControllerBase
    {
        [HttpPost("ClearDataTableTransactions", Name = "ClearDataTableTransactions")]
        //[ValidateAntiForgeryToken]
        private string ReadTransactionLog(string token_log)
        {
            ClassTransactionLog clstranlog = new ClassTransactionLog();
            return clstranlog.load_log(token_log);
        }
        private string InsertTransactionLog(string module, string sub_software, Object param, ref string token_log)
        {
            try
            {
                string param_text = JsonConvert.SerializeObject(param, Formatting.Indented) ?? "";
                int maxLength = 4000; // Set this to the maximum length your column supports
                if (param_text.Length > maxLength)
                {
                    param_text = param_text.Substring(0, maxLength);
                }

                //insert log    
                ClassTransactionLog clstranlog = new ClassTransactionLog();
                string msg = clstranlog.insert_log(module, sub_software, param_text, ref token_log);
                if ((msg?.ToString() == ""))
                {
                    msg = "error insert log : " + token_log;
                }
                return msg;
            }
            catch (Exception ex) { return ex.Message.ToString(); }
        }

        [HttpPost("ClearDataTableTransactions", Name = "ClearDataTableTransactions")]
        //[ValidateAntiForgeryToken]
        public string ClearDataTableTransactions()
        {
            string ret = "";
            List<SqlParameter> parameter = new List<SqlParameter>();
            string sqlstr = "SELECT name FROM SYSOBJECTS where lower(name) like 'epha_t%' ";
            DataTable dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameter);

            try
            {
                using (TransactionScope scope = new TransactionScope())
                using (SqlConnection conn = new SqlConnection(ClassConnectionDb.ConnectionString()))
                {
                    conn.Open();

                    foreach (DataRow row in dt.Rows)
                    {
                        sqlstr = "DELETE FROM " + row["name"];

                        using (SqlCommand cmd = new SqlCommand(sqlstr, conn))
                        {
                            ret = cmd.ExecuteNonQuery() > 0 ? "true" : "false";
                            if (ret.ToLower() != "true") throw new Exception("Failed to delete data from table: " + row["name"]);
                        }
                    }

                    scope.Complete();
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message + " - sqlstr: " + sqlstr;
            }

            return ret;
        }

        [HttpPost("ConnectionSting", Name = "ConnectionSting")]
        ////[ValidateAntiForgeryToken]
        public string ConnectionSting()
        {
            // ดึง ConnectionString จาก appsettings.json
            string connStrSQL = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json")
                .Build()
                .GetSection("ConnectionConfig")["ConnString"] ?? "";

            // ส่งค่า ConnectionString กลับในรูปแบบ JSON
            return connStrSQL;
        }

        [HttpPost("ExQueryString", Name = "ExQueryString")]
        //[ValidateAntiForgeryToken]
        public string ExQueryString(string param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("ExQueryString", "", param, ref token_log);
            try
            {
                ClassConnectionDb cls_conn = new ClassConnectionDb();
                System.Data.DataTable dt = new System.Data.DataTable();

                string sqlstr = param ?? "";
                dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

                string json = Newtonsoft.Json.JsonConvert.SerializeObject(dt, Newtonsoft.Json.Formatting.Indented);

                return json;
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        //[HttpPost("copy_document_file_responder_to_reviewer", Name = "copy_document_file_responder_to_reviewer")]
        ////[ValidateAntiForgeryToken]
        //public string copy_document_file_responder_to_reviewer(LoadDocModel param)
        //{
        //    string token_log = "";
        //    string msg = InsertTransactionLog("copy_document_file_responder_to_reviewer", "", param, ref token_log);
        //    try
        //    {
        //        string token_doc = (param.token_doc ?? "");
        //        string sub_software = (param.sub_software ?? "");
        //        string user_name = (param.user_name ?? "");
        //        ClassHazopSet cls = new ClassHazopSet();
        //        return cls.copy_document_file_responder_to_reviewer(user_name, token_doc, sub_software);
        //    }
        //    catch (Exception e)
        //    {
        //        msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
        //    }
        //    return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        //}

        [HttpPost("importfile_data_jsea", Name = "importfile_data_jsea")]
        //[ValidateAntiForgeryToken]
        public string importfile_data_jsea([FromForm] uploadFile param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("importfile_data_jsea", "jsea", param, ref token_log);
            try
            {
                // กำหนดชนิดไฟล์ที่อนุญาต
                string[] allowedExtensions = { ".xlsx", ".xls", ".pdf", ".doc", ".docx", ".png", ".jpg", ".gif", ".eml", ".msg" };

                // ตรวจสอบไฟล์แต่ละไฟล์ที่อัพโหลด
                foreach (var file in param.file_obj)
                {
                    var extension = Path.GetExtension(file.FileName).ToLowerInvariant();
                    if (!allowedExtensions.Contains(extension))
                    {
                        // ถ้าไฟล์มีชนิดที่ไม่ได้รับอนุญาต ให้คืนค่าข้อความผิดพลาด
                        return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave("false", "File type not allowed.", "", "", "", ""));
                    }
                }

                string sub_software = "jsea";
                try { sub_software = param?.sub_software ?? ""; } catch { }

                ClassHazopSet cls = new ClassHazopSet();
                return cls.importfile_data_jsea(param, sub_software);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));


        }
        [HttpPost("uploadfile_data", Name = "uploadfile_data")]
        //[ValidateAntiForgeryToken]
        public string uploadfile_data([FromForm] uploadFile param)
        {
            string sub_software = "hazop";
            try { sub_software = param?.sub_software ?? ""; } catch { }

            string token_log = "";
            string msg = InsertTransactionLog("uploadfile_data", sub_software, param, ref token_log);
            try
            {
                // กำหนดชนิดไฟล์ที่อนุญาต
                string[] allowedExtensions = { ".xlsx", ".xls", ".pdf", ".doc", ".docx", ".png", ".jpg", ".gif", ".eml", ".msg" };

                // ตรวจสอบไฟล์แต่ละไฟล์ที่อัพโหลด
                foreach (var file in param.file_obj)
                {
                    var extension = Path.GetExtension(file.FileName).ToLowerInvariant();
                    if (!allowedExtensions.Contains(extension))
                    {
                        // ถ้าไฟล์มีชนิดที่ไม่ได้รับอนุญาต ให้คืนค่าข้อความผิดพลาด
                        return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave("false", "File type not allowed.", "", "", "", ""));
                    }
                }

                ClassHazopSet cls = new ClassHazopSet();
                return cls.uploadfile_data(param, sub_software);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }

        [HttpPost("uploadfile_data_followup", Name = "uploadfile_data_followup")]
        //[ValidateAntiForgeryToken]
        public string uploadfile_data_followup([FromForm] uploadFile param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("uploadfile_data_followup", "followup", param, ref token_log);

            try
            {
                // กำหนดชนิดไฟล์ที่อนุญาต
                string[] allowedExtensions = { ".xlsx", ".xls", ".pdf", ".doc", ".docx", ".png", ".jpg", ".gif", ".eml", ".msg" };

                // ตรวจสอบไฟล์แต่ละไฟล์ที่อัพโหลด
                foreach (var file in param.file_obj)
                {
                    var extension = Path.GetExtension(file.FileName).ToLowerInvariant();
                    if (!allowedExtensions.Contains(extension))
                    {
                        // ถ้าไฟล์มีชนิดที่ไม่ได้รับอนุญาต ให้คืนค่าข้อความผิดพลาด
                        return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave("false", "File type not allowed.", "", "", "", ""));
                    }
                }

                // เรียกใช้ฟังก์ชัน uploadfile_data ถ้าผ่านการตรวจสอบแล้ว
                ClassHazopSet cls = new ClassHazopSet();
                return cls.uploadfile_data(param, "followup");
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }


        [HttpPost("get_hazop_details", Name = "get_hazop_details")]
        ////[ValidateAntiForgeryToken]
        public string get_hazop_details(LoadDocModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("get_hazop_details", "hazop", param, ref token_log);
            try
            {
                param.sub_software = "hazop";
                ClassHazop cls = new ClassHazop();
                return cls.get_details(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));


        }
        [HttpPost("get_jsea_details", Name = "get_jsea_details")]
        //[ValidateAntiForgeryToken]
        public string get_jsea_details(LoadDocModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("get_jsea_details", "jsea", param, ref token_log);
            try
            {
                param.sub_software = "jsea";
                ClassHazop cls = new ClassHazop();
                return cls.get_details(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }
        [HttpPost("get_max_id", Name = "get_max_id")]
        //[ValidateAntiForgeryToken]
        public string get_max_id(string table_name = "", string id_pha = "")
        {
            table_name = table_name ?? "epha_t_header";
            id_pha = id_pha ?? "";

            ClassHazop cls = new ClassHazop();
            string token_log = "";
            string msg = "";
            try
            {
                int id_pha_ref = (cls.get_max(table_name, id_pha));
                return id_pha_ref.ToString() ?? "";
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }
        [HttpPost("get_whatif_details", Name = "get_whatif_details")]
        //[ValidateAntiForgeryToken]
        public string get_whatif_details(LoadDocModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("get_whatif_details", "whatif", param, ref token_log);
            try
            {
                param.sub_software = "whatif";
                ClassHazop cls = new ClassHazop();
                return cls.get_details(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }
        [HttpPost("get_hra_details", Name = "get_hra_details")]
        //[ValidateAntiForgeryToken]
        public string get_hra_details(LoadDocModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("get_hra_details", "hra", param, ref token_log);
            try
            {
                param.sub_software = "hra";
                ClassHazop cls = new ClassHazop();
                return cls.get_details(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }

        [HttpPost("load_page_search_details", Name = "load_page_search_details")]
        //[ValidateAntiForgeryToken]
        public string load_page_search_details(LoadDocModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("load_page_search_details", "all", param, ref token_log);
            try
            {
                ClassHazop cls = new ClassHazop();
                return cls.get_search_details(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }

        [HttpPost("set_hazop", Name = "set_hazop")]
        //[ValidateAntiForgeryToken]
        public string set_hazop(SetDataWorkflowModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("set_hazop", "hazop", param, ref token_log);
            try
            {
                ClassHazopSet cls = new ClassHazopSet();
                param.sub_software = "hazop";
                //param.flow_action = "change_action_owner";

                //20240419 เพิ่ม flow_action : change_action_owner --> เป็นการแก้ไขรายชื่อ action owner 
                if (param.flow_action == "change_action_owner") { return cls.set_workflow_change_employee(param); }

                //20240516 เพิ่ม flow_action : change_approver --> เป็นการแก้ไขรายชื่อ approver 
                if (param.flow_action == "change_approver") { return cls.set_workflow_change_employee(param); }


                return cls.set_workflow(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }
        [HttpPost("set_jsea", Name = "set_jsea")]
        //[ValidateAntiForgeryToken]
        public string set_jsea(SetDataWorkflowModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("set_jsea", "jsea", param, ref token_log);
            try
            {
                ClassHazopSet cls = new ClassHazopSet();
                param.sub_software = "jsea";
                return cls.set_workflow(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }

        [HttpPost("set_whatif", Name = "set_whatif")]
        //[ValidateAntiForgeryToken]
        public string set_whatif(SetDataWorkflowModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("set_whatif", "whatif", param, ref token_log);
            try
            {
                ClassHazopSet cls = new ClassHazopSet();
                param.sub_software = "whatif";

                //20240419 เพิ่ม flow_action : change_action_owner --> เป็นการแก้ไขรายชื่อ action owner 
                if (param.flow_action == "change_action_owner") { return cls.set_workflow_change_employee(param); }

                //20240516 เพิ่ม flow_action : change_approver --> เป็นการแก้ไขรายชื่อ approver 
                if (param.flow_action == "change_approver") { return cls.set_workflow_change_employee(param); }

                return cls.set_workflow(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }
        [HttpPost("set_hra", Name = "set_hra")]
        //[ValidateAntiForgeryToken]
        public string set_hra(SetDataWorkflowModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("set_hra", "hra", param, ref token_log);
            try
            {
                ClassHazopSet cls = new ClassHazopSet();
                param.sub_software = "hra";

                //20240419 เพิ่ม flow_action : change_action_owner --> เป็นการแก้ไขรายชื่อ action owner 
                if (param.flow_action == "change_action_owner") { return cls.set_workflow_change_employee(param); }

                //20240516 เพิ่ม flow_action : change_approver --> เป็นการแก้ไขรายชื่อ approver 
                if (param.flow_action == "change_approver") { return cls.set_workflow_change_employee(param); }

                return cls.set_workflow(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }

        [HttpPost("set_master_ram", Name = "set_master_ram")]
        //[ValidateAntiForgeryToken]
        public string set_master_ram(SetDataWorkflowModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("set_master_ram", "all", param, ref token_log);
            try
            {
                ClassHazopSet cls = new ClassHazopSet();
                return cls.set_master_ram(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }
        [HttpPost("edit_worksheet", Name = "edit_worksheet")]
        //[ValidateAntiForgeryToken]
        public string edit_worksheet(SetDocWorksheetModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("edit_worksheet", (param?.sub_software ?? "all"), param, ref token_log);
            try
            {
                ClassHazopSet cls = new ClassHazopSet();
                return cls.edit_worksheet(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }
        [HttpPost("set_approve", Name = "set_approve")]
        //[ValidateAntiForgeryToken]
        public string set_approve(SetDocApproveModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("set_approve", (param?.sub_software ?? "all"), param, ref token_log);
            try
            {
                ClassHazopSet cls = new ClassHazopSet();
                return cls.set_approve(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }
        [HttpPost("set_transfer_monitoring", Name = "set_transfer_monitoring")]
        //[ValidateAntiForgeryToken]
        public string set_transfer_monitoring(SetDocTransferMonitoringModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("set_transfer_monitoring", (param?.sub_software ?? "all"), param, ref token_log);
            try
            {
                ClassHazopSet cls = new ClassHazopSet();
                return cls.set_transfer_monitoring(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }
        [HttpPost("set_approve_ta3", Name = "set_approve_ta3")]
        //[ValidateAntiForgeryToken]
        public string set_approve_ta3(SetDocApproveTa3Model param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("set_approve_ta3", (param?.sub_software ?? "all"), param, ref token_log);
            try
            {
                ClassHazopSet cls = new ClassHazopSet();
                return cls.set_approve_ta3(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }
        #region Mail
        [HttpPost("MailToPHAConduct", Name = "MailToPHAConduct")]
        //[ValidateAntiForgeryToken]
        public string MailToPHAConduct(string seq, string sub_software)
        {
            string token_log = "";
            string msg = InsertTransactionLog("MailToPHAConduct", "all", new object(), ref token_log);
            try
            {
                ClassEmail cls = new ClassEmail();
                return cls.MailNotificationWorkshopInvitation(seq, sub_software);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }
        [HttpPost("MailToActionOwner", Name = "MailToActionOwner")]
        //[ValidateAntiForgeryToken]
        public string MailToActionOwner(string seq, string sub_software)
        {
            string token_log = "";
            string msg = InsertTransactionLog("MailToActionOwner", "all", new object(), ref token_log);
            try
            {
                ClassEmail cls = new ClassEmail();
                return cls.MailToActionOwner(seq, sub_software);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }


        [HttpPost("load_notification", Name = "load_notification")]
        //[ValidateAntiForgeryToken]
        public string load_notification(LoadDocModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("load_notification", "all", new object(), ref token_log);
            try
            {
                ClassHazop cls = new ClassHazop();
                return cls.get_notification(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }
        [HttpPost("send_notification_member_review", Name = "send_notification_member_review")]
        //[ValidateAntiForgeryToken]
        public string send_notification_member_review(SetDocModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("send_notification_member_review", "all", new object(), ref token_log);
            try
            {
                ClassEmail cls = new ClassEmail();
                string id_session = "";
                string ret = cls.MailNotificationMemberReview((param.pha_seq + ""), (param.sub_software + ""));

                return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave(ret, "Msg :" + (ret == "" ? "true" : ret) + ",Session Last :" + id_session, "", "", "", ""));
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }
        [HttpPost("send_notification_daily", Name = "send_notification_daily")]
        //[ValidateAntiForgeryToken]
        public string send_notification_daily(SetDocModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("send_notification_daily", "all", new object(), ref token_log);
            try
            {
                ClassEmail cls = new ClassEmail();
                string ret = cls.MailNotificationOutstandingAction((param.user_name + ""), (param.pha_seq + ""), (param.sub_software + ""));

                return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave(ret, "Msg :" + (ret == "" ? "true" : ret), "", "", "", ""));
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }
        #endregion Mail

        #region follow up  
        [HttpPost("load_follow_up", Name = "load_follow_up")]
        //[ValidateAntiForgeryToken]
        public string load_follow_up(LoadDocModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("load_follow_up", "follow_up", param, ref token_log);
            try
            {
                ClassHazop cls = new ClassHazop();
                return cls.get_followup(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }
        [HttpPost("load_follow_up_details", Name = "load_follow_up_details")]
        //[ValidateAntiForgeryToken]
        public string load_follow_up_details(LoadDocFollowModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("load_follow_up_details", "follow_up", param, ref token_log);
            try
            {
                ClassHazop cls = new ClassHazop();
                return cls.get_followup_detail(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }
        [HttpPost("set_follow_up", Name = "set_follow_up")]
        //[ValidateAntiForgeryToken]
        public string set_follow_up(SetDataWorkflowModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("set_follow_up", "follow_up", param, ref token_log);
            try
            {
                ClassHazopSet cls = new ClassHazopSet();
                return cls.set_follow_up(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }
        [HttpPost("set_follow_up_review", Name = "set_follow_up_review")]
        //[ValidateAntiForgeryToken]
        public string set_follow_up_review(SetDataWorkflowModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("set_follow_up_review", "follow_up_review", param, ref token_log);
            try
            {
                ClassHazopSet cls = new ClassHazopSet();
                return cls.set_follow_up_review(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }
        #endregion follow up 

        #region home tasks
        [HttpPost("load_home_tasks", Name = "load_home_tasks")]
        //[ValidateAntiForgeryToken]
        public string load_home_tasks(LoadDocModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("load_home_tasks", "home_tasks", param, ref token_log);
            try
            {
                ClassHazop cls = new ClassHazop();
                return cls.get_hometasks(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));


        }
        #endregion home tasks

        #region export hazop 
        [HttpPost("export_hazop_report", Name = "export_hazop_report")]
        //[ValidateAntiForgeryToken]
        public string export_hazop_report(ReportModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("export_hazop_report", "hazop", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                return cls.export_full_report(param, true);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        [HttpPost("export_hazop_worksheet", Name = "export_hazop_worksheet")]
        //[ValidateAntiForgeryToken]
        public string export_hazop_worksheet(ReportModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("export_hazop_worksheet", "hazop", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                //return cls.export_hazop_worksheet(param);
                return cls.export_full_report(param, false);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        [HttpPost("export_hazop_recommendation", Name = "export_hazop_recommendation")]
        //[ValidateAntiForgeryToken]
        public string export_hazop_recommendation(ReportByWorksheetModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("export_hazop_recommendation", "hazop", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                return cls.export_report_recommendation(param, false, false);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        [HttpPost("export_hazop_ram", Name = "export_hazop_ram")]
        //[ValidateAntiForgeryToken]
        public string export_hazop_ram(ReportModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("export_hazop_ram", "hazop", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                return cls.export_template_ram(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }
        [HttpPost("export_hazop_guidewords", Name = "export_hazop_guidewords")]
        //[ValidateAntiForgeryToken]
        public string export_hazop_guidewords(ReportModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("export_hazop_guidewords", "hazop", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                return cls.export_hazop_guidewords(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        #endregion export hazop


        #region export what's if 
        [HttpPost("export_template_whatif", Name = "export_template_whatif")]
        //[ValidateAntiForgeryToken]
        public string export_template_whatif(ReportModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("export_template_whatif", "whatif", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                return cls.export_full_report(param, true);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        [HttpPost("export_whatif_report", Name = "export_whatif_report")]
        //[ValidateAntiForgeryToken]
        public string export_whatif_report(ReportModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("export_whatif_report", "whatif", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                return cls.export_full_report(param, true);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));


        }

        [HttpPost("export_whatif_worksheet", Name = "export_whatif_worksheet")]
        //[ValidateAntiForgeryToken]
        public string export_whatif_worksheet(ReportModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("export_whatif_worksheet", "whatif", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                return cls.export_full_report(param, false);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        [HttpPost("export_whatif_recommendation", Name = "export_whatif_recommendation")]
        //[ValidateAntiForgeryToken]
        public string export_whatif_recommendation(ReportByWorksheetModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("export_whatif_recommendation", "whatif", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                return cls.export_report_recommendation(param, false, false);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        [HttpPost("export_whatif_ram", Name = "export_whatif_ram")]
        //[ValidateAntiForgeryToken]
        public string export_whatif_ram(ReportModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("export_whatif_ram", "whatif", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                return cls.export_template_ram(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }
        #endregion export what's if

        #region export jsea

        [HttpPost("export_template_jsea", Name = "export_template_jsea")]
        //[ValidateAntiForgeryToken]
        public string export_template_jsea(ReportModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("export_template_jsea", "jsea", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                //return cls.export_template_jsea(param);
                return cls.export_full_report(param, true);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        [HttpPost("export_jsea_report", Name = "export_jsea_report")]
        //[ValidateAntiForgeryToken]
        public string export_jsea_report(ReportModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("export_jsea_report", "jsea", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                return cls.export_full_report(param, true);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));


        }

        [HttpPost("export_jsea_worksheet", Name = "export_jsea_worksheet")]
        //[ValidateAntiForgeryToken]
        public string export_jsea_worksheet(ReportModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("export_jsea_worksheet", "jsea", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                //return cls.export_jsea_worksheet(param);
                return cls.export_full_report(param, false);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        #endregion export jsea

        #region export hra 
        [HttpPost("export_hra_report", Name = "export_hra_report")]
        //[ValidateAntiForgeryToken]
        public string export_hra_report(ReportModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("export_hra_report", "hra", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                return cls.export_full_report(param, true);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        [HttpPost("export_hra_worksheet", Name = "export_hra_worksheet")]
        //[ValidateAntiForgeryToken]
        public string export_hra_worksheet(ReportModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("export_hra_worksheet", "hra", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                return cls.export_potential_health_checklist_template(param, false);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        [HttpPost("export_hra_recommendation", Name = "export_hra_recommendation")]
        //[ValidateAntiForgeryToken]
        public string export_hra_recommendation(ReportByWorksheetModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("export_hra_recommendation", "hra", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                return cls.export_report_recommendation(param, false, false);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }

        [HttpPost("export_hra_template_moc", Name = "export_hra_template_moc")]
        //[ValidateAntiForgeryToken]
        public string export_hra_template_moc(ReportModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("export_hra_template_moc", "hra", param, ref token_log);
            try
            {
                if (param != null)
                {
                    param.export_type = "template";

                    ClassExcel cls = new ClassExcel();
                    return cls.export_full_report(param, false);
                }
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }
        #endregion export hra

        #region other case function 
        [HttpPost("MailNotificationDaily", Name = "MailNotificationDaily")]
        //[ValidateAntiForgeryToken]
        public string MailNotificationDaily(LoadDocModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("MailNotificationDaily", "noti", param, ref token_log);
            try
            {
                string user_name = (param.user_name + "");
                string seq = (param.token_doc + "");
                string sub_software = (param.sub_software + "");

                ClassEmail classEmail = new ClassEmail();
                return classEmail.MailNotificationOutstandingAction(user_name, seq, sub_software);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }


        [HttpPost("copy_pdf_file", Name = "copy_pdf_file")]
        //[ValidateAntiForgeryToken]
        public string copy_pdf_file(CopyFileModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("", "copy_pdf_file", param, ref token_log);
            try
            {
                string _file_fullpath_name = param.file_path ?? "";
                string _folder = param.sub_software ?? "hazop";
                if (string.IsNullOrEmpty(_file_fullpath_name) || string.IsNullOrEmpty(_folder))
                {
                    if (string.IsNullOrEmpty(ClassFile.check_file_on_server(_folder, _file_fullpath_name)))
                    {
                        ClassExcel classExcel = new ClassExcel();
                        return classExcel.copy_pdf_file(param);
                    }
                    else { msg = "The file is not within the allowed directory."; }
                }
                else
                {
                    msg = "Invalid file path/folder.";
                }
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }

        #endregion other case function

        #region Function Search
        [HttpPost("employees_search", Name = "employees_search")]
        public string employees_search(EmployeeModel param)
        {
            //param.max_rows = (param.max_rows == null ? "10" : param.max_rows); 
            ClassHazop cls = new ClassHazop();
            return cls.employees_search(param);

        }
        [HttpPost("employees_list", Name = "employees_list")]
        public string employees_list(EmployeeListModel param)
        {
            ClassHazop cls = new ClassHazop();
            return cls.employees_list(param);

        }
        #endregion  Function Search


        #region Function Manage Document

        [HttpPost("manage_document_copy", Name = "manage_document_copy")]
        public string manage_document_copy(ManageDocModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("manage_document_copy", "all", param, ref token_log);
            try
            {
                ClassManage cls = new ClassManage();
                string ret = cls.DocumentCopy(param);

                return ret;
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }
        [HttpPost("manage_document_cancel", Name = "manage_document_cancel")]
        public string manage_document_cancel(ManageDocModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("manage_document_cancel", "all", param, ref token_log);
            try
            {
                ClassManage cls = new ClassManage();
                string ret = cls.DocumentCancel(param);

                return ret;
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }
        #endregion Function Manage Document


        [HttpPost("export_recommendation_by_action_owner", Name = "export_recommendation_by_action_owner")]
        //[ValidateAntiForgeryToken]
        public string export_recommendation_by_action_owner(ReportByWorksheetModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("export_recommendation_by_action_owner", "all", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                msg = cls.export_report_recommendation(param, true, false);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        [HttpPost("export_recommendation_by_item", Name = "export_recommendation_by_item")]
        //[ValidateAntiForgeryToken]
        public string export_recommendation_by_item(ReportByWorksheetModel param)
        {
            string token_log = "";
            string msg = InsertTransactionLog("export_recommendation_by_item", "all", param, ref token_log);
            try
            {
                ClassExcel cls = new ClassExcel();
                msg = cls.export_report_recommendation(param, false, false);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

    }
}
