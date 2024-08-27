using DocumentFormat.OpenXml.Spreadsheet;
using dotnet6_epha_api.Class;
using Model;
using System.Data;
using System.Data.SqlClient;
using System.DirectoryServices;

namespace Class
{

    public class ClassLogin
    {
        string sqlstr = "";
        ClassFunctions cls = new ClassFunctions();
        ClassJSON cls_json = new ClassJSON();
        ClassConnectionDb cls_conn = new ClassConnectionDb();
          
        public DataTable dataEmployeeRole(string user_name)
        {
            user_name = user_name ?? "";

            //กรณีที่เป็น Employee ทั่วไปเข้าใช้งานระบบ  
            DataTable dt = new DataTable();
            var parameters = new List<SqlParameter>();

            sqlstr = "usp_GetQueryEmployeeRole";
            parameters = new List<SqlParameter>();
            if (!string.IsNullOrEmpty(user_name))
            {
                parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name });
            }

            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters, "data", true);

            return dt;
        }
        public DataTable dataUserRole(string user_name)
        {
            user_name = user_name ?? "";

            //กรณีที่เป็น Employee ที่กำหนดสิทธิ์ในระบบ
            DataTable dt = new DataTable();
            var parameters = new List<SqlParameter>();

            sqlstr = "usp_GetQueryUserRole";
            parameters = new List<SqlParameter>();
            if (!string.IsNullOrEmpty(user_name))
            {
                parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name });
            }

            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters, "data", true);

            return dt;
        }
        private DataTable _dataUser_Role(LoginUserModel param)
        {
            string user_name = (param.user_name ?? "").Trim();
            try
            {
                if (user_name.IndexOf("@") > -1)
                {
                    string[] x = user_name.Split('@');
                    if (x.Length > 1)
                    {
                        user_name = x[0];
                    }
                }
            }
            catch { }

            DataTable dt = new DataTable();
            cls = new ClassFunctions();

            dt = dataUserRole(user_name);

            if (dt?.Rows.Count == 0)
            {
                cls_conn = new ClassConnectionDb();
                dt = new DataTable();
                dt = dataEmployeeRole(user_name);
            }
            else if (user_name.ToLower() == "admin")
            {
                dt.Rows[0]["role_type"] = "admin";
                dt.Rows[0]["user_name"] = "admin";
                dt.Rows[0]["user_id"] = "00000000";
                dt.Rows[0]["user_email"] = "admin-epha@thaioilgroup.com";
                dt.Rows[0]["user_display"] = user_name + "(Admin)";
                dt.Rows[0]["user_img"] = "images/user-avatar.png";
                dt.AcceptChanges();
            }

            //
             
            return dt;
        }
        public string login(LoginUserModel param)
        {
            DataTable dt = new DataTable();
            dt = _dataUser_Role(param);

            return cls_json.SetJSONresult(dt);
        }

        public string authorization_page(PageRoleListModel param)
        {
            string page_controller = (param.page_controller ?? "").Trim();
            string user_name = (param.user_name ?? "").Trim();
            try
            {
                if (user_name.IndexOf("@") > -1)
                {
                    string[] x = user_name.Split('@');
                    if (x.Length > 1)
                    {
                        user_name = x[0];
                    }
                }
            }
            catch { }

            DataTable dt = new DataTable();
            dt = _dtAuthorization_Page(user_name, page_controller);

            return cls_json.SetJSONresult(dt);
        }

        public string check_authorization_page_fix(PageRoleListModel param)
        {
            string page_controller = (param.page_controller ?? "").Trim();
            string user_name = (param.user_name ?? "").Trim();

            try
            {
                if (user_name.IndexOf("@") > -1)
                {
                    string[] x = user_name.Split('@');
                    if (x.Length > 1)
                    {
                        user_name = x[0];
                    }
                }
            }
            catch { }

            string role_type = _dtAuthorization_RoleType(user_name);

            DataTable dt = new DataTable();
            dt = _dtAuthorization_Page(user_name, page_controller);

            dt.Columns.Add("followup_page", typeof(int));
            dt.AcceptChanges();

            ClassHazop cls = new ClassHazop();
            DataTable dtFollow = new DataTable();
            dtFollow = cls.DataHomeTask((role_type == "admin" ? "" : user_name), "", "", false, true, "13");
            if (dtFollow?.Rows.Count == 0)
            {
                dtFollow = cls.DataHomeTask((role_type == "admin" ? "" : user_name), "", "", false, true, "14");
            }

            for (int i = 0; i < dt?.Rows.Count; i++)
            {
                dt.Rows[i]["followup_page"] = 0;
                string filterExpression = "pha_type = '" + (dt.Rows[i]["page_controller"]?.ToString() ?? "").ToUpper() + "'";
                if (dtFollow?.Select(filterExpression).Length > 0)
                {
                    dt.Rows[i]["followup_page"] = 1;
                }
            }
            dt.AcceptChanges();

            return cls_json.SetJSONresult(dt);
        }
        public string _dtAuthorization_RoleType(string user_name)
        {
            user_name = (user_name ?? "").ToLower();

            DataTable dt = new DataTable();
            var parameters = new List<SqlParameter>();

            sqlstr = "usp_GetQueryUserRole";
            parameters = new List<SqlParameter>();
            if (!string.IsNullOrEmpty(user_name))
            {
                parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name });
            }

            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters, "data", true);


            if (dt?.Rows.Count > 0)
            {
                for (int i = 0; i < dt?.Rows.Count; i++)
                {
                    string role_type = (dt.Rows[i]["role_type"]?.ToString() ?? "").ToLower();

                    if (role_type == "admin") { return "admin"; }

                    return role_type;
                }
            }
            return "employee";
        }
        public DataTable _dtAuthorization_Page(string user_name, string page_controller)
        {
            user_name = (user_name ?? "").ToLower();
            user_name = (user_name == "admin" ? "" : user_name);
            page_controller = (page_controller ?? "").ToLower();


            DataTable dt = new DataTable();
            var parameters = new List<SqlParameter>();

            sqlstr = "usp_GetQueryPageRole";
            parameters = new List<SqlParameter>();
            if (!string.IsNullOrEmpty(user_name))
            {
                parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name });
            }
            if (!string.IsNullOrEmpty(page_controller))
            {
                parameters.Add(new SqlParameter("@page_controller", SqlDbType.VarChar, 1000) { Value = page_controller });
            }

            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters, "data", true);

            if (dt?.Rows.Count == 0)
            {
                //กรณีที่ไม่มี menu?? ถึงขั้นตอนนี้น่าจะต้องมี --> create new row
                dt.NewRow(); dt.AcceptChanges();
            }

            return dt;
        }

        public DataTable _dtAuthorization_Page_By_Doc(string user_name, string page_controller)
        {
            user_name = (user_name ?? "").ToLower();
            user_name = (user_name == "admin" ? "" : user_name);
            page_controller = (page_controller ?? "").ToLower();


            DataTable dt = new DataTable();
            var parameters = new List<SqlParameter>();

            sqlstr = @" select distinct a.user_name, a.pha_sub_software as page_controller from ( 
                             select ht.pha_sub_software, dt.*
                             from epha_t_header ht 
                             inner join (
                             select distinct a.responder_user_name as user_name, a.id_pha 
                             from epha_t_node_worksheet a inner join vw_epha_max_seq_by_pha_no b on a.id_pha = b.id_pha
                             where a.responder_user_name is not null and isnull(a.action_project_team,0) = 0 
                             union
                             select distinct a.user_name, a.id_pha 
                             from epha_t_member_team a inner join vw_epha_max_seq_by_pha_no b on a.id_pha = b.id_pha
                             where a.user_name is not null
                             union
                             select distinct a.user_name, a.id_pha 
                             from epha_t_approver a inner join vw_epha_max_seq_by_pha_no b on a.id_pha = b.id_pha
                             where a.user_name is not null
                             union
                             select distinct a.user_name, a.id_pha 
                             from (select ta3.user_name, ta3.id_pha from epha_t_approver_ta3 ta3 inner join epha_t_approver ta2 on ta3.id_approver = ta2.id)a inner join vw_epha_max_seq_by_pha_no b on a.id_pha = b.id_pha
                             where a.user_name is not null
                             union
                             select distinct a.request_user_name as user_name, a.id as id_pha 
                             from epha_t_header a inner join vw_epha_max_seq_by_pha_no b on a.id = b.id_pha
                             where a.request_user_name is not null
                             )dt on ht.id = dt.id_pha 
                        )a where a.pha_sub_software is not null";

            if (!string.IsNullOrEmpty(user_name))
            {
                sqlstr += " and lower(a.user_name)  = lower(@user_name)  ";
                parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name ?? "" });
            }
            if (!string.IsNullOrEmpty(page_controller))
            {
                sqlstr += " and lower(a.page_controller)  = lower(@page_controller)  ";
                parameters.Add(new SqlParameter("@page_controller", SqlDbType.VarChar, 1000) { Value = page_controller ?? "" });
            }
            sqlstr += " order by a.pha_sub_software ";

            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);

            if (dt?.Rows.Count == 0)
            {
                // กรณีที่ไม่มี data?? ถึงขั้นตอนนี้น่าจะต้องมี --> create new row
                DataRow newRow = dt.NewRow();
                dt.Rows.Add(newRow);
                dt.AcceptChanges();
            }

            return dt;
        }
        private int get_max_seq(string table_name)
        {
            if (string.IsNullOrEmpty(table_name))
            {
                throw new ArgumentException("Table name cannot be null or empty", nameof(table_name));
            }

            string sqlstr = string.Format("SELECT COALESCE(MAX(seq), 0) + 1 AS seq FROM {0}", table_name);

            DataTable dt = new DataTable();
            var parameters = new List<SqlParameter>();

            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);

            if (dt?.Rows.Count > 0)
            {
                return Convert.ToInt32(dt.Rows[0]["seq"]);
            }

            return 1; // Default value if no rows returned
        }
        public string register_account(RegisterAccountModel param)
        {
            string user_displayname = (param.user_displayname ?? "").Trim();
            string user_email = (param.user_email ?? "").Trim();
            string user_password = (param.user_password ?? "").Trim();
            string user_password_confirm = (param.user_password_confirm ?? "").Trim();

            string ret = "";
            string msg = "";

            DataTable dt = new DataTable();
            cls = new ClassFunctions();

            string sqlstr = @"SELECT a.user_name, a.user_id, a.user_email, a.user_displayname
                      ,LOWER(COALESCE(c.name, 'employee')) AS role_type
                      FROM VW_EPHA_PERSON_DETAILS a
                      INNER JOIN EPHA_M_ROLE_SETTING b ON LOWER(a.user_name) = LOWER(b.user_name) AND b.active_type = 1
                      INNER JOIN EPHA_M_ROLE_TYPE c ON LOWER(c.id) = LOWER(b.id_role_group) AND c.active_type = 1
                      WHERE a.active_type = 1";

            var parameters = new List<SqlParameter>();
            if (!string.IsNullOrEmpty(user_email))
            {
                sqlstr += " AND LOWER(a.user_email) = LOWER(@user_email)";
                parameters.Add(new SqlParameter("@user_email", SqlDbType.VarChar, 100) { Value = user_email });
            }
            sqlstr += " ORDER BY a.user_name, c.name";

            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters, "data", false);

            if (dt?.Rows.Count > 0)
            {
                ret = "true";
                msg = "User already has data in the system.";
            }
            else
            {
                int seq = get_max_seq("EPHA_REGISTER_ACCOUNT");

                using (var conn = new SqlConnection(ClassConnectionDb.ConnectionString()))
                {
                    conn.Open();

                    using (var transaction = conn.BeginTransaction())
                    {
                        try
                        {
                            sqlstr = @"INSERT INTO EPHA_REGISTER_ACCOUNT
                               (SEQ, ID, REGISTER_TYPE, USER_DISPLAYNAME, USER_EMAIL, USER_PASSWORD, USER_PASSWORD_CONFIRM, ACCEPT_STATUS, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY)
                               VALUES (@SEQ, @ID, @REGISTER_TYPE, @USER_DISPLAYNAME, @USER_EMAIL, @USER_PASSWORD, @USER_PASSWORD_CONFIRM, @ACCEPT_STATUS, GETDATE(), NULL, @CREATE_BY, NULL)";

                            parameters = new List<SqlParameter>
                            {
                                new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq },
                                new SqlParameter("@ID", SqlDbType.Int) { Value = seq },
                                new SqlParameter("@REGISTER_TYPE", SqlDbType.Int) { Value = 1 },
                                new SqlParameter("@USER_DISPLAYNAME", SqlDbType.VarChar, 4000) { Value = user_displayname },
                                new SqlParameter("@USER_EMAIL", SqlDbType.VarChar, 100) { Value = user_email },
                                new SqlParameter("@USER_PASSWORD", SqlDbType.VarChar, 50) { Value = user_password },
                                new SqlParameter("@USER_PASSWORD_CONFIRM", SqlDbType.VarChar, 50) { Value = user_password_confirm },
                                new SqlParameter("@ACCEPT_STATUS", SqlDbType.VarChar, 50) { Value = DBNull.Value },
                                new SqlParameter("@CREATE_BY", SqlDbType.VarChar, 50) { Value = "system" }
                            };

                            ClassConnectionDb cls_conn = new ClassConnectionDb();
                            ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                            if (ret == "true")
                            {
                                transaction.Commit();
                            }
                            else
                            {
                                transaction.Rollback();
                                ret = "error";
                                msg = ret;
                            }
                        }
                        catch (Exception ex)
                        {
                            transaction.Rollback();
                            ret = "error: " + ex.Message;
                        }
                    }
                }

                if (ret.ToLower() == "true")
                {
                    ret = "true";
                    msg = "User registration is complete. Please wait for the login credentials from the system administrator.";
                }
                else
                {
                    ret = "error";
                    msg = ret;
                }
            }

            if (ret.ToLower() == "true")
            {
                // Email แจ้ง admin ให้ accept การ register
                ClassEmail clsmail = new ClassEmail();
                clsmail.MailToAdminRegisterAccount(user_displayname, user_email, user_password, user_password_confirm);
            }

            dt = new DataTable();
            dt.Columns.Add("status");
            dt.Columns.Add("msg");
            dt.AcceptChanges();

            dt.Rows.Add(dt.NewRow());
            dt.AcceptChanges();
            dt.Rows[0]["status"] = ret;
            dt.Rows[0]["msg"] = msg;

            return cls_json.SetJSONresult(dt);
        }

        public string update_register_account(RegisterAccountModel param)
        {
            string user_active = (param.user_active ?? "").Trim();
            string user_email = (param.user_email ?? "").Trim();
            string accept_status = (param.accept_status ?? "").Trim();

            string ret = "";
            string msg = "";

            DataTable dt = new DataTable();
            cls = new ClassFunctions();

            string sqlstr = @"SELECT a.user_name, a.user_id, a.user_email, a.user_displayname
                      ,LOWER(COALESCE(c.name, 'employee')) AS role_type
                      FROM VW_EPHA_PERSON_DETAILS a
                      INNER JOIN EPHA_M_ROLE_SETTING b ON LOWER(a.user_name) = LOWER(b.user_name) AND b.active_type = 1
                      INNER JOIN EPHA_M_ROLE_TYPE c ON LOWER(c.id) = LOWER(b.id_role_group) AND c.active_type = 1
                      WHERE a.active_type = 1";

            var parameters = new List<SqlParameter>();
            if (!string.IsNullOrEmpty(user_email))
            {
                sqlstr += " AND LOWER(a.user_email) = LOWER(@user_email)";
                parameters.Add(new SqlParameter("@user_email", SqlDbType.VarChar, 100) { Value = user_email });
            }
            sqlstr += " ORDER BY a.user_name, c.name";

            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters, "data", false);

            #region insert/update  
            int seq = get_max_seq("EPHA_REGISTER_ACCOUNT");

            using (var conn = new SqlConnection(ClassConnectionDb.ConnectionString()))
            {
                conn.Open();

                using (var transaction = conn.BeginTransaction())
                {
                    try
                    {
                        sqlstr = @"UPDATE EPHA_REGISTER_ACCOUNT SET
                           ACCEPT_STATUS = @ACCEPT_STATUS,
                           UPDATE_DATE = GETDATE(),
                           UPDATE_BY = @UPDATE_BY
                           WHERE USER_EMAIL = @USER_EMAIL";

                        parameters = new List<SqlParameter>
                {
                    new SqlParameter("@ACCEPT_STATUS", SqlDbType.VarChar, 50) { Value = accept_status },
                    new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 400) { Value = user_active },
                    new SqlParameter("@USER_EMAIL", SqlDbType.VarChar, 100) { Value = user_email }
                };

                        ClassConnectionDb cls_conn = new ClassConnectionDb();
                        ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                        if (ret == "true")
                        {
                            transaction.Commit();
                        }
                        else
                        {
                            transaction.Rollback();
                            ret = "error";
                            msg = ret;
                        }
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        ret = "error: " + ex.Message;
                    }
                }
            }
            #endregion insert/update    

            if (ret.ToLower() == "true")
            {
                ret = "true";
                msg = "User registration update is complete.";
            }
            else
            {
                ret = "error";
                msg = ret;
            }

            if (ret.ToLower() == "true")
            {
                if (dt?.Rows.Count > 0)
                {
                    string user_displayname = (dt.Rows[0]["user_displayname"] + "").Trim();
                    string user_password = (dt.Rows[0]["user_password"] + "").Trim();
                    string user_password_confirm = (dt.Rows[0]["user_password_confirm"] + "").Trim();

                    // email แจ้งผู้ใช้งานว่าการลงทะเบียนสำเร็จ
                    ClassEmail clsmail = new ClassEmail();
                    clsmail.MailToUserRegisterAccount(user_displayname, user_email, user_password, user_password_confirm, accept_status);
                }
            }

            dt = new DataTable();
            dt.Columns.Add("status");
            dt.Columns.Add("msg");
            dt.AcceptChanges();

            dt.Rows.Add(dt.NewRow());
            dt.AcceptChanges();
            dt.Rows[0]["status"] = ret;
            dt.Rows[0]["msg"] = msg;

            return cls_json.SetJSONresult(dt);
        } 
    }
}
