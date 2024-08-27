
using dotnet_epha_api.Class;
using dotnet6_epha_api.Class;
using Microsoft.Exchange.WebServices.Data;
using Model;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.Data;
using System.Data.SqlClient;
using System.Transactions;

namespace Class
{
    public class ClassMasterData
    {
        private string sqlstr = "";
        private string jsper = "";
        private ClassFunctions cls = new ClassFunctions();
        private ClassJSON cls_json = new ClassJSON();
        private ClassConnectionDb cls_conn = new ClassConnectionDb();
        private ClassHazop clshazop = new ClassHazop();
        private DataSet dsData = new DataSet();
        private DataTable dt = new DataTable();
        private DataTable dtma = new DataTable();

        #region function
        string[] sMonth = ("JAN,FEB,MAR,APR,MAY,JUN,JUL,AUG,SEP,OCT,NOV,DEC").Split(',');

        private static DataTable refMsg(string status, string remark)
        {
            DataTable dtMsg = new DataTable();
            dtMsg.Columns.Add("status");
            dtMsg.Columns.Add("remark");
            dtMsg.Columns.Add("seq_ref");
            dtMsg.AcceptChanges();

            dtMsg.Rows.Add(dtMsg.NewRow());
            dtMsg.Rows[0]["status"] = status;
            dtMsg.Rows[0]["remark"] = remark;
            return dtMsg;
        }
        private static DataTable refMsg(string status, string remark, string seq_new)
        {
            DataTable dtMsg = new DataTable();
            dtMsg.Columns.Add("status");
            dtMsg.Columns.Add("remark");
            dtMsg.Columns.Add("seq_new");
            dtMsg.AcceptChanges();

            dtMsg.Rows.Add(dtMsg.NewRow());
            dtMsg.Rows[0]["status"] = status;
            dtMsg.Rows[0]["remark"] = remark;
            dtMsg.Rows[0]["seq_new"] = seq_new;
            return dtMsg;
        }
        private DataTable ConvertDStoDT(DataSet _ds, string _table_name)
        {
            DataTable _dt = new DataTable();
            try
            {
                _dt = dsData.Tables[_table_name].Copy();
            }
            catch { }
            return _dt;
        }
        private int get_max(string table_name, string Neverkey = "")
        {
            if (string.IsNullOrEmpty(Neverkey)) { Neverkey = "id"; }

            DataTable _dt = new DataTable();
            cls = new ClassFunctions();
            try
            {

                sqlstr = $@" select coalesce(max({Neverkey}),0)+1 as id from {table_name}";

                _dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

                return Convert.ToInt32(_dt.Rows[0]["id"].ToString() + "");
            }
            catch
            {
                sqlstr = $@" select coalesce(max(seq),0)+1 as id from {table_name}";

                _dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

                return Convert.ToInt32(_dt.Rows[0]["id"].ToString() + "");
            }
        }
        private void ConvertJSONListresultToDataSet(ref string msg, ref string ret, ref DataSet dsData, DataMasterListModel param)
        {
            string table_name = param.json_name ?? "";
            jsper = param.json_data + "";
            if (jsper.Trim() == "") { msg = "No Data."; ret = "Error"; return; }
            try
            {
                dt = new DataTable();
                dt = cls_json.ConvertJSONresult(jsper);

                dt.TableName = table_name;
                dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }

        }
        private void ConvertJSONresultToDataSet(ref string msg, ref string ret, ref DataSet dsData, SetDataMasterModel param, string table_name = "data")
        {
            jsper = param.json_data + "";
            if (jsper.Trim() == "") { msg = "No Data."; ret = "Error"; return; }
            try
            {
                dt = new DataTable();
                dt = cls_json.ConvertJSONresult(jsper);

                dt.TableName = table_name;
                dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }

        }
        private void ConvertJSONresultToDataSetManageuser(ref string msg, ref string ret, ref DataSet dsData, SetManageuser param)
        {
            jsper = param.json_register_account + "";
            if (jsper.Trim() == "") { msg = "No Data."; ret = "Error"; return; }
            try
            {
                dt = new DataTable();
                dt = cls_json.ConvertJSONresult(jsper);

                dt.TableName = "register_account";
                dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }

        }
        private void ConvertJSONresultToDataSetAuthorizationSetting(ref string msg, ref string ret, ref DataSet dsData, SetAuthorizationSetting param)
        {
            jsper = param.json_role_type + "";
            if (jsper.Trim() == "") { msg = "No Data."; ret = "Error"; return; }
            try
            {
                dt = new DataTable();
                dt = cls_json.ConvertJSONresult(jsper);

                dt.TableName = "role_type";
                dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }

            jsper = param.json_menu_setting + "";
            //if (jsper.Trim() == "") { msg = "No Data."; ret = "Error"; return; }
            try
            {
                dt = new DataTable();
                dt = cls_json.ConvertJSONresult(jsper);

                dt.TableName = "menu_setting";
                dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }

            jsper = param.json_role_setting + "";
            //if (jsper.Trim() == "") { msg = "No Data."; ret = "Error"; return; }
            try
            {
                dt = new DataTable();
                dt = cls_json.ConvertJSONresult(jsper);

                dt.TableName = "role_setting";
                dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }

        }
        public DataTable refMsgSaveMaster(string status, string remark, string seq_new)
        {
            DataTable dtMsg = new DataTable();
            dtMsg.Columns.Add("status");
            dtMsg.Columns.Add("remark");
            dtMsg.Columns.Add("seq_new");
            dtMsg.AcceptChanges();

            dtMsg.Rows.Add(dtMsg.NewRow());
            dtMsg.Rows[0]["status"] = status;
            dtMsg.Rows[0]["remark"] = remark;
            dtMsg.Rows[0]["seq_new"] = seq_new;
            return dtMsg;
        }
        private string MapPathFiles(string _folder)
        {
            return (Path.Combine(Directory.GetCurrentDirectory(), "") + _folder.Replace("~", ""));
        }
        #endregion function

        #region employee pis 
        public string importfile_data_employe(uploadFile uploadFile, string folder)
        {
            string msg_error = "";
            DataSet _dsData = new DataSet();
            DataTable dtdef = (DataTable)ClassFile.DatatableFile();
            string _file_name = "";
            string _file_download_name = "";
            string _file_fullpath_name = "";

            if (dtdef != null && dt != null && uploadFile != null)
            {
                IFormFileCollection files = uploadFile?.file_obj;
                if (files?.Count > 0)
                {
                    var file_seq = uploadFile?.file_seq ?? "";
                    var file_name = uploadFile?.file_name ?? "";
                    var file_part = uploadFile?.file_part ?? "";
                    var file_doc = uploadFile?.file_doc ?? "";
                    var file_sub_software = uploadFile?.sub_software ?? "";
                    folder = "PersonalData";

                    if (string.IsNullOrEmpty(folder)) { msg_error = "Invalid folder."; }
                    else
                    {
                        msg_error = ClassFile.copy_file_data_to_server(ref _file_name, ref _file_download_name, ref _file_fullpath_name
                        , files, folder, "employee_list", file_doc, true, false);

                        if (string.IsNullOrEmpty(file_seq) || string.IsNullOrEmpty(_file_name) || string.IsNullOrEmpty(_file_download_name) || string.IsNullOrEmpty(_file_fullpath_name))
                        { msg_error = "Invalid file."; }
                        else
                        {
                            #region import file data to database 
                            // ตรวจสอบว่าไฟล์ที่ได้มามีอยู่จริงหรือไม่
                            if (!File.Exists(_file_fullpath_name))
                            {
                                msg_error = $"The specified file does not exist.{_file_fullpath_name}";
                            }
                            else
                            {
                                if (string.IsNullOrEmpty(ClassFile.check_file_on_server(folder, _file_fullpath_name)))
                                {
                                    DataTable dtData = new DataTable();
                                    dtData = excel_person_details(_file_fullpath_name);
                                    if (dsData != null) _dsData.Tables.Add(dtData.Copy());
                                    msg_error = "";
                                }
                                else { msg_error = "The file is not within the allowed directory."; }
                            }
                            #endregion import file data to database 
                        }
                    }
                }
            }
            else { msg_error = "No Data."; }

            if (dtdef != null)
            {
                ClassFile.AddRowToDataTable(ref dtdef, _file_name, _file_download_name, msg_error);
                if (dsData != null) _dsData.Tables.Add(dtdef.Copy());
            }

            return JsonConvert.SerializeObject(_dsData, Formatting.Indented);
        }
        public DataTable excel_person_details(string _excel_file_name)
        {
            string msg = "";
            DataTable dtWorksheet = new DataTable();

            sqlstr = "select * from EPHA_PERSON_DETAILS where 1=2 ";
            cls_conn = new ClassConnectionDb();
            dtWorksheet = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            // ตรวจสอบว่าไฟล์ที่ได้มามีอยู่จริงหรือไม่
            if (!File.Exists(_excel_file_name))
            {
                throw new ArgumentException("The specified file does not exist: {_excel_file_name}");
            }
            if (!string.IsNullOrEmpty(ClassFile.check_file_on_server("PersonalData", _excel_file_name)))
            {
                throw new ArgumentException("The file is not within the allowed directory.");
            }

            FileInfo template = new FileInfo(_excel_file_name);

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage(template))
            {
                ExcelWorksheet sourceWorksheet = excelPackage.Workbook.Worksheets[0];
                ExcelWorksheet worksheet = sourceWorksheet;

                //Worksheet
                #region Worksheet
                if (true)
                {
                    int irows = 0;
                    int startRows = 1;
                    int endRowsWorksheet = 1;
                    int icol_start = 1;
                    int icol_end = 27;
                    ;
                    int iWorksheetRows = worksheet.Dimension.Rows;
                    for (int i = 2; i <= iWorksheetRows; i++) // ใช้ Dimension.Rows เพื่อหาขนาดของชีทแน่นอน
                    {
                        startRows = i;
                        icol_start = 1;//เริ่มจาก row 1 = A 

                        //Check PK
                        if (worksheet.Cells[startRows, icol_start].Value?.ToString() == "") { break; }

                        //New Row 
                        dtWorksheet.Rows.Add(dtWorksheet.NewRow());
                        try
                        {
                            for (int j = 0; j < dtWorksheet.Columns.Count; j++)
                            {
                                if (j == 0) { icol_start = j; }

                                string col_name = dtWorksheet.Columns[j].ColumnName;
                                icol_start += 1; dtWorksheet.Rows[irows][col_name] = worksheet.Cells[startRows, icol_start].Value?.ToString();

                            }
                        }
                        catch { }
                        irows += 1;

                        //if (i == 5) { break; }
                    }
                }
                #endregion Worksheet

            }

            if (dtWorksheet?.Rows.Count > 0 && dtWorksheet.AsEnumerable().Any(row => !string.IsNullOrEmpty(row.Field<string>("user_type"))))
            {
                return dtWorksheet;
            }

            //กรณีที่ไม่มีข้อมูลให้ส่ง datatable เปล่ากลับไป
            dtWorksheet.Rows.Clear();
            dt.AcceptChanges();

            return dtWorksheet;
        }

        private string person_details_insert_data(DataTable _dt, string user_type)
        {
            string ret = "";
            int batchSize = 500; // กำหนดขนาดของแต่ละชุดย่อย
            int totalRows = _dt.Rows.Count;

            if (totalRows > 0)
            {
                string sColFix = @"USER_TYPE,EMPLOYEEID,USERID,COMPANYCODE,DEPARTMENT,DIVISION,SECTIONS,UNITS,ORGID,POSID,OBJENFULLNAME,OBJTHFULLNAME,PERSAREA,PERSUBAREA" +
                                 ",THTITLE,THFIRSTNAME,THLASTNAME,ENFIRSTNAME,ENLASTNAME,EMAIL,CONTRACT,HOLDERPOSITION,EMPSUBGROUP,MANAGERIAL" +
                                 ",REPORTTOPOS,REPORTTOID,REPORTTONAME,REPORTTOEMAIL,ACTIVE_TYPE";
                string[] sColFixSplit = sColFix.Split(",");

                string connectionString = ClassConnectionDb.ConnectionString();

                try
                {
                    using (var scope = new TransactionScope())
                    {
                        using (var conn = new SqlConnection(connectionString))
                        {
                            conn.Open();

                            // Delete existing records for the user type
                            string deleteSql = "DELETE FROM epha_person_details WHERE user_type = @user_type";
                            using (var deleteCmd = new SqlCommand(deleteSql, conn))
                            {
                                deleteCmd.Parameters.Add(new SqlParameter("@user_type", SqlDbType.NVarChar, 255) { Value = user_type });
                                deleteCmd.ExecuteNonQuery();
                            }

                            for (int startRow = 0; startRow < totalRows; startRow += batchSize)
                            {
                                int endRow = Math.Min(startRow + batchSize, totalRows);
                                try
                                {
                                    for (int i = startRow; i < endRow; i++)
                                    {
                                        string action_type = _dt.Rows[i]["action_type"]?.ToString() ?? "";
                                        string sqlstr = "";
                                        List<SqlParameter> parameters = new List<SqlParameter>();

                                        if (action_type == "insert")
                                        {
                                            sqlstr = "INSERT INTO epha_person_details(" + sColFix +
                                                     ", CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) VALUES (" +
                                                     "@USER_TYPE, @EMPLOYEEID, @USERID, @COMPANYCODE, @DEPARTMENT, @DIVISION, @SECTIONS, @UNITS, @ORGID, @POSID, @OBJENFULLNAME, @OBJTHFULLNAME, @PERSAREA, @PERSUBAREA, " +
                                                     "@THTITLE, @THFIRSTNAME, @THLASTNAME, @ENFIRSTNAME, @ENLASTNAME, @EMAIL, @CONTRACT, @HOLDERPOSITION, @EMPSUBGROUP, @MANAGERIAL, " +
                                                     "@REPORTTOPOS, @REPORTTOID, @REPORTTONAME, @REPORTTOEMAIL, @ACTIVE_TYPE, GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                                            foreach (string col in sColFixSplit)
                                            {
                                                parameters.Add(new SqlParameter("@" + col, SqlDbType.NVarChar, 255) { Value = _dt.Rows[i][col] ?? DBNull.Value });
                                            }
                                            parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = _dt.Rows[i]["CREATE_BY"] ?? DBNull.Value });
                                            parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = _dt.Rows[i]["UPDATE_BY"] ?? DBNull.Value });
                                        }
                                        else if (action_type == "update")
                                        {
                                            sqlstr = "UPDATE epha_person_details SET ";

                                            for (int k = 0; k < sColFixSplit.Length; k++)
                                            {
                                                if (k > 0) sqlstr += ", ";
                                                sqlstr += sColFixSplit[k] + " = @" + sColFixSplit[k];
                                                parameters.Add(new SqlParameter("@" + sColFixSplit[k], SqlDbType.NVarChar, 255) { Value = _dt.Rows[i][sColFixSplit[k]] ?? DBNull.Value });
                                            }
                                            sqlstr += " WHERE EMPLOYEEID = @EMPLOYEEID AND USERID = @USERID AND USER_TYPE = @USER_TYPE";

                                            parameters.Add(new SqlParameter("@EMPLOYEEID", SqlDbType.NVarChar, 255) { Value = _dt.Rows[i]["EMPLOYEEID"] ?? DBNull.Value });
                                            parameters.Add(new SqlParameter("@USERID", SqlDbType.NVarChar, 255) { Value = _dt.Rows[i]["USERID"] ?? DBNull.Value });
                                            parameters.Add(new SqlParameter("@USER_TYPE", SqlDbType.NVarChar, 255) { Value = user_type });
                                        }
                                        else if (action_type == "delete")
                                        {
                                            sqlstr = "DELETE FROM epha_person_details WHERE EMPLOYEEID = @EMPLOYEEID AND USERID = @USERID AND USER_TYPE = @USER_TYPE";

                                            parameters.Add(new SqlParameter("@EMPLOYEEID", SqlDbType.NVarChar, 255) { Value = _dt.Rows[i]["EMPLOYEEID"] ?? DBNull.Value });
                                            parameters.Add(new SqlParameter("@USERID", SqlDbType.NVarChar, 255) { Value = _dt.Rows[i]["USERID"] ?? DBNull.Value });
                                            parameters.Add(new SqlParameter("@USER_TYPE", SqlDbType.NVarChar, 255) { Value = user_type });
                                        }

                                        if (!string.IsNullOrEmpty(sqlstr))
                                        {
                                            using (var cmd = new SqlCommand(sqlstr, conn))
                                            {
                                                cmd.Parameters.AddRange(parameters.ToArray());
                                                cmd.ExecuteNonQuery();
                                            }
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    ret = "error: " + ex.Message;
                                    throw; // Rethrow the exception to be caught by the outer catch block
                                }
                            }
                        }
                        scope.Complete();
                        ret = "true";
                    }
                }
                catch (Exception ex)
                {
                    ret = "error: " + ex.Message;
                }
            }

            return ret;
        }

        #endregion employee pis 

        #region Manage User
        public string get_manageuser(LoadMasterPageModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = (param.user_name + "").Trim();

            clshazop = new ClassHazop();
            clshazop.get_employee_list(true, ref dsData);

            //epha_register_account -> register_account
            int iMaxSeqRegister = get_max("epha_register_account");

            sqlstr = @" select a.*, 'update' as action_type, 0 as action_change
                        from epha_register_account a 
                        order by seq ";

            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            if (dt?.Rows.Count == 0)
            {
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = iMaxSeqRegister;
                dt.Rows[0]["id"] = iMaxSeqRegister;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();

                iMaxSeqRegister += 1;
            }
            dt.AcceptChanges();

            dt.TableName = "register_account";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();



            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "register_account", (iMaxSeqRegister + 1).ToString());
            dtma.TableName = "max";

            dsData.Tables.Add(dtma.Copy()); dsData.AcceptChanges();

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;

        }
        public string set_manageuser(SetManageuser param)
        {
            dtma = new DataTable();
            dsData = new DataSet();

            int iMaxSeq = 0;
            string msg = "";
            string ret = "";
            string user_name = (param.user_name + "").Trim();
            string role_type = (param.role_type + "").Trim();
            string json_register_account = (param.json_register_account + "").Trim();

            SetManageuser param_def = new SetManageuser();
            param_def.user_name = (param.user_name + "").Trim();
            param_def.role_type = (param.role_type + "").Trim();
            param_def.json_register_account = (param.json_register_account + "").Trim();

            ConvertJSONresultToDataSetManageuser(ref msg, ref ret, ref dsData, param_def);

            if (ret.ToLower() == "error") { }
            else
            {
                //set role_type,role_setting,menu_setting 
                var iMaxSeqRoleRole = get_max("epha_register_account");
                ret = set_register_account(ref dsData, ref iMaxSeqRoleRole);
            }
            return cls_json.SetJSONresult(refMsgSaveMaster(ret, msg, (iMaxSeq.ToString() + "").ToString()));
        }
        public string set_register_account(ref DataSet dsData, ref int seq_now)
        {
            string ret = "";

            #region update data 
            DataTable dt = ConvertDStoDT(dsData, "register_account");

            string connectionString = ClassConnectionDb.ConnectionString();

            using (var conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (var transaction = conn.BeginTransaction())
                {
                    try
                    {
                        foreach (DataRow row in dt.Rows)
                        {
                            string action_type = row["action_type"]?.ToString() ?? "";
                            string sqlstr = "";
                            List<SqlParameter> parameters = new List<SqlParameter>();

                            if (action_type == "insert")
                            {
                                sqlstr = @"
                            INSERT INTO EPHA_REGISTER_ACCOUNT (
                                SEQ, ID, REGISTER_TYPE, ACCEPT_STATUS, USER_NAME, USER_DISPLAYNAME, USER_EMAIL, USER_PASSWORD, USER_PASSWORD_CONFIRM,
                                CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY
                            ) VALUES (
                                @SEQ, @ID, @REGISTER_TYPE, @ACCEPT_STATUS, @USER_NAME, @USER_DISPLAYNAME, @USER_EMAIL, @USER_PASSWORD, @USER_PASSWORD_CONFIRM,
                                GETDATE(), NULL, @CREATE_BY, @UPDATE_BY
                            )";

                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                                parameters.Add(new SqlParameter("@REGISTER_TYPE", SqlDbType.Int) { Value = row["REGISTER_TYPE"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@ACCEPT_STATUS", SqlDbType.Int) { Value = row["ACCEPT_STATUS"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@USER_NAME", SqlDbType.NVarChar, 50) { Value = row["USER_NAME"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@USER_DISPLAYNAME", SqlDbType.NVarChar, 4000) { Value = row["USER_DISPLAYNAME"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@USER_EMAIL", SqlDbType.NVarChar, 50) { Value = row["USER_EMAIL"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@USER_PASSWORD", SqlDbType.NVarChar, 50) { Value = row["USER_PASSWORD"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@USER_PASSWORD_CONFIRM", SqlDbType.NVarChar, 50) { Value = row["USER_PASSWORD_CONFIRM"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = row["CREATE_BY"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] ?? DBNull.Value });

                                seq_now += 1;
                            }
                            else if (action_type == "update")
                            {
                                seq_now = Convert.ToInt32(row["SEQ"]);

                                sqlstr = @"
                            UPDATE EPHA_REGISTER_ACCOUNT SET
                                ACCEPT_STATUS = @ACCEPT_STATUS, USER_NAME = @USER_NAME, USER_DISPLAYNAME = @USER_DISPLAYNAME,
                                USER_EMAIL = @USER_EMAIL, USER_PASSWORD = @USER_PASSWORD, USER_PASSWORD_CONFIRM = @USER_PASSWORD_CONFIRM
                            WHERE SEQ = @SEQ AND ID = @ID";

                                parameters.Add(new SqlParameter("@ACCEPT_STATUS", SqlDbType.Int) { Value = row["ACCEPT_STATUS"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@USER_NAME", SqlDbType.NVarChar, 50) { Value = row["USER_NAME"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@USER_DISPLAYNAME", SqlDbType.NVarChar, 4000) { Value = row["USER_DISPLAYNAME"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@USER_EMAIL", SqlDbType.NVarChar, 50) { Value = row["USER_EMAIL"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@USER_PASSWORD", SqlDbType.NVarChar, 50) { Value = row["USER_PASSWORD"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@USER_PASSWORD_CONFIRM", SqlDbType.NVarChar, 50) { Value = row["USER_PASSWORD_CONFIRM"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] ?? DBNull.Value });
                            }
                            else if (action_type == "delete")
                            {
                                sqlstr = @"
                            DELETE FROM EPHA_REGISTER_ACCOUNT
                            WHERE SEQ = @SEQ AND ID = @ID";

                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] ?? DBNull.Value });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] ?? DBNull.Value });
                            }

                            if (!string.IsNullOrEmpty(sqlstr))
                            {
                                ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                                if (ret != "true") throw new Exception(ret);
                            }
                        }

                        if (ret == "true")
                        {
                            transaction.Commit();
                        }
                        else
                        {
                            transaction.Rollback();
                        }
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        ret = "error: " + ex.Message;
                        return ret;
                    }
                }
            }

            #endregion update data 

            return ret;
        }

        #endregion Manage User

        #region AuthorizationSetting
        public string get_authorizationsetting(LoadMasterPageModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = (param.user_name + "").Trim();

            int iMaxSeqMenu = 0;
            int iMaxSeqRoleType = 0;
            int iMaxSeqMenuSetting = 0;
            int iMaxSeqRoleSetting = 0;

            #region menu -> fix data
            sqlstr = @"   select a.page_type, a.page_controller
                         , case when a.page_type in ('main') then 0 else 1 end disable_page
                         , 0 as choos_menu
                         , a.seq, a.name
                         from epha_m_menu a 
                         where a.active_type = 1 
                         order by a.page_type, a.seq";

            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            dt.AcceptChanges();
            dt.TableName = "menu";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion menu -> fix data

            #region role_type
            iMaxSeqRoleType = get_max("epha_m_role_type");

            sqlstr = @"  select a.*, 'update' as action_type, 0 as action_change 
                         from epha_m_role_type a 
                         order by seq";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            if (dt?.Rows.Count == 0)
            {
                //กรณีที่เป็นข้อมูลใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = iMaxSeqRoleType;
                dt.Rows[0]["id"] = iMaxSeqRoleType;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();

                iMaxSeqRoleType += 1;
            }
            dt.AcceptChanges();
            dt.TableName = "role_type";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion role_type

            #region menu_setting
            iMaxSeqMenuSetting = get_max("epha_m_menu_setting");

            sqlstr = @"  select a.*, 'update' as action_type, 0 as action_change 
                         , 1 as choos_data 
                         from epha_m_menu_setting a 
                         order by id_role_group, seq ";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            if (dt?.Rows.Count == 0)
            {
                //กรณีที่เป็นข้อมูลใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = iMaxSeqMenuSetting;
                dt.Rows[0]["id"] = iMaxSeqMenuSetting;
                dt.Rows[0]["id_role_group"] = iMaxSeqRoleType;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();

                iMaxSeqMenuSetting += 1;
            }
            dt.AcceptChanges();
            dt.TableName = "menu_setting";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion menu_setting

            #region role_setting
            iMaxSeqRoleSetting = get_max("epha_m_role_setting");

            sqlstr = @"  select a.*, 'update' as action_type, 0 as action_change
                         , emp.user_displayname
                         , 'assets/img/team/avatar.webp' as user_img
                         from epha_m_role_setting a 
                         left join vw_epha_person_details emp on lower(emp.user_name)  = lower(a.user_name) 
                         order by id_role_group, seq  ";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            if (dt?.Rows.Count == 0)
            {
                //กรณีที่เป็นข้อมูลใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = iMaxSeqRoleSetting;
                dt.Rows[0]["id"] = iMaxSeqRoleSetting;
                dt.Rows[0]["id_role_group"] = iMaxSeqRoleType;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();

                iMaxSeqRoleSetting += 1;
            }
            dt.AcceptChanges();
            dt.TableName = "role_setting";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion role_setting

            clshazop = new ClassHazop();
            clshazop.get_employee_list(true, ref dsData);

            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "role_type", (iMaxSeqRoleType + 1).ToString());
            clshazop.set_max_id(ref dtma, "menu_setting", (iMaxSeqMenuSetting + 1).ToString());
            clshazop.set_max_id(ref dtma, "role_setting", (iMaxSeqRoleSetting + 1).ToString());
            dtma.TableName = "max";

            dsData.Tables.Add(dtma.Copy()); dsData.AcceptChanges();

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;

        }
        public string set_authorizationsetting(SetAuthorizationSetting param)
        {
            dtma = new DataTable();
            dsData = new DataSet();

            int iMaxSeq = 0;
            string msg = "";
            string ret = "";
            string user_name = (param.user_name + "").Trim();
            string role_type = (param.role_type + "").Trim();
            string json_role_type = (param.json_role_type + "").Trim();
            string json_menu_setting = (param.json_menu_setting + "").Trim();
            string json_role_setting = (param.json_role_setting + "").Trim();

            SetAuthorizationSetting param_def = new SetAuthorizationSetting();
            param_def.user_name = (param.user_name + "").Trim();
            param_def.role_type = (param.role_type + "").Trim();
            param_def.json_role_type = (param.json_role_type + "").Trim();
            param_def.json_menu_setting = (param.json_menu_setting + "").Trim();
            param_def.json_role_setting = (param.json_role_setting + "").Trim();

            ConvertJSONresultToDataSetAuthorizationSetting(ref msg, ref ret, ref dsData, param_def);

            if (ret.ToLower() == "error") { goto Next_Line; }

            //set role_type,role_setting,menu_setting 
            var iMaxSeqRoleRole = get_max("epha_m_role_type");
            var iMaxSeqRoleSetting = get_max("epha_m_role_setting");
            var iMaxSeqMenuSetting = get_max("epha_m_menu_setting");

            ret = set_role_type(ref dsData, ref iMaxSeqRoleRole);
            if (ret == "true") { goto Next_Line; }

            ret = set_role_setting(ref dsData, ref iMaxSeqRoleSetting);
            if (ret == "true") { goto Next_Line; }

            ret = set_menu_setting(ref dsData, ref iMaxSeqMenuSetting);
            if (ret == "true") { goto Next_Line; }


        Next_Line:;

            return cls_json.SetJSONresult(refMsgSaveMaster(ret, msg, (iMaxSeq.ToString() + "").ToString()));
        }
        public string set_role_type(ref DataSet dsData, ref int seq_now)
        {
            string ret = "";
            string connectionString = ClassConnectionDb.ConnectionString();

            try
            {
                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    using (var transaction = conn.BeginTransaction())
                    {
                        try
                        {
                            DataTable dt = ConvertDStoDT(dsData, "role_type");

                            for (int i = 0; i < dt?.Rows.Count; i++)
                            {
                                string action_type = dt.Rows[i]["action_type"]?.ToString() ?? "";
                                string sqlstr = "";
                                List<SqlParameter> parameters = new List<SqlParameter>();

                                if (action_type == "insert")
                                {
                                    sqlstr = "INSERT INTO EPHA_M_ROLE_TYPE " +
                                             "(SEQ, ID, NAME, DESCRIPTIONS, DEFAULT_TYPE, ACTIVE_TYPE, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                             "VALUES (@SEQ, @ID, @NAME, @DESCRIPTIONS, @DEFAULT_TYPE, @ACTIVE_TYPE, GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                                    parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 255) { Value = dt.Rows[i]["NAME"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTIONS"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@DEFAULT_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["DEFAULT_TYPE"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["CREATE_BY"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"] ?? DBNull.Value });

                                    seq_now += 1;
                                }
                                else if (action_type == "update")
                                {
                                    sqlstr = "UPDATE EPHA_M_ROLE_TYPE SET " +
                                             "NAME = @NAME, DESCRIPTIONS = @DESCRIPTIONS, ACTIVE_TYPE = @ACTIVE_TYPE, DEF_SELECTED = @DEF_SELECTED, " +
                                             "UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                             "WHERE SEQ = @SEQ AND ID = @ID";

                                    parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 255) { Value = dt.Rows[i]["NAME"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTIONS"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@DEF_SELECTED", SqlDbType.Int) { Value = dt.Rows[i]["DEF_SELECTED"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"] ?? DBNull.Value });
                                }
                                else if (action_type == "delete")
                                {
                                    sqlstr = "DELETE FROM EPHA_M_ROLE_TYPE WHERE SEQ = @SEQ AND ID = @ID";

                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"] ?? DBNull.Value });
                                }

                                if (!string.IsNullOrEmpty(sqlstr))
                                {
                                    using (var cmd = new SqlCommand(sqlstr, conn, transaction))
                                    {
                                        cmd.Parameters.AddRange(parameters.ToArray());
                                        ret = cmd.ExecuteNonQuery() > 0 ? "true" : "false";
                                        if (ret != "true") throw new Exception("Database operation failed");
                                    }
                                }
                            }

                            transaction.Commit();
                            ret = "true";
                        }
                        catch (Exception ex)
                        {
                            transaction.Rollback();
                            ret = "error: " + ex.Message;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return ret;
        }

        public string set_role_setting(ref DataSet dsData, ref int seq_now)
        {
            string ret = "";
            string connectionString = ClassConnectionDb.ConnectionString();

            try
            {
                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    using (var transaction = conn.BeginTransaction())
                    {
                        try
                        {
                            DataTable dt = ConvertDStoDT(dsData, "role_setting");

                            for (int i = 0; i < dt?.Rows.Count; i++)
                            {
                                string action_type = dt.Rows[i]["action_type"]?.ToString() ?? "";
                                string sqlstr = "";
                                List<SqlParameter> parameters = new List<SqlParameter>();

                                if (action_type == "insert")
                                {
                                    sqlstr = "INSERT INTO EPHA_M_ROLE_SETTING " +
                                             "(SEQ, ID_ROLE_GROUP, USER_NAME, ACTIVE_TYPE, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                             "VALUES (@SEQ, @ID_ROLE_GROUP, @USER_NAME, @ACTIVE_TYPE, GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                                    parameters.Add(new SqlParameter("@ID_ROLE_GROUP", SqlDbType.Int) { Value = dt.Rows[i]["ID_ROLE_GROUP"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@USER_NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["USER_NAME"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["CREATE_BY"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"] ?? DBNull.Value });

                                    seq_now += 1;
                                }
                                else if (action_type == "update")
                                {
                                    sqlstr = "UPDATE EPHA_M_ROLE_SETTING SET " +
                                             "ID_ROLE_GROUP = @ID_ROLE_GROUP, USER_NAME = @USER_NAME, UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                             "WHERE SEQ = @SEQ";

                                    parameters.Add(new SqlParameter("@ID_ROLE_GROUP", SqlDbType.Int) { Value = dt.Rows[i]["ID_ROLE_GROUP"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@USER_NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["USER_NAME"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"] ?? DBNull.Value });
                                }
                                else if (action_type == "delete")
                                {
                                    sqlstr = "DELETE FROM EPHA_M_ROLE_SETTING WHERE SEQ = @SEQ";

                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"] ?? DBNull.Value });
                                }

                                if (!string.IsNullOrEmpty(sqlstr))
                                {
                                    using (var cmd = new SqlCommand(sqlstr, conn, transaction))
                                    {
                                        cmd.Parameters.AddRange(parameters.ToArray());
                                        ret = cmd.ExecuteNonQuery() > 0 ? "true" : "false";
                                        if (ret != "true") throw new Exception("Database operation failed");
                                    }
                                }
                            }

                            transaction.Commit();
                            ret = "true";
                        }
                        catch (Exception ex)
                        {
                            transaction.Rollback();
                            ret = "error: " + ex.Message;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return ret;
        }

        public string set_menu_setting(ref DataSet dsData, ref int seq_now)
        {
            string ret = "";
            string connectionString = ClassConnectionDb.ConnectionString();
            string sqlstr = "";

            try
            {
                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    using (var transaction = conn.BeginTransaction())
                    {
                        try
                        {
                            DataTable dt = ConvertDStoDT(dsData, "menu_setting");

                            for (int i = 0; i < dt?.Rows.Count; i++)
                            {
                                string action_type = dt.Rows[i]["action_type"]?.ToString() ?? "";
                                sqlstr = "";
                                List<SqlParameter> parameters = new List<SqlParameter>();

                                if (action_type == "insert")
                                {
                                    sqlstr = "INSERT INTO EPHA_M_MENU_SETTING " +
                                             "(SEQ, ID_ROLE_GROUP, ID_MENU, DESCRIPTIONS, ACTIVE_TYPE, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                             "VALUES (@SEQ, @ID_ROLE_GROUP, @ID_MENU, @DESCRIPTIONS, @ACTIVE_TYPE, GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                                    parameters.Add(new SqlParameter("@ID_ROLE_GROUP", SqlDbType.Int) { Value = dt.Rows[i]["ID_ROLE_GROUP"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ID_MENU", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["ID_MENU"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTIONS"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["CREATE_BY"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"] ?? DBNull.Value });

                                    seq_now += 1;
                                }
                                else if (action_type == "update")
                                {
                                    sqlstr = "UPDATE EPHA_M_MENU_SETTING SET " +
                                             "ID_ROLE_GROUP = @ID_ROLE_GROUP, ID_MENU = @ID_MENU, DESCRIPTIONS = @DESCRIPTIONS, " +
                                             "UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                             "WHERE SEQ = @SEQ";

                                    parameters.Add(new SqlParameter("@ID_ROLE_GROUP", SqlDbType.Int) { Value = dt.Rows[i]["ID_ROLE_GROUP"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ID_MENU", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["ID_MENU"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTIONS"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"] ?? DBNull.Value });
                                }
                                else if (action_type == "delete")
                                {
                                    sqlstr = "DELETE FROM EPHA_M_MENU_SETTING WHERE SEQ = @SEQ";

                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"] ?? DBNull.Value });
                                }

                                if (!string.IsNullOrEmpty(sqlstr))
                                {
                                    using (var cmd = new SqlCommand(sqlstr, conn, transaction))
                                    {
                                        cmd.Parameters.AddRange(parameters.ToArray());
                                        ret = cmd.ExecuteNonQuery() > 0 ? "true" : "false";
                                        if (ret != "true") throw new Exception("Database operation failed");
                                    }
                                }
                            }

                            transaction.Commit();
                            ret = "true";
                        }
                        catch (Exception ex)
                        {
                            transaction.Rollback();
                            ret = "error: " + ex.Message;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return ret;
        }

        #endregion AuthorizationSetting

        #region Manage User
        public string get_master_contractlist(LoadMasterPageModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = (param.user_name + "").Trim();

            int iMaxSeq = get_max("epha_person_details", "employeeid");

            sqlstr = @" select employeeid as seq, a.*, 'update' as action_type, 0 as action_change 
                        from epha_person_details a 
                        where a.user_type = 'contract'
                        order by a.employeeid  ";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            if (dt?.Rows.Count == 0)
            {
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = iMaxSeq;
                dt.Rows[0]["employeeid"] = iMaxSeq;
                dt.Rows[0]["user_type"] = "contract";

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();

                iMaxSeq += 1;
            }
            dt.AcceptChanges();

            dt.TableName = "data";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "seq", (iMaxSeq + 1).ToString());
            dtma.TableName = "max";

            dsData.Tables.Add(dtma.Copy()); dsData.AcceptChanges();

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;

        }
        public string set_master_contractlist(SetMasterGuideWordsModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();

            int iMaxSeq = 0;
            string msg = "";
            string ret = "";
            string user_name = (param.user_name + "").Trim();
            string json_data = (param.json_data + "").Trim();

            DataMasterListModel param_def = new DataMasterListModel();
            param_def.json_name = "data";
            param_def.json_data = param.json_data;
            ConvertJSONListresultToDataSet(ref msg, ref ret, ref dsData, param_def);
            if (ret.ToLower() == "error") { goto Next_Line_Convert; }
            else
            {
                iMaxSeq = get_max("epha_person_details");
                ret = set_contractlist(ref dsData, ref iMaxSeq);
            }


        Next_Line_Convert:;

            return cls_json.SetJSONresult(refMsgSaveMaster(ret, msg, (iMaxSeq.ToString() + "").ToString()));

        }
        public string set_contractlist(ref DataSet dsData, ref int seq_now)
        {
            string ret = "";
            string connectionString = ClassConnectionDb.ConnectionString();
            string sqlstr = "";

            try
            {
                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    using (var transaction = conn.BeginTransaction())
                    {
                        try
                        {
                            DataTable dt = ConvertDStoDT(dsData, "data");

                            for (int i = 0; i < dt?.Rows.Count; i++)
                            {
                                string action_type = dt.Rows[i]["action_type"]?.ToString() ?? "";
                                sqlstr = "";
                                List<SqlParameter> parameters = new List<SqlParameter>();

                                if (action_type == "insert")
                                {
                                    sqlstr = "INSERT INTO EPHA_PERSON_DETAILS " +
                                             "(SEQ, ID, NO, NO_DEVIATIONS, NO_GUIDE_WORDS, DEVIATIONS, GUIDE_WORDS, PROCESS_DEVIATION, AREA_APPLICATION, PARAMETER, ACTIVE_TYPE, DEF_SELECTED, " +
                                             "CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                             "VALUES (@SEQ, @ID, @NO, @NO_DEVIATIONS, @NO_GUIDE_WORDS, @DEVIATIONS, @GUIDE_WORDS, @PROCESS_DEVIATION, @AREA_APPLICATION, @PARAMETER, @ACTIVE_TYPE, @DEF_SELECTED, " +
                                             "GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                                    parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = dt.Rows[i]["NO"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@NO_DEVIATIONS", SqlDbType.Int) { Value = dt.Rows[i]["NO_DEVIATIONS"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@NO_GUIDE_WORDS", SqlDbType.Int) { Value = dt.Rows[i]["NO_GUIDE_WORDS"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@DEVIATIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DEVIATIONS"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@GUIDE_WORDS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["GUIDE_WORDS"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@PROCESS_DEVIATION", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["PROCESS_DEVIATION"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@AREA_APPLICATION", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["AREA_APPLICATION"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@PARAMETER", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["PARAMETER"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@DEF_SELECTED", SqlDbType.Int) { Value = dt.Rows[i]["DEF_SELECTED"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["CREATE_BY"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"] ?? DBNull.Value });

                                    seq_now += 1;
                                }
                                else if (action_type == "update")
                                {
                                    seq_now = Convert.ToInt32(dt.Rows[i]["seq"]?.ToString() ?? "0");

                                    sqlstr = "UPDATE EPHA_M_GUIDE_WORDS SET " +
                                             "NO = @NO, NO_DEVIATIONS = @NO_DEVIATIONS, NO_GUIDE_WORDS = @NO_GUIDE_WORDS, " +
                                             "DEVIATIONS = @DEVIATIONS, GUIDE_WORDS = @GUIDE_WORDS, PROCESS_DEVIATION = @PROCESS_DEVIATION, AREA_APPLICATION = @AREA_APPLICATION, PARAMETER = @PARAMETER, " +
                                             "ACTIVE_TYPE = @ACTIVE_TYPE, DEF_SELECTED = @DEF_SELECTED, UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                             "WHERE SEQ = @SEQ AND ID = @ID";

                                    parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = dt.Rows[i]["NO"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@NO_DEVIATIONS", SqlDbType.Int) { Value = dt.Rows[i]["NO_DEVIATIONS"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@NO_GUIDE_WORDS", SqlDbType.Int) { Value = dt.Rows[i]["NO_GUIDE_WORDS"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@DEVIATIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DEVIATIONS"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@GUIDE_WORDS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["GUIDE_WORDS"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@PROCESS_DEVIATION", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["PROCESS_DEVIATION"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@AREA_APPLICATION", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["AREA_APPLICATION"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@PARAMETER", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["PARAMETER"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@DEF_SELECTED", SqlDbType.Int) { Value = dt.Rows[i]["DEF_SELECTED"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"] ?? DBNull.Value });
                                }
                                else if (action_type == "delete")
                                {
                                    sqlstr = "DELETE FROM EPHA_M_GUIDE_WORDS WHERE SEQ = @SEQ AND ID = @ID";

                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"] ?? DBNull.Value });
                                }

                                if (!string.IsNullOrEmpty(sqlstr))
                                {
                                    ClassConnectionDb cls_conn = new ClassConnectionDb();
                                    ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                                    if (ret != "true") { break; }
                                }
                            }

                            if (ret == "true")
                            {
                                transaction.Commit();
                            }
                            else
                            {
                                transaction.Rollback();
                            }
                        }
                        catch (Exception ex)
                        {
                            transaction.Rollback();
                            ret = "error: " + ex.Message;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return ret;
        }

        #endregion Manage User

        #region Company, Department and Sections

        //get_master_systemwide
        public string get_master_company(LoadMasterPageModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = (param.user_name + "").Trim();

            int iMaxSeq = get_max("epha_m_company");

            #region company
            sqlstr = @" select * from epha_m_company t order by id ";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            dt.TableName = "company";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion company

            #region Departments
            sqlstr = @" select distinct departments as id, functions +'-'+ departments as name, lower(a.departments) as text_check 
                        from vw_epha_person_details a
                        where isnull(functions,'') <> '' and isnull(departments,'') <> ''
                        order by   functions +'-'+ departments";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            dt.TableName = "departments";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion Departments

            #region Sections
            sqlstr = @"select distinct emp.sections as id, emp.sections as name,
                       emp.functions, emp.departments, emp.sections
                       from vw_epha_person_details emp
                       where isnull(emp.functions,'') <> '' and isnull(emp.departments,'') <> '' and isnull(emp.sections,'') <> '' 
                       order by emp.functions, emp.departments, emp.sections";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            dt.TableName = "sections";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion Sections

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;
        }
        public string get_master_area(LoadMasterPageModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = (param.user_name + "").Trim();

            int iMaxSeq = get_max("epha_m_area");

            #region Area Process Unit
            sqlstr = @" select a.*, 'update' as action_type, 0 as action_change  
                        from epha_m_area a 
                        order by a.id";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            if (dt?.Rows.Count == 0)
            {
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = iMaxSeq;
                dt.Rows[0]["id"] = iMaxSeq;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;

                dt.AcceptChanges();
            }
            dt.TableName = "area";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion Area Process Unit

            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "seq", (iMaxSeq + 1).ToString());
            dtma.TableName = "max";

            dsData.Tables.Add(dtma.Copy()); dsData.AcceptChanges();

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;
        }
        public string get_master_toc(LoadMasterPageModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = (param.user_name + "").Trim();

            int iMaxSeq = get_max("epha_m_area_complex");
            int id_company = 1;
            int id_area = 1;

            #region Plant
            sqlstr = @" select t.seq as id, t.plant as name from epha_m_company t order by t.plant ";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            dt.TableName = "plant";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion Plant

            #region Area Process Unit
            sqlstr = @" select a.id, a.name, a.name as area_check  
                        from epha_m_area a 
                        order by a.id";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            dt.TableName = "area";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion Area Process Unit

            #region Complex
            sqlstr = @"select t.*, 'update' as action_type, 0 as action_change
                       from epha_m_area_complex t 
                       order by t.id";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            if (dt?.Rows.Count == 0)
            {
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = iMaxSeq;
                dt.Rows[0]["id"] = iMaxSeq;
                dt.Rows[0]["id_company"] = id_company;
                dt.Rows[0]["id_area"] = id_area;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;

                dt.AcceptChanges();
            }
            dt.TableName = "toc";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            #endregion Complex

            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "seq", (iMaxSeq + 1).ToString());
            dtma.TableName = "max";

            dsData.Tables.Add(dtma.Copy()); dsData.AcceptChanges();

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;
        }
        public string get_master_unit(LoadMasterPageModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = (param.user_name + "").Trim();

            int iMaxSeq = get_max("epha_m_business_unit");
            int id_company = 1;
            int id_area = 1;
            int id_plant_area = 1;

            #region Plant
            sqlstr = @" select t.seq as id, t.plant as name from epha_m_company t order by t.plant ";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            dt.TableName = "plant";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion Plant

            #region Area Process Unit
            sqlstr = @" select a.id, a.name, a.name as area_check  
                        from epha_m_area a 
                        order by a.id";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            dt.TableName = "area";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion Area Process Unit

            #region Complex
            sqlstr = @" select a.id, a.name, a.name as area_check, a.id_company, a.id_area
                        from epha_m_area_complex a 
                        order by a.id";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            dt.TableName = "toc";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion Complex

            #region Unit No
            sqlstr = @"select t.*, 'update' as action_type, 0 as action_change
                       from epha_m_business_unit t 
                       order by t.id";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            if (dt?.Rows.Count == 0)
            {
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = iMaxSeq;
                dt.Rows[0]["id"] = iMaxSeq;
                dt.Rows[0]["id_company"] = id_company;
                dt.Rows[0]["id_area"] = id_area;
                dt.Rows[0]["id_plant_area"] = id_plant_area;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;

                dt.AcceptChanges();
            }
            dt.TableName = "unit";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            #endregion Unit No

            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "seq", (iMaxSeq + 1).ToString());
            dtma.TableName = "max";

            dsData.Tables.Add(dtma.Copy()); dsData.AcceptChanges();

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;
        }

        public string set_master_systemwide(SetDataMasterModel param)
        {
            string role_type = (param.role_type + "");
            string user_name = (param.user_name + "");
            string table_name = param.page_name ?? "";

            string ret = "";
            string msg = "";

            dsData = new DataSet();
            ConvertJSONresultToDataSet(ref msg, ref ret, ref dsData, param, table_name);
            if (ret.ToLower() == "error") { goto Next_Line_Convert; }

            try
            {
                using (var scope = new TransactionScope())
                {
                    if (table_name == "area") { ret = set_master_area(dsData, table_name, ref ret); }
                    else if (table_name == "toc") { ret = set_master_toc(dsData, table_name, ref ret); }
                    else if (table_name == "unit") { ret = set_master_unit(dsData, table_name, ref ret); }

                    if (ret == "true")
                    {
                        scope.Complete();
                    }
                    else
                    {
                        ret = "error";
                    }
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

        Next_Line_Convert:;
            string json = ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
            return json;
        }

        public string set_master_area(DataSet dsData, string json_table_name, ref string ret)
        {
            #region update data 
            int seq_now = Convert.ToInt32(get_max("epha_m_area").ToString() ?? "1");
            DataTable dt = ConvertDStoDT(dsData, json_table_name);

            foreach (DataRow row in dt.Rows)
            {
                string action_type = row["action_type"]?.ToString() ?? "";
                string sqlstr = "";
                List<SqlParameter> parameters = new List<SqlParameter>();

                if (action_type == "insert")
                {
                    #region insert 
                    sqlstr = "INSERT INTO EPHA_M_AREA (SEQ, ID, NAME, DESCRIPTIONS, ACTIVE_TYPE, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                             "VALUES (@SEQ, @ID, @NAME, @DESCRIPTIONS, @ACTIVE_TYPE, GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                    parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = row["NAME"] ?? DBNull.Value });
                    parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = row["DESCRIPTIONS"] ?? DBNull.Value });
                    parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = row["ACTIVE_TYPE"] ?? DBNull.Value });
                    parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = row["CREATE_BY"] ?? DBNull.Value });
                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] ?? DBNull.Value });

                    seq_now += 1;
                    #endregion insert
                }
                else if (action_type == "update")
                {
                    seq_now = Convert.ToInt32(row["seq"]?.ToString() ?? "0");

                    #region update
                    sqlstr = "UPDATE EPHA_M_AREA SET " +
                             "ID = @ID, NAME = @NAME, DESCRIPTIONS = @DESCRIPTIONS, ACTIVE_TYPE = @ACTIVE_TYPE, " +
                             "UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                             "WHERE SEQ = @SEQ";

                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] ?? DBNull.Value });
                    parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = row["NAME"] ?? DBNull.Value });
                    parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = row["DESCRIPTIONS"] ?? DBNull.Value });
                    parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = row["ACTIVE_TYPE"] ?? DBNull.Value });
                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] ?? DBNull.Value });
                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] ?? DBNull.Value });
                    #endregion update
                }
                else if (action_type == "delete")
                {
                    #region delete
                    sqlstr = "DELETE FROM EPHA_M_AREA WHERE SEQ = @SEQ";

                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] ?? DBNull.Value });
                    #endregion delete
                }

                if (!string.IsNullOrEmpty(sqlstr))
                {
                    ClassConnectionDb cls_conn = new ClassConnectionDb();
                    ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, null, null);
                    if (ret != "true") break;
                }
            }
            #endregion update data 

            return ret;
        }

        public string set_master_toc(DataSet dsData, string json_table_name, ref string ret)
        {
            #region update data 
            int seq_now = Convert.ToInt32(get_max("epha_m_area_complex").ToString() ?? "1");
            DataTable dt = ConvertDStoDT(dsData, json_table_name);

            foreach (DataRow row in dt.Rows)
            {
                string action_type = row["action_type"]?.ToString() ?? "";
                string sqlstr = "";
                List<SqlParameter> parameters = new List<SqlParameter>();

                if (action_type == "insert")
                {
                    #region insert 
                    sqlstr = "INSERT INTO epha_m_area_complex (SEQ, ID, ID_COMPANY, ID_AREA, NAME, DESCRIPTIONS, ACTIVE_TYPE, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                             "VALUES (@SEQ, @ID, @ID_COMPANY, @ID_AREA, @NAME, @DESCRIPTIONS, @ACTIVE_TYPE, GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                    parameters.Add(new SqlParameter("@ID_COMPANY", SqlDbType.Int) { Value = row["ID_COMPANY"] ?? DBNull.Value });
                    parameters.Add(new SqlParameter("@ID_AREA", SqlDbType.Int) { Value = row["ID_AREA"] ?? DBNull.Value });
                    parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = row["NAME"] ?? DBNull.Value });
                    parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = row["DESCRIPTIONS"] ?? DBNull.Value });
                    parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = row["ACTIVE_TYPE"] ?? DBNull.Value });
                    parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = row["CREATE_BY"] ?? DBNull.Value });
                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] ?? DBNull.Value });

                    seq_now += 1;
                    #endregion insert
                }
                else if (action_type == "update")
                {
                    seq_now = Convert.ToInt32(row["seq"]?.ToString() ?? "0");

                    #region update
                    sqlstr = "UPDATE epha_m_area_complex SET " +
                             "ID = @ID, NAME = @NAME, DESCRIPTIONS = @DESCRIPTIONS, ACTIVE_TYPE = @ACTIVE_TYPE, " +
                             "UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                             "WHERE SEQ = @SEQ";

                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] ?? DBNull.Value });
                    parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = row["NAME"] ?? DBNull.Value });
                    parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = row["DESCRIPTIONS"] ?? DBNull.Value });
                    parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = row["ACTIVE_TYPE"] ?? DBNull.Value });
                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] ?? DBNull.Value });
                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] ?? DBNull.Value });
                    #endregion update
                }
                else if (action_type == "delete")
                {
                    #region delete
                    sqlstr = "DELETE FROM epha_m_area_complex WHERE SEQ = @SEQ";

                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] ?? DBNull.Value });
                    #endregion delete
                }

                if (!string.IsNullOrEmpty(sqlstr))
                {
                    ClassConnectionDb cls_conn = new ClassConnectionDb();
                    ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, null, null);
                    if (ret != "true") break;
                }
            }
            #endregion update data 

            return ret;
        }

        public string set_master_unit(DataSet dsData, string json_table_name, ref string ret)
        {
            int seq_now = Convert.ToInt32(get_max("EPHA_M_BUSINESS_UNIT").ToString() ?? "1");

            #region update data 
            DataTable dt = ConvertDStoDT(dsData, json_table_name);

            foreach (DataRow row in dt.Rows)
            {
                string action_type = row["action_type"]?.ToString() ?? "";
                string sqlstr = "";
                List<SqlParameter> parameters = new List<SqlParameter>();

                if (action_type == "insert")
                {
                    #region insert 
                    sqlstr = "INSERT INTO EPHA_M_BUSINESS_UNIT (SEQ, ID, ID_COMPANY, ID_AREA, ID_PLANT_AREA, NAME, DESCRIPTIONS, ACTIVE_TYPE, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                             "VALUES (@SEQ, @ID, @ID_COMPANY, @ID_AREA, @ID_PLANT_AREA, @NAME, @DESCRIPTIONS, @ACTIVE_TYPE, GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                    parameters.Add(new SqlParameter("@ID_COMPANY", SqlDbType.Int) { Value = row["ID_COMPANY"] ?? DBNull.Value });
                    parameters.Add(new SqlParameter("@ID_AREA", SqlDbType.Int) { Value = row["ID_AREA"] ?? DBNull.Value });
                    parameters.Add(new SqlParameter("@ID_PLANT_AREA", SqlDbType.Int) { Value = row["ID_PLANT_AREA"] ?? DBNull.Value });
                    parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = row["NAME"] ?? DBNull.Value });
                    parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = row["DESCRIPTIONS"] ?? DBNull.Value });
                    parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = row["ACTIVE_TYPE"] ?? DBNull.Value });
                    parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = row["CREATE_BY"] ?? DBNull.Value });
                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] ?? DBNull.Value });

                    seq_now += 1;
                    #endregion insert
                }
                else if (action_type == "update")
                {
                    seq_now = Convert.ToInt32(row["seq"]?.ToString() ?? "0");

                    #region update
                    sqlstr = "UPDATE EPHA_M_BUSINESS_UNIT SET " +
                             "ID = @ID, NAME = @NAME, DESCRIPTIONS = @DESCRIPTIONS, ACTIVE_TYPE = @ACTIVE_TYPE, " +
                             "UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                             "WHERE SEQ = @SEQ";

                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] ?? DBNull.Value });
                    parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = row["NAME"] ?? DBNull.Value });
                    parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = row["DESCRIPTIONS"] ?? DBNull.Value });
                    parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = row["ACTIVE_TYPE"] ?? DBNull.Value });
                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] ?? DBNull.Value });
                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] ?? DBNull.Value });
                    #endregion update
                }
                else if (action_type == "delete")
                {
                    #region delete
                    sqlstr = "DELETE FROM EPHA_M_BUSINESS_UNIT WHERE SEQ = @SEQ";

                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] ?? DBNull.Value });
                    #endregion delete
                }

                if (!string.IsNullOrEmpty(sqlstr))
                {
                    ClassConnectionDb cls_conn = new ClassConnectionDb();
                    ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, null, null);
                    if (ret != "true") break;
                }
            }
            #endregion update data 

            return ret;
        }

        #endregion  Company, Department and Sections

        #region HAZOP Module : Functional Location
        public string get_master_functionallocation(LoadMasterPageModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = (param.user_name + "").Trim();


            int iMaxSeq = get_max("epha_m_functional_location");

            sqlstr = @" select a.*, 'update' as action_type, 0 as action_change, case when a.seq = 0 then 1 else 0 end disable_page
                        from epha_m_functional_location a 
                        order by seq ";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            if (dt?.Rows.Count == 0)
            {
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = iMaxSeq;
                dt.Rows[0]["id"] = iMaxSeq;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();

                iMaxSeq += 1;
            }
            dt.AcceptChanges();

            dt.TableName = "data";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            #region drawing 
            sqlstr = @" select a.* , 'update' as action_type, 0 as action_change
                        from epha_m_drawing a
                        where a.module = 'functional_location' ";
            sqlstr += " order by a.seq ";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            int id_drawing = get_max("epha_m_drawing");

            if (dt?.Rows.Count == 0)
            {
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = id_drawing;
                dt.Rows[0]["id"] = id_drawing;

                dt.Rows[0]["module"] = "functional_location";

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();
            }
            dt.TableName = "drawing";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion drawing

            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "seq", (iMaxSeq + 1).ToString());
            clshazop.set_max_id(ref dtma, "drawing", (id_drawing + 1).ToString());
            dtma.TableName = "max";

            dsData.Tables.Add(dtma.Copy()); dsData.AcceptChanges();

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;

        }
        public string set_master_functionallocation(SetMasterGuideWordsModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();

            int iMaxSeq = 0;
            string msg = "";
            string ret = "";
            string user_name = (param.user_name + "").Trim();
            string json_data = (param.json_data + "").Trim();

            DataMasterListModel param_def = new DataMasterListModel();
            param_def.json_name = "data";
            param_def.json_data = param.json_data;
            ConvertJSONListresultToDataSet(ref msg, ref ret, ref dsData, param_def);
            if (ret.ToLower() == "error") { goto Next_Line_Convert; }

            param_def = new DataMasterListModel();
            param_def.json_name = "drawing";
            param_def.json_data = param.json_drawing;
            ConvertJSONListresultToDataSet(ref msg, ref ret, ref dsData, param_def);
            if (ret.ToLower() == "error") { goto Next_Line_Convert; }

            iMaxSeq = get_max("epha_m_functional_location");
            int iMaxSeqDrawing = get_max("epha_m_drawing");

            try
            {
                using (TransactionScope scope = new TransactionScope())
                {
                    ret = set_functional_location(ref dsData, ref iMaxSeq);
                    if (ret != "true") throw new Exception(ret);

                    ret = set_master_drawing(ref dsData, ref iMaxSeqDrawing);
                    if (ret != "true") throw new Exception(ret);

                    scope.Complete();
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

        Next_Line_Convert:;

            return cls_json.SetJSONresult(refMsgSaveMaster(ret, msg, (iMaxSeq.ToString() + "").ToString()));
        }

        public string set_functional_location(ref DataSet dsData, ref int seq_now)
        {
            string ret = "";

            dt = new DataTable();
            dt = ConvertDStoDT(dsData, "data");

            using (SqlConnection conn = new SqlConnection(ClassConnectionDb.ConnectionString()))
            {
                conn.Open();
                using (SqlTransaction transaction = conn.BeginTransaction())
                {
                    try
                    {
                        foreach (DataRow row in dt.Rows)
                        {
                            string action_type = row["action_type"]?.ToString() ?? "";
                            List<SqlParameter> parameters = new List<SqlParameter>();
                            string sqlstr = "";

                            if (!string.IsNullOrEmpty(action_type))
                            {
                                if (action_type == "insert")
                                {
                                    sqlstr = "insert into EPHA_M_FUNCTIONAL_LOCATION " +
                                             "(SEQ, ID, NOTIF_DATE, NOTIF_TIME, NOTIFICATION, ORDERS, TYP, P, PLS, FUNCTIONAL_LOCATION, DESCRIPTIONS_FUNC, DESCRIPTIONS, " +
                                             "MN_WK, PLNT, REPORTED_BY, REQUIRED_START, REQUIRED_END, USER_STATUS, DATE_UPDATE, ACTIVE_TYPE, CREATE_DATE, CREATE_BY, UPDATE_BY) " +
                                             "values (@SEQ, @ID, @NOTIF_DATE, @NOTIF_TIME, @NOTIFICATION, @ORDERS, @TYP, @P, @PLS, @FUNCTIONAL_LOCATION, @DESCRIPTIONS_FUNC, @DESCRIPTIONS, " +
                                             "@MN_WK, @PLNT, @REPORTED_BY, @REQUIRED_START, @REQUIRED_END, @USER_STATUS, @DATE_UPDATE, @ACTIVE_TYPE, getdate(), @CREATE_BY, @UPDATE_BY)";

                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                                    parameters.Add(new SqlParameter("@NOTIF_DATE", SqlDbType.NVarChar, 4000) { Value = row["NOTIF_DATE"] });
                                    parameters.Add(new SqlParameter("@NOTIF_TIME", SqlDbType.NVarChar, 4000) { Value = row["NOTIF_TIME"] });
                                    parameters.Add(new SqlParameter("@NOTIFICATION", SqlDbType.NVarChar, 4000) { Value = row["NOTIFICATION"] });
                                    parameters.Add(new SqlParameter("@ORDERS", SqlDbType.NVarChar, 4000) { Value = row["ORDERS"] });
                                    parameters.Add(new SqlParameter("@TYP", SqlDbType.NVarChar, 4000) { Value = row["TYP"] });
                                    parameters.Add(new SqlParameter("@P", SqlDbType.NVarChar, 4000) { Value = row["P"] });
                                    parameters.Add(new SqlParameter("@PLS", SqlDbType.NVarChar, 4000) { Value = row["PLS"] });
                                    parameters.Add(new SqlParameter("@FUNCTIONAL_LOCATION", SqlDbType.NVarChar, 4000) { Value = row["FUNCTIONAL_LOCATION"] });
                                    parameters.Add(new SqlParameter("@DESCRIPTIONS_FUNC", SqlDbType.NVarChar, 4000) { Value = row["DESCRIPTIONS_FUNC"] });
                                    parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = row["DESCRIPTIONS"] });
                                    parameters.Add(new SqlParameter("@MN_WK", SqlDbType.NVarChar, 4000) { Value = row["MN_WK"] });
                                    parameters.Add(new SqlParameter("@PLNT", SqlDbType.NVarChar, 4000) { Value = row["PLNT"] });
                                    parameters.Add(new SqlParameter("@REPORTED_BY", SqlDbType.NVarChar, 4000) { Value = row["REPORTED_BY"] });
                                    parameters.Add(new SqlParameter("@REQUIRED_START", SqlDbType.NVarChar, 4000) { Value = row["REQUIRED_START"] });
                                    parameters.Add(new SqlParameter("@REQUIRED_END", SqlDbType.NVarChar, 4000) { Value = row["REQUIRED_END"] });
                                    parameters.Add(new SqlParameter("@USER_STATUS", SqlDbType.NVarChar, 4000) { Value = row["USER_STATUS"] });
                                    parameters.Add(new SqlParameter("@DATE_UPDATE", SqlDbType.NVarChar, 4000) { Value = row["DATE_UPDATE"] });
                                    parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = row["ACTIVE_TYPE"] });
                                    parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = row["CREATE_BY"] });
                                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] });

                                    seq_now++;
                                }
                                else if (action_type == "update")
                                {
                                    sqlstr = "update EPHA_M_FUNCTIONAL_LOCATION set " +
                                             "NOTIF_DATE = @NOTIF_DATE, NOTIF_TIME = @NOTIF_TIME, NOTIFICATION = @NOTIFICATION, ORDERS = @ORDERS, TYP = @TYP, P = @P, PLS = @PLS, " +
                                             "FUNCTIONAL_LOCATION = @FUNCTIONAL_LOCATION, DESCRIPTIONS_FUNC = @DESCRIPTIONS_FUNC, DESCRIPTIONS = @DESCRIPTIONS, " +
                                             "MN_WK = @MN_WK, PLNT = @PLNT, REPORTED_BY = @REPORTED_BY, REQUIRED_START = @REQUIRED_START, REQUIRED_END = @REQUIRED_END, " +
                                             "USER_STATUS = @USER_STATUS, DATE_UPDATE = @DATE_UPDATE, ACTIVE_TYPE = @ACTIVE_TYPE, UPDATE_DATE = getdate(), UPDATE_BY = @UPDATE_BY " +
                                             "where SEQ = @SEQ and ID = @ID";

                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] });
                                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] });
                                    parameters.Add(new SqlParameter("@NOTIF_DATE", SqlDbType.NVarChar, 4000) { Value = row["NOTIF_DATE"] });
                                    parameters.Add(new SqlParameter("@NOTIF_TIME", SqlDbType.NVarChar, 4000) { Value = row["NOTIF_TIME"] });
                                    parameters.Add(new SqlParameter("@NOTIFICATION", SqlDbType.NVarChar, 4000) { Value = row["NOTIFICATION"] });
                                    parameters.Add(new SqlParameter("@ORDERS", SqlDbType.NVarChar, 4000) { Value = row["ORDERS"] });
                                    parameters.Add(new SqlParameter("@TYP", SqlDbType.NVarChar, 4000) { Value = row["TYP"] });
                                    parameters.Add(new SqlParameter("@P", SqlDbType.NVarChar, 4000) { Value = row["P"] });
                                    parameters.Add(new SqlParameter("@PLS", SqlDbType.NVarChar, 4000) { Value = row["PLS"] });
                                    parameters.Add(new SqlParameter("@FUNCTIONAL_LOCATION", SqlDbType.NVarChar, 4000) { Value = row["FUNCTIONAL_LOCATION"] });
                                    parameters.Add(new SqlParameter("@DESCRIPTIONS_FUNC", SqlDbType.NVarChar, 4000) { Value = row["DESCRIPTIONS_FUNC"] });
                                    parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = row["DESCRIPTIONS"] });
                                    parameters.Add(new SqlParameter("@MN_WK", SqlDbType.NVarChar, 4000) { Value = row["MN_WK"] });
                                    parameters.Add(new SqlParameter("@PLNT", SqlDbType.NVarChar, 4000) { Value = row["PLNT"] });
                                    parameters.Add(new SqlParameter("@REPORTED_BY", SqlDbType.NVarChar, 4000) { Value = row["REPORTED_BY"] });
                                    parameters.Add(new SqlParameter("@REQUIRED_START", SqlDbType.NVarChar, 4000) { Value = row["REQUIRED_START"] });
                                    parameters.Add(new SqlParameter("@REQUIRED_END", SqlDbType.NVarChar, 4000) { Value = row["REQUIRED_END"] });
                                    parameters.Add(new SqlParameter("@USER_STATUS", SqlDbType.NVarChar, 4000) { Value = row["USER_STATUS"] });
                                    parameters.Add(new SqlParameter("@DATE_UPDATE", SqlDbType.NVarChar, 4000) { Value = row["DATE_UPDATE"] });
                                    parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = row["ACTIVE_TYPE"] });
                                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] });
                                }
                                else if (action_type == "delete")
                                {
                                    sqlstr = "delete from EPHA_M_FUNCTIONAL_LOCATION where SEQ = @SEQ and ID = @ID";

                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] });
                                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] });
                                }

                                if (!string.IsNullOrEmpty(sqlstr))
                                {
                                    ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                                    if (ret != "true") throw new Exception(ret);
                                }
                            }
                        }

                        transaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        ret = "error: " + ex.Message;
                    }
                }
            }

            return ret;
        }

        public string set_master_drawing(ref DataSet dsData, ref int seq_now)
        {
            string ret = "";

            if (dsData.Tables["drawing"] != null)
            {
                dt = new DataTable();
                dt = dsData?.Tables["drawing"]?.Copy() ?? new DataTable();
                dt.AcceptChanges();

                using (SqlConnection conn = new SqlConnection(ClassConnectionDb.ConnectionString()))
                {
                    conn.Open();
                    using (SqlTransaction transaction = conn.BeginTransaction())
                    {
                        try
                        {
                            foreach (DataRow row in dt.Rows)
                            {
                                string action_type = row["action_type"].ToString();
                                List<SqlParameter> parameters = new List<SqlParameter>();
                                string sqlstr = "";

                                if (!string.IsNullOrEmpty(action_type))
                                {
                                    if (action_type == "insert")
                                    {
                                        sqlstr = "insert into EPHA_M_DRAWING " +
                                                 "(SEQ, ID, MODULE, NAME, DESCRIPTIONS, DOCUMENT_FILE_SIZE, DOCUMENT_FILE_PATH, DOCUMENT_FILE_NAME, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                                 "values (@SEQ, @ID, @MODULE, @NAME, @DESCRIPTIONS, @DOCUMENT_FILE_SIZE, @DOCUMENT_FILE_PATH, @DOCUMENT_FILE_NAME, getdate(), null, @CREATE_BY, @UPDATE_BY)";

                                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                                        parameters.Add(new SqlParameter("@MODULE", SqlDbType.NVarChar, 4000) { Value = row["MODULE"] });
                                        parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = row["NAME"] });
                                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = row["DESCRIPTIONS"] });
                                        parameters.Add(new SqlParameter("@DOCUMENT_FILE_SIZE", SqlDbType.Int) { Value = row["DOCUMENT_FILE_SIZE"] });
                                        parameters.Add(new SqlParameter("@DOCUMENT_FILE_PATH", SqlDbType.NVarChar, 4000) { Value = row["DOCUMENT_FILE_PATH"] });
                                        parameters.Add(new SqlParameter("@DOCUMENT_FILE_NAME", SqlDbType.NVarChar, 4000) { Value = row["DOCUMENT_FILE_NAME"] });
                                        parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = row["CREATE_BY"] });
                                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] });

                                        seq_now++;
                                    }
                                    else if (action_type == "update")
                                    {
                                        sqlstr = "update EPHA_M_DRAWING set " +
                                                 "MODULE = @MODULE, NAME = @NAME, DESCRIPTIONS = @DESCRIPTIONS, DOCUMENT_FILE_SIZE = @DOCUMENT_FILE_SIZE, " +
                                                 "DOCUMENT_FILE_PATH = @DOCUMENT_FILE_PATH, DOCUMENT_FILE_NAME = @DOCUMENT_FILE_NAME, " +
                                                 "UPDATE_DATE = getdate(), UPDATE_BY = @UPDATE_BY " +
                                                 "where SEQ = @SEQ and ID = @ID";

                                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] });
                                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] });
                                        parameters.Add(new SqlParameter("@MODULE", SqlDbType.NVarChar, 4000) { Value = row["MODULE"] });
                                        parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = row["NAME"] });
                                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = row["DESCRIPTIONS"] });
                                        parameters.Add(new SqlParameter("@DOCUMENT_FILE_SIZE", SqlDbType.Int) { Value = row["DOCUMENT_FILE_SIZE"] });
                                        parameters.Add(new SqlParameter("@DOCUMENT_FILE_PATH", SqlDbType.NVarChar, 4000) { Value = row["DOCUMENT_FILE_PATH"] });
                                        parameters.Add(new SqlParameter("@DOCUMENT_FILE_NAME", SqlDbType.NVarChar, 4000) { Value = row["DOCUMENT_FILE_NAME"] });
                                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] });
                                    }
                                    else if (action_type == "delete")
                                    {
                                        sqlstr = "delete from EPHA_M_DRAWING where SEQ = @SEQ and ID = @ID";

                                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] });
                                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] });
                                    }

                                    if (!string.IsNullOrEmpty(sqlstr))
                                    {
                                        ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                                        if (ret != "true") throw new Exception(ret);
                                    }
                                }
                            }

                            transaction.Commit();
                        }
                        catch (Exception ex)
                        {
                            transaction.Rollback();
                            ret = "error: " + ex.Message;
                        }
                    }
                }
            }

            return ret;
        }


        #endregion HAZOP Module : Functional Location

        #region HAZOP Module : Guide Words 
        public string get_master_guidewords(LoadMasterPageModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = (param.user_name + "").Trim();

            #region master data in page

            sqlstr = @" select p.id, p.name 
                        from epha_m_parameter p 
                        where p.active_type = 1
                        order by p.id";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            dt.TableName = "parameter";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            sqlstr = @" select a.id, a.name 
                        , a.id_parameter, p.name as parameter
                        from epha_m_area_application a 
                        inner join epha_m_parameter p on a.id_parameter = p.id
                        where a.active_type = 1
                        order by a.id";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            dt.TableName = "area_application";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();


            #endregion master data in page


            int iMaxSeq = get_max("epha_m_guide_words");

            sqlstr = @" select a.*, 'update' as action_type, 0 as action_change, case when a.seq = 0 then 1 else 0 end disable_page
                        from epha_m_guide_words a 
                        order by seq ";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            if (dt?.Rows.Count == 0)
            {
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = iMaxSeq;
                dt.Rows[0]["id"] = iMaxSeq;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();

                iMaxSeq += 1;
            }
            dt.AcceptChanges();

            dt.TableName = "data";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            #region drawing 
            sqlstr = @" select a.* , 'update' as action_type, 0 as action_change
                        from epha_m_drawing a
                        where a.module = 'guide_words' ";
            sqlstr += " order by a.seq ";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            int id_drawing = get_max("epha_m_drawing");

            if (dt?.Rows.Count == 0)
            {
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = id_drawing;
                dt.Rows[0]["id"] = id_drawing;

                dt.Rows[0]["module"] = "guide_words";

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();
            }
            dt.TableName = "drawing";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion drawing

            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "seq", (iMaxSeq + 1).ToString());
            clshazop.set_max_id(ref dtma, "drawing", (id_drawing + 1).ToString());
            dtma.TableName = "max";

            dsData.Tables.Add(dtma.Copy()); dsData.AcceptChanges();

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;

        }
        public string set_master_guidewords(SetMasterGuideWordsModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();

            int iMaxSeq = 0;
            string msg = "";
            string ret = "";
            string user_name = (param.user_name ?? "").Trim();
            string json_data = (param.json_data ?? "").Trim();

            DataMasterListModel param_def = new DataMasterListModel
            {
                json_name = "data",
                json_data = param.json_data
            };
            ConvertJSONListresultToDataSet(ref msg, ref ret, ref dsData, param_def);
            if (ret.ToLower() == "error") { goto Next_Line; }

            param_def = new DataMasterListModel
            {
                json_name = "drawing",
                json_data = param.json_drawing
            };
            ConvertJSONListresultToDataSet(ref msg, ref ret, ref dsData, param_def);
            if (ret.ToLower() == "error") { goto Next_Line; }

            iMaxSeq = get_max("epha_m_guide_words");
            int iMaxSeqDrawing = get_max("epha_m_drawing");

            using (TransactionScope scope = new TransactionScope())
            using (SqlConnection conn = new SqlConnection(ClassConnectionDb.ConnectionString()))
            {
                conn.Open();

                ret = set_guidewords(ref dsData, ref iMaxSeq, conn);

                if (ret == "true")
                {
                    ret = set_master_drawing(ref dsData, ref iMaxSeqDrawing, conn);
                }

                if (ret == "true")
                {
                    scope.Complete();
                }
            }

        Next_Line:;
            return cls_json.SetJSONresult(refMsgSaveMaster(ret, msg, (iMaxSeq.ToString() + "").ToString()));
        }

        public string set_guidewords(ref DataSet dsData, ref int seq_now, SqlConnection conn)
        {
            string ret = "";

            #region update data
            dt = new DataTable();
            dt = ConvertDStoDT(dsData, "data");

            for (int i = 0; i < dt?.Rows.Count; i++)
            {
                string action_type = (dt.Rows[i]["action_type"]?.ToString() ?? "").ToLower();
                string sqlstr = "";
                List<SqlParameter> parameters = new List<SqlParameter>();

                if (action_type == "insert")
                {
                    #region insert
                    sqlstr = "insert into EPHA_M_GUIDE_WORDS " +
                             "(SEQ, ID, NO, NO_DEVIATIONS, NO_GUIDE_WORDS, DEVIATIONS, GUIDE_WORDS, PROCESS_DEVIATION, AREA_APPLICATION, PARAMETER, ACTIVE_TYPE, DEF_SELECTED, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                             "values (@SEQ, @ID, @NO, @NO_DEVIATIONS, @NO_GUIDE_WORDS, @DEVIATIONS, @GUIDE_WORDS, @PROCESS_DEVIATION, @AREA_APPLICATION, @PARAMETER, @ACTIVE_TYPE, @DEF_SELECTED, getdate(), null, @CREATE_BY, @UPDATE_BY)";

                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                    parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = dt.Rows[i]["NO"]?.ToString() ?? "" });
                    parameters.Add(new SqlParameter("@NO_DEVIATIONS", SqlDbType.Int) { Value = dt.Rows[i]["NO_DEVIATIONS"]?.ToString() ?? "" });
                    parameters.Add(new SqlParameter("@NO_GUIDE_WORDS", SqlDbType.Int) { Value = dt.Rows[i]["NO_GUIDE_WORDS"]?.ToString() ?? "" });
                    parameters.Add(new SqlParameter("@DEVIATIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DEVIATIONS"]?.ToString() ?? "" });
                    parameters.Add(new SqlParameter("@GUIDE_WORDS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["GUIDE_WORDS"]?.ToString() ?? "" });
                    parameters.Add(new SqlParameter("@PROCESS_DEVIATION", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["PROCESS_DEVIATION"]?.ToString() ?? "" });
                    parameters.Add(new SqlParameter("@AREA_APPLICATION", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["AREA_APPLICATION"]?.ToString() ?? "" });
                    parameters.Add(new SqlParameter("@PARAMETER", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["PARAMETER"]?.ToString() ?? "" });
                    parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"]?.ToString() ?? "" });
                    parameters.Add(new SqlParameter("@DEF_SELECTED", SqlDbType.Int) { Value = dt.Rows[i]["DEF_SELECTED"]?.ToString() ?? "" });
                    parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["CREATE_BY"]?.ToString() ?? "" });
                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"]?.ToString() ?? "" });

                    seq_now += 1;
                    #endregion
                }
                else if (action_type == "update")
                {
                    #region update
                    sqlstr = "update EPHA_M_GUIDE_WORDS set " +
                             "NO = @NO, NO_DEVIATIONS = @NO_DEVIATIONS, NO_GUIDE_WORDS = @NO_GUIDE_WORDS, DEVIATIONS = @DEVIATIONS, GUIDE_WORDS = @GUIDE_WORDS, " +
                             "PROCESS_DEVIATION = @PROCESS_DEVIATION, AREA_APPLICATION = @AREA_APPLICATION, PARAMETER = @PARAMETER, ACTIVE_TYPE = @ACTIVE_TYPE, " +
                             "DEF_SELECTED = @DEF_SELECTED, UPDATE_DATE = getdate(), UPDATE_BY = @UPDATE_BY " +
                             "where SEQ = @SEQ and ID = @ID";

                    parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = dt.Rows[i]["NO"]?.ToString() ?? "" });
                    parameters.Add(new SqlParameter("@NO_DEVIATIONS", SqlDbType.Int) { Value = dt.Rows[i]["NO_DEVIATIONS"]?.ToString() ?? "" });
                    parameters.Add(new SqlParameter("@NO_GUIDE_WORDS", SqlDbType.Int) { Value = dt.Rows[i]["NO_GUIDE_WORDS"]?.ToString() ?? "" });
                    parameters.Add(new SqlParameter("@DEVIATIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DEVIATIONS"]?.ToString() ?? "" });
                    parameters.Add(new SqlParameter("@GUIDE_WORDS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["GUIDE_WORDS"]?.ToString() ?? "" });
                    parameters.Add(new SqlParameter("@PROCESS_DEVIATION", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["PROCESS_DEVIATION"]?.ToString() ?? "" });
                    parameters.Add(new SqlParameter("@AREA_APPLICATION", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["AREA_APPLICATION"]?.ToString() ?? "" });
                    parameters.Add(new SqlParameter("@PARAMETER", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["PARAMETER"]?.ToString() ?? "" });
                    parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"]?.ToString() ?? "" });
                    parameters.Add(new SqlParameter("@DEF_SELECTED", SqlDbType.Int) { Value = dt.Rows[i]["DEF_SELECTED"]?.ToString() ?? "" });
                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"]?.ToString() ?? "" });
                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"]?.ToString() ?? "" });
                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"]?.ToString() ?? "" });
                    #endregion
                }
                else if (action_type == "delete")
                {
                    #region delete
                    sqlstr = "delete from EPHA_M_GUIDE_WORDS where SEQ = @SEQ and ID = @ID";

                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"]?.ToString() ?? "" });
                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"]?.ToString() ?? "" });
                    #endregion
                }

                if (action_type != "")
                {
                    ClassConnectionDb cls_conn = new ClassConnectionDb();
                    ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, null); // transaction is null because it's handled by TransactionScope
                    if (ret != "true") { throw new Exception("Database operation failed"); }
                }
            }

            #endregion update data

            return ret;
        }

        public string set_master_drawing(ref DataSet dsData, ref int seq_now, SqlConnection conn)
        {
            string ret = "";

            #region update data drawing
            if (dsData.Tables["drawing"] != null)
            {
                dt = new DataTable();
                dt = dsData.Tables["drawing"].Copy(); dt.AcceptChanges();

                for (int i = 0; i < dt?.Rows.Count; i++)
                {
                    string action_type = (dt.Rows[i]["action_type"]?.ToString() ?? "").ToLower();
                    string sqlstr = "";
                    List<SqlParameter> parameters = new List<SqlParameter>();

                    if (action_type == "insert")
                    {
                        #region insert
                        sqlstr = "insert into EPHA_M_DRAWING (" +
                                 "SEQ, ID, MODULE, NAME, DESCRIPTIONS, DOCUMENT_FILE_SIZE, DOCUMENT_FILE_PATH, DOCUMENT_FILE_NAME, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                 "values (@SEQ, @ID, @MODULE, @NAME, @DESCRIPTIONS, @DOCUMENT_FILE_SIZE, @DOCUMENT_FILE_PATH, @DOCUMENT_FILE_NAME, getdate(), null, @CREATE_BY, @UPDATE_BY)";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"]?.ToString() ?? "" });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"]?.ToString() ?? "" });
                        parameters.Add(new SqlParameter("@MODULE", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["MODULE"]?.ToString() ?? "" });
                        parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["NAME"]?.ToString() ?? "" });
                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTIONS"]?.ToString() ?? "" });
                        parameters.Add(new SqlParameter("@DOCUMENT_FILE_SIZE", SqlDbType.Int) { Value = dt.Rows[i]["DOCUMENT_FILE_SIZE"]?.ToString() ?? "" });
                        parameters.Add(new SqlParameter("@DOCUMENT_FILE_PATH", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DOCUMENT_FILE_PATH"]?.ToString() ?? "" });
                        parameters.Add(new SqlParameter("@DOCUMENT_FILE_NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DOCUMENT_FILE_NAME"]?.ToString() ?? "" });
                        parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["CREATE_BY"]?.ToString() ?? "" });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"]?.ToString() ?? "" });
                        #endregion
                    }
                    else if (action_type == "update")
                    {
                        #region update
                        sqlstr = "update EPHA_M_DRAWING set " +
                                 "MODULE = @MODULE, NAME = @NAME, DESCRIPTIONS = @DESCRIPTIONS, DOCUMENT_FILE_SIZE = @DOCUMENT_FILE_SIZE, " +
                                 "DOCUMENT_FILE_PATH = @DOCUMENT_FILE_PATH, DOCUMENT_FILE_NAME = @DOCUMENT_FILE_NAME, UPDATE_DATE = getdate(), UPDATE_BY = @UPDATE_BY " +
                                 "where SEQ = @SEQ and ID = @ID";

                        parameters.Add(new SqlParameter("@MODULE", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["MODULE"]?.ToString() ?? "" });
                        parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["NAME"]?.ToString() ?? "" });
                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTIONS"]?.ToString() ?? "" });
                        parameters.Add(new SqlParameter("@DOCUMENT_FILE_SIZE", SqlDbType.Int) { Value = dt.Rows[i]["DOCUMENT_FILE_SIZE"]?.ToString() ?? "" });
                        parameters.Add(new SqlParameter("@DOCUMENT_FILE_PATH", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DOCUMENT_FILE_PATH"]?.ToString() ?? "" });
                        parameters.Add(new SqlParameter("@DOCUMENT_FILE_NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DOCUMENT_FILE_NAME"]?.ToString() ?? "" });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"]?.ToString() ?? "" });
                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"]?.ToString() ?? "" });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"]?.ToString() ?? "" });
                        #endregion
                    }
                    else if (action_type == "delete")
                    {
                        #region delete
                        sqlstr = "delete from EPHA_M_DRAWING where SEQ = @SEQ and ID = @ID";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"]?.ToString() ?? "" });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"]?.ToString() ?? "" });
                        #endregion
                    }

                    if (action_type != "")
                    {
                        ClassConnectionDb cls_conn = new ClassConnectionDb();
                        ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, null); // transaction is null because it's handled by TransactionScope
                        if (ret != "true") { throw new Exception("Database operation failed"); }
                    }
                }

                if (ret != "true") { return ret; }
            }
            #endregion update data drawing

            return ret;
        }

        #endregion HAZOP Module : Guide Words 

        #region JSEA Module : Mandatory Note
        public string get_master_mandatorynote(LoadMasterPageModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = (param.user_name + "").Trim();

            int iMaxSeq = get_max("epha_m_mandatory_note");

            sqlstr = @" select a.*, 'update' as action_type, 0 as action_change
                        from epha_m_mandatory_note a 
                        order by seq ";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            if (dt?.Rows.Count == 0)
            {
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = iMaxSeq;
                dt.Rows[0]["id"] = iMaxSeq;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();

                iMaxSeq += 1;
            }
            dt.AcceptChanges();

            dt.TableName = "data";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "seq", (iMaxSeq + 1).ToString());
            dtma.TableName = "max";

            dsData.Tables.Add(dtma.Copy()); dsData.AcceptChanges();

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;

        }
        public string set_master_mandatorynote(SetDataMasterModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();

            int iMaxSeq = 0;
            string msg = "";
            string ret = "";
            string user_name = (param.user_name ?? "").Trim();
            string role_type = (param.role_type ?? "").Trim();
            string json_data = (param.json_data ?? "").Trim();

            SetDataMasterModel param_def = new SetDataMasterModel
            {
                user_name = user_name,
                role_type = role_type,
                json_data = json_data
            };

            ConvertJSONresultToDataSet(ref msg, ref ret, ref dsData, param_def);

            if (!(ret.ToLower() == "error"))
            {
                iMaxSeq = get_max("epha_m_mandatory_note");
                ret = set_mandatory_note(ref dsData, ref iMaxSeq);
            }

            return cls_json.SetJSONresult(refMsgSaveMaster(ret, msg, iMaxSeq.ToString()));
        }

        public string set_mandatory_note(ref DataSet dsData, ref int seq_now)
        {
            string ret = "";

            #region update data
            dt = new DataTable();
            dt = ConvertDStoDT(dsData, "data");

            try
            {
                using (TransactionScope scope = new TransactionScope())
                using (SqlConnection conn = new SqlConnection(ClassConnectionDb.ConnectionString()))
                {
                    conn.Open();

                    for (int i = 0; i < dt?.Rows.Count; i++)
                    {
                        string action_type = (dt.Rows[i]["action_type"]?.ToString() ?? "").ToLower();
                        string sqlstr = "";
                        List<SqlParameter> parameters = new List<SqlParameter>();

                        if (action_type == "insert")
                        {
                            #region insert
                            sqlstr = "insert into EPHA_M_MANDATORY_NOTE " +
                                     "(SEQ, ID, NAME, DESCRIPTION, ACTIVE_DEF, ACTIVE_TYPE, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                     "values (@SEQ, @ID, @NAME, @DESCRIPTION, @ACTIVE_DEF, @ACTIVE_TYPE, getdate(), null, @CREATE_BY, @UPDATE_BY)";

                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                            parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["NAME"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@DESCRIPTION", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTION"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@ACTIVE_DEF", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_DEF"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["CREATE_BY"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"]?.ToString() ?? "" });

                            seq_now += 1;
                            #endregion
                        }
                        else if (action_type == "update")
                        {
                            #region update
                            sqlstr = "update EPHA_M_MANDATORY_NOTE set " +
                                     "NAME = @NAME, DESCRIPTION = @DESCRIPTION, ACTIVE_DEF = @ACTIVE_DEF, ACTIVE_TYPE = @ACTIVE_TYPE, UPDATE_DATE = getdate(), UPDATE_BY = @UPDATE_BY " +
                                     "where SEQ = @SEQ and ID = @ID";

                            parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["NAME"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@DESCRIPTION", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTION"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@ACTIVE_DEF", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_DEF"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"]?.ToString() ?? "" });
                            #endregion
                        }
                        else if (action_type == "delete")
                        {
                            #region delete
                            sqlstr = "delete from EPHA_M_MANDATORY_NOTE where SEQ = @SEQ and ID = @ID";

                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"]?.ToString() ?? "" });
                            #endregion
                        }

                        if (action_type != "")
                        {
                            ClassConnectionDb cls_conn = new ClassConnectionDb();
                            ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, null); // transaction is null because it's handled by TransactionScope
                            if (ret != "true") { throw new Exception("Database operation failed"); }
                        }
                    }

                    if (ret == "true")
                    {
                        scope.Complete();
                    }
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            #endregion update data

            return ret;
        }


        #endregion JSEA Module : Mandatory Note

        #region JSEA Module : Task Type
        public string get_master_tasktype(LoadMasterPageModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = (param.user_name + "").Trim();

            int iMaxSeq = get_max("epha_m_request_type");

            sqlstr = @" select a.*, 'update' as action_type, 0 as action_change
                        from epha_m_request_type a 
                        order by seq ";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            if (dt?.Rows.Count == 0)
            {
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = iMaxSeq;
                dt.Rows[0]["id"] = iMaxSeq;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();

                iMaxSeq += 1;
            }
            dt.AcceptChanges();

            dt.TableName = "data";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "seq", (iMaxSeq + 1).ToString());
            dtma.TableName = "max";

            dsData.Tables.Add(dtma.Copy()); dsData.AcceptChanges();

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;

        }
        public string set_master_tasktype(SetDataMasterModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();

            int iMaxSeq = 0;
            string msg = "";
            string ret = "";
            string user_name = param.user_name?.Trim() ?? "";
            string role_type = param.role_type?.Trim() ?? "";
            string json_data = param.json_data?.Trim() ?? "";

            SetDataMasterModel param_def = new SetDataMasterModel
            {
                user_name = user_name,
                role_type = role_type,
                json_data = json_data
            };

            ConvertJSONresultToDataSet(ref msg, ref ret, ref dsData, param_def);

            if (!(ret.ToLower() == "error"))
            {
                iMaxSeq = get_max("epha_m_request_type");
                ret = set_request_type(ref dsData, ref iMaxSeq);
            }
            return cls_json.SetJSONresult(refMsgSaveMaster(ret, msg, iMaxSeq.ToString()));
        }

        public string set_request_type(ref DataSet dsData, ref int seq_now)
        {
            string ret = "";

            #region update data
            dt = new DataTable();
            dt = ConvertDStoDT(dsData, "data");

            try
            {
                using (TransactionScope scope = new TransactionScope())
                using (SqlConnection conn = new SqlConnection(ClassConnectionDb.ConnectionString()))
                {
                    conn.Open();

                    for (int i = 0; i < dt?.Rows.Count; i++)
                    {
                        string action_type = (dt.Rows[i]["action_type"]?.ToString() ?? "").ToLower();
                        string sqlstr = "";
                        List<SqlParameter> parameters = new List<SqlParameter>();

                        if (action_type == "insert")
                        {
                            #region insert
                            sqlstr = "insert into EPHA_M_REQUEST_TYPE " +
                                     "(SEQ, ID, NAME, DESCRIPTION, PHA_SUB_SOFTWARE, ACTIVE_TYPE, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                     "values (@SEQ, @ID, @NAME, @DESCRIPTION, @PHA_SUB_SOFTWARE, @ACTIVE_TYPE, getdate(), null, @CREATE_BY, @UPDATE_BY)";

                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                            parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["NAME"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@DESCRIPTION", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTION"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@PHA_SUB_SOFTWARE", SqlDbType.NVarChar, 50) { Value = "JSEA" });
                            parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["CREATE_BY"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"]?.ToString() ?? "" });

                            seq_now += 1;
                            #endregion
                        }
                        else if (action_type == "update")
                        {
                            #region update
                            sqlstr = "update EPHA_M_REQUEST_TYPE set " +
                                     "NAME = @NAME, DESCRIPTION = @DESCRIPTION, ACTIVE_TYPE = @ACTIVE_TYPE, UPDATE_DATE = getdate(), UPDATE_BY = @UPDATE_BY " +
                                     "where SEQ = @SEQ and ID = @ID";

                            parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["NAME"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@DESCRIPTION", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTION"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"]?.ToString() ?? "" });
                            #endregion
                        }
                        else if (action_type == "delete")
                        {
                            #region delete
                            sqlstr = "delete from EPHA_M_REQUEST_TYPE where SEQ = @SEQ and ID = @ID";

                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"]?.ToString() ?? "" });
                            #endregion
                        }

                        if (action_type != "")
                        {
                            ClassConnectionDb cls_conn = new ClassConnectionDb();
                            ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, null); // transaction is null because it's handled by TransactionScope
                            if (ret != "true") { throw new Exception("Database operation failed"); }
                        }
                    }

                    if (ret == "true")
                    {
                        scope.Complete();
                    }
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            #endregion update data

            return ret;
        }


        #endregion JSEA Module : Task Type

        #region JSEA Module : Tag ID/Equipment
        public string get_master_tagid(LoadMasterPageModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = (param.user_name + "").Trim();

            int iMaxSeq = get_max("epha_m_tagid");

            sqlstr = @" select a.*, 'update' as action_type, 0 as action_change
                        from epha_m_tagid a 
                        order by seq ";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            if (dt?.Rows.Count == 0)
            {
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = iMaxSeq;
                dt.Rows[0]["id"] = iMaxSeq;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();

                iMaxSeq += 1;
            }
            dt.AcceptChanges();

            dt.TableName = "data";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "seq", (iMaxSeq + 1).ToString());
            dtma.TableName = "max";

            dsData.Tables.Add(dtma.Copy()); dsData.AcceptChanges();

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;

        }
        public string set_master_tagid(SetDataMasterModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();

            int iMaxSeq = 0;
            string msg = "";
            string ret = "";
            string user_name = param.user_name?.Trim() ?? "";
            string role_type = param.role_type?.Trim() ?? "";
            string json_data = param.json_data?.Trim() ?? "";

            SetDataMasterModel param_def = new SetDataMasterModel();
            param_def.user_name = user_name;
            param_def.role_type = role_type;
            param_def.json_data = json_data;

            ConvertJSONresultToDataSet(ref msg, ref ret, ref dsData, param_def);
            if (!(ret.ToLower() == "error"))
            {
                iMaxSeq = get_max("epha_m_tagid");
                ret = set_tagid(ref dsData, ref iMaxSeq);
            }
            return cls_json.SetJSONresult(refMsgSaveMaster(ret, msg, iMaxSeq.ToString()));
        }

        public string set_tagid(ref DataSet dsData, ref int seq_now)
        {
            string ret = "";

            #region update data
            dt = new DataTable();
            dt = ConvertDStoDT(dsData, "data");

            try
            {
                using (TransactionScope scope = new TransactionScope())
                using (SqlConnection conn = new SqlConnection(ClassConnectionDb.ConnectionString()))
                {
                    conn.Open();

                    for (int i = 0; i < dt?.Rows.Count; i++)
                    {
                        string action_type = (dt.Rows[i]["action_type"]?.ToString() ?? "").ToLower();
                        string sqlstr = "";
                        List<SqlParameter> parameters = new List<SqlParameter>();

                        if (action_type == "insert")
                        {
                            #region insert
                            sqlstr = "insert into EPHA_M_TAGID " +
                                     "(SEQ, ID, NAME, DESCRIPTIONS, ACTIVE_TYPE, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                     "values (@SEQ, @ID, @NAME, @DESCRIPTIONS, @ACTIVE_TYPE, getdate(), null, @CREATE_BY, @UPDATE_BY)";

                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                            parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["NAME"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTION"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["CREATE_BY"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"]?.ToString() ?? "" });

                            seq_now += 1;
                            #endregion
                        }
                        else if (action_type == "update")
                        {
                            #region update
                            sqlstr = "update EPHA_M_TAGID set " +
                                     "NAME = @NAME, DESCRIPTION = @DESCRIPTIONS, ACTIVE_TYPE = @ACTIVE_TYPE, " +
                                     "UPDATE_DATE = getdate(), UPDATE_BY = @UPDATE_BY " +
                                     "where SEQ = @SEQ and ID = @ID";

                            parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["NAME"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTION"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"]?.ToString() ?? "" });
                            #endregion
                        }
                        else if (action_type == "delete")
                        {
                            #region delete
                            sqlstr = "delete from EPHA_M_TAGID where SEQ = @SEQ and ID = @ID";

                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"]?.ToString() ?? "" });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"]?.ToString() ?? "" });
                            #endregion
                        }

                        if (action_type != "")
                        {
                            ClassConnectionDb cls_conn = new ClassConnectionDb();
                            ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, null); // transaction is null because it's handled by TransactionScope
                            if (ret != "true") { throw new Exception("Database operation failed"); }
                        }
                    }

                    if (ret == "true")
                    {
                        scope.Complete();
                    }
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            #endregion update data

            return ret;
        }

        #endregion JSEA Module : Tag ID/Equipment

        #region HRA : Sections Group => Sub Area Group
        public string get_master_sections_group(LoadMasterPageModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = (param.user_name + "").Trim();

            int iMaxSeq = get_max("epha_m_sections_group");

            sqlstr = @" select a.*, 'update' as action_type, 0 as action_change, case when a.seq = 0 then 1 else 0 end disable_page
                        from epha_m_sections_group a 
                        order by seq ";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            if (dt?.Rows.Count == 0)
            {
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = iMaxSeq;
                dt.Rows[0]["id"] = iMaxSeq;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();

                iMaxSeq += 1;
            }
            dt.AcceptChanges();

            dt.TableName = "data";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "seq", (iMaxSeq + 1).ToString());
            dtma.TableName = "max";

            dsData.Tables.Add(dtma.Copy()); dsData.AcceptChanges();

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;

        }
        public string set_master_sections_group(SetMasterGuideWordsModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();

            int iMaxSeq = 0;
            string msg = "";
            string ret = "";
            string user_name = (param.user_name + "").Trim();
            string json_data = (param.json_data + "").Trim();

            SetDataMasterModel param_def = new SetDataMasterModel();
            param_def.user_name = param.user_name;
            param_def.json_data = param.json_data;

            ConvertJSONresultToDataSet(ref msg, ref ret, ref dsData, param_def);
            if (!(ret.ToLower() == "error"))
            {
                iMaxSeq = get_max("epha_m_sections_group");
                ret = set_sections_group(ref dsData, ref iMaxSeq);
            }
            return cls_json.SetJSONresult(refMsgSaveMaster(ret, msg, (iMaxSeq.ToString() + "").ToString()));
        }

        public string set_sections_group(ref DataSet dsData, ref int seq_now)
        {
            string ret = "";

            #region update data
            dt = new DataTable();
            dt = ConvertDStoDT(dsData, "data");

            try
            {
                using (TransactionScope scope = new TransactionScope())
                using (SqlConnection conn = new SqlConnection(ClassConnectionDb.ConnectionString()))
                {
                    conn.Open();

                    for (int i = 0; i < dt?.Rows.Count; i++)
                    {
                        string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                        string sqlstr = "";
                        List<SqlParameter> parameters = new List<SqlParameter>();

                        if (action_type == "insert")
                        {
                            #region insert
                            sqlstr = "insert into EPHA_M_SECTIONS_GROUP " +
                                     "(SEQ, ID, NAME, DESCRIPTIONS, ACTIVE_TYPE, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                     "values (@SEQ, @ID, @NAME, @DESCRIPTIONS, @ACTIVE_TYPE, getdate(), null, @CREATE_BY, @UPDATE_BY)";

                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                            parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["NAME"] });
                            parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTIONS"] });
                            parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"] });
                            parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["CREATE_BY"] });
                            parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"] });

                            seq_now += 1;
                            #endregion
                        }
                        else if (action_type == "update")
                        {
                            #region update
                            sqlstr = "update EPHA_M_SECTIONS_GROUP set " +
                                     "NAME = @NAME, DESCRIPTIONS = @DESCRIPTIONS, ACTIVE_TYPE = @ACTIVE_TYPE, " +
                                     "UPDATE_DATE = getdate(), UPDATE_BY = @UPDATE_BY " +
                                     "where SEQ = @SEQ and ID = @ID";

                            parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["NAME"] });
                            parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTIONS"] });
                            parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"] });
                            parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"] });
                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"] });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"] });
                            #endregion
                        }
                        else if (action_type == "delete")
                        {
                            #region delete
                            sqlstr = "delete from EPHA_M_SECTIONS_GROUP where SEQ = @SEQ and ID = @ID";

                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"] });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"] });
                            #endregion
                        }

                        if (action_type != "")
                        {
                            ClassConnectionDb cls_conn = new ClassConnectionDb();
                            ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, null); // transaction is null because it's handled by TransactionScope
                            if (ret != "true") { throw new Exception("Database operation failed"); }
                        }
                    }

                    if (ret == "true")
                    {
                        scope.Complete();
                    }
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            #endregion update data

            return ret;
        }

        #endregion HRA : Sections Group => Sub Area Group

        #region HRA : Group of Sub Area
        public string get_master_sub_area_group(LoadMasterPageModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = (param.user_name + "").Trim();

            #region Departments
            sqlstr = @"select distinct emp.departments as id,emp.departments as name
                         ,emp.functions, emp.departments
                         from vw_epha_person_details emp
                         where emp.departments is not null
                         order by emp.functions, emp.departments";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            dt.TableName = "departments";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion Departments

            #region Sections
            sqlstr = @"select distinct emp.sections as id, emp.sections as name,
                         emp.functions, emp.departments, emp.sections
                         from vw_epha_person_details emp
                         where emp.departments is not null and emp.sections is not null
                         order by emp.functions, emp.departments, emp.sections";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            dt.TableName = "sections";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion Sections


            #region Group of Sub Area 
            sqlstr = @" select  a.id, a.name, lower(a.name) as  field_check  
                        from epha_m_sections_group  a
                        where a.active_type = 1
                        order by a.name";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            dt.TableName = "sections_group";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion Group of Sub Area


            int iMaxSeq = get_max("epha_m_hazard_type");

            sqlstr = @" select a.*, 'update' as action_type, 0 as action_change, case when a.seq = 0 then 1 else 0 end disable_page
                        from epha_m_sub_area a 
                        order by id_sections, id_sections_group, seq";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            if (dt?.Rows.Count == 0)
            {
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = iMaxSeq;
                dt.Rows[0]["id"] = iMaxSeq;
                dt.Rows[0]["id_sections_group"] = null;
                dt.Rows[0]["id_sections"] = null;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();

                iMaxSeq += 1;
            }
            dt.AcceptChanges();

            dt.TableName = "data";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "seq", (iMaxSeq + 1).ToString());
            dtma.TableName = "max";

            dsData.Tables.Add(dtma.Copy()); dsData.AcceptChanges();

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;

        }
        public string set_master_sub_area_group(SetMasterGuideWordsModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();

            int iMaxSeq = 0;
            string msg = "";
            string ret = "";
            string user_name = (param.user_name + "").Trim();
            string json_data = (param.json_data + "").Trim();

            SetDataMasterModel param_def = new SetDataMasterModel();
            param_def.user_name = param.user_name;
            param_def.json_data = param.json_data;

            ConvertJSONresultToDataSet(ref msg, ref ret, ref dsData, param_def);
            if (!(ret.ToLower() == "error"))
            {
                iMaxSeq = get_max("epha_m_sub_area");
                ret = set_sub_area_group(ref dsData, ref iMaxSeq);
            }

            return cls_json.SetJSONresult(refMsgSaveMaster(ret, msg, (iMaxSeq.ToString() + "").ToString()));
        }

        public string set_sub_area_group(ref DataSet dsData, ref int seq_now)
        {
            string ret = "";

            #region update data
            dt = new DataTable();
            dt = ConvertDStoDT(dsData, "data");

            try
            {
                using (TransactionScope scope = new TransactionScope())
                using (SqlConnection conn = new SqlConnection(ClassConnectionDb.ConnectionString()))
                {
                    conn.Open();

                    for (int i = 0; i < dt?.Rows.Count; i++)
                    {
                        string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                        string sqlstr = "";
                        List<SqlParameter> parameters = new List<SqlParameter>();

                        if (action_type == "insert")
                        {
                            #region insert
                            sqlstr = "insert into EPHA_M_SUB_AREA " +
                                     "(ID_SECTIONS, ID_SECTIONS_GROUP, SEQ, ID, NAME, DESCRIPTIONS, ACTIVE_TYPE, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                     "values (@ID_SECTIONS, @ID_SECTIONS_GROUP, @SEQ, @ID, @NAME, @DESCRIPTIONS, @ACTIVE_TYPE, getdate(), null, @CREATE_BY, @UPDATE_BY)";

                            parameters.Add(new SqlParameter("@ID_SECTIONS", SqlDbType.Int) { Value = dt.Rows[i]["ID_SECTIONS"] });
                            parameters.Add(new SqlParameter("@ID_SECTIONS_GROUP", SqlDbType.Int) { Value = dt.Rows[i]["ID_SECTIONS_GROUP"] });
                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                            parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["NAME"] });
                            parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTIONS"] });
                            parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"] });
                            parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["CREATE_BY"] });
                            parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"] });

                            seq_now += 1;
                            #endregion
                        }
                        else if (action_type == "update")
                        {
                            #region update
                            sqlstr = "update EPHA_M_SUB_AREA set " +
                                     "NAME = @NAME, DESCRIPTIONS = @DESCRIPTIONS, ACTIVE_TYPE = @ACTIVE_TYPE, " +
                                     "UPDATE_DATE = getdate(), UPDATE_BY = @UPDATE_BY " +
                                     "where SEQ = @SEQ and ID = @ID";

                            parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["NAME"] });
                            parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTIONS"] });
                            parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"] });
                            parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"] });
                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"] });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"] });
                            #endregion
                        }
                        else if (action_type == "delete")
                        {
                            #region delete
                            sqlstr = "delete from EPHA_M_SUB_AREA where SEQ = @SEQ and ID = @ID";

                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"] });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"] });
                            #endregion
                        }

                        if (action_type != "")
                        {
                            ClassConnectionDb cls_conn = new ClassConnectionDb();
                            ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, null); // transaction is null because it's handled by TransactionScope
                            if (ret != "true") { throw new Exception("Database operation failed"); }
                        }
                    }

                    if (ret == "true")
                    {
                        scope.Complete();
                    }
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            #endregion update data

            return ret;
        }

        #endregion HRA : Group of Sub Area

        #region HRA : Equipmet of Sub Area 
        public void get_master_data_of_area(ref DataSet dsData)
        {
            #region Plant
            sqlstr = @"  select distinct a.id_plant as id, a.plant as name, a.plant_check from vw_epha_data_of_area a order by a.plant ";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            dt.TableName = "plant";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion Plant

            #region Area Process Unit
            sqlstr = @"  select distinct a.id_area as id, a.area as name, a.plant_check, a.area_check from vw_epha_data_of_area a order by a.area ";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            dt.TableName = "area";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion Area Process Unit

            #region Complex

            sqlstr = @" select distinct a.id_toc as id, a.toc as name, a.plant_check, a.area_check, a.toc_check from vw_epha_data_of_area a order by a.toc ";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            dt.TableName = "toc";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion Complex

            #region master apu

            sqlstr = @"  select distinct a.id_unit as id, a.unit +'-'+ a.toc as name, a.plant_check, a.area_check, a.toc_check from vw_epha_data_of_area a order by a.unit +'-'+ a.toc  ";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            dt.TableName = "apu";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion master apu

        }
        public string get_master_sub_area_equipmet(LoadMasterPageModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = (param.user_name + "").Trim();

            get_master_data_of_area(ref dsData);

            #region Group of Sub Area 
            sqlstr = @" select  a.id, a.name, lower(a.name) as  field_check  
                        from epha_m_sections_group  a
                        where a.active_type = 1
                        order by a.name";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            dt.TableName = "sections_group";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion Group of Sub Area


            int iMaxSeq = get_max("epha_m_hazard_type");

            sqlstr = @" select a.*, 'update' as action_type, 0 as action_change, case when a.seq = 0 then 1 else 0 end disable_page
                        from epha_m_sub_area a 
                        order by id_sections, id_sections_group, seq";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            if (dt?.Rows.Count == 0)
            {
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = iMaxSeq;
                dt.Rows[0]["id"] = iMaxSeq;
                dt.Rows[0]["id_sections_group"] = null;
                dt.Rows[0]["id_sections"] = null;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();

                iMaxSeq += 1;
            }
            dt.AcceptChanges();

            dt.TableName = "data";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "seq", (iMaxSeq + 1).ToString());
            dtma.TableName = "max";

            dsData.Tables.Add(dtma.Copy()); dsData.AcceptChanges();

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;

        }
        public string set_master_sub_area_equipmet(SetMasterGuideWordsModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();

            int iMaxSeq = 0;
            string msg = "";
            string ret = "";
            string user_name = (param.user_name + "").Trim();
            string json_data = (param.json_data + "").Trim();

            SetDataMasterModel param_def = new SetDataMasterModel();
            param_def.user_name = param.user_name;
            param_def.json_data = param.json_data;

            ConvertJSONresultToDataSet(ref msg, ref ret, ref dsData, param_def);
            if (!(ret.ToLower() == "error"))
            {
                iMaxSeq = get_max("epha_m_sub_area");
                ret = set_sub_area_equipmet(ref dsData, ref iMaxSeq);
            }

            return cls_json.SetJSONresult(refMsgSaveMaster(ret, msg, (iMaxSeq.ToString() + "").ToString()));
        }

        public string set_sub_area_equipmet(ref DataSet dsData, ref int seq_now)
        {
            string ret = "";

            #region update data
            dt = new DataTable();
            dt = ConvertDStoDT(dsData, "data");

            try
            {
                using (TransactionScope scope = new TransactionScope())
                using (SqlConnection conn = new SqlConnection(ClassConnectionDb.ConnectionString()))
                {
                    conn.Open();

                    for (int i = 0; i < dt?.Rows.Count; i++)
                    {
                        string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                        string sqlstr = "";
                        List<SqlParameter> parameters = new List<SqlParameter>();

                        if (action_type == "insert")
                        {
                            #region insert
                            sqlstr = "insert into EPHA_M_SUB_AREA " +
                                     "(ID_SECTIONS, ID_SECTIONS_GROUP, SEQ, ID, NAME, DESCRIPTIONS, ACTIVE_TYPE, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                     "values (@ID_SECTIONS, @ID_SECTIONS_GROUP, @SEQ, @ID, @NAME, @DESCRIPTIONS, @ACTIVE_TYPE, getdate(), null, @CREATE_BY, @UPDATE_BY)";

                            parameters.Add(new SqlParameter("@ID_SECTIONS", SqlDbType.Int) { Value = dt.Rows[i]["ID_SECTIONS"] });
                            parameters.Add(new SqlParameter("@ID_SECTIONS_GROUP", SqlDbType.Int) { Value = dt.Rows[i]["ID_SECTIONS_GROUP"] });
                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                            parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["NAME"] });
                            parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTIONS"] });
                            parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"] });
                            parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["CREATE_BY"] });
                            parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"] });

                            seq_now += 1;
                            #endregion
                        }
                        else if (action_type == "update")
                        {
                            #region update
                            sqlstr = "update EPHA_M_SUB_AREA set " +
                                     "NAME = @NAME, DESCRIPTIONS = @DESCRIPTIONS, ACTIVE_TYPE = @ACTIVE_TYPE, " +
                                     "UPDATE_DATE = getdate(), UPDATE_BY = @UPDATE_BY " +
                                     "where SEQ = @SEQ and ID = @ID";

                            parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["NAME"] });
                            parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = dt.Rows[i]["DESCRIPTIONS"] });
                            parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = dt.Rows[i]["ACTIVE_TYPE"] });
                            parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = dt.Rows[i]["UPDATE_BY"] });
                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"] });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"] });
                            #endregion
                        }
                        else if (action_type == "delete")
                        {
                            #region delete
                            sqlstr = "delete from EPHA_M_SUB_AREA where SEQ = @SEQ and ID = @ID";

                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[i]["SEQ"] });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[i]["ID"] });
                            #endregion
                        }

                        if (action_type != "")
                        {
                            ClassConnectionDb cls_conn = new ClassConnectionDb();
                            ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, null); // transaction is null because it's handled by TransactionScope
                            if (ret != "true") { throw new Exception("Database operation failed"); }
                        }
                    }

                    if (ret == "true")
                    {
                        scope.Complete();
                    }
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            #endregion update data

            return ret;
        }

        #endregion HRA : Equipmet of Sub Area

        #region HRA : Hazard Type & Hazard Riskfactors
        public string get_master_hazard_type(LoadMasterPageModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = (param.user_name + "").Trim();

            int iMaxSeq = get_max("epha_m_hazard_type");

            sqlstr = @" select a.*, 'update' as action_type, 0 as action_change, case when a.seq = 0 then 1 else 0 end disable_page
                        from epha_m_hazard_type a 
                        order by seq ";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            if (dt?.Rows.Count == 0)
            {
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = iMaxSeq;
                dt.Rows[0]["id"] = iMaxSeq;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();

                iMaxSeq += 1;
            }
            dt.AcceptChanges();

            dt.TableName = "data";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "seq", (iMaxSeq + 1).ToString());
            dtma.TableName = "max";

            dsData.Tables.Add(dtma.Copy()); dsData.AcceptChanges();

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;

        }
        public string set_master_hazard_type(SetMasterGuideWordsModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();

            int iMaxSeq = 0;
            string msg = "";
            string ret = "";
            string user_name = (param.user_name + "").Trim();
            string json_data = (param.json_data + "").Trim();

            SetDataMasterModel param_def = new SetDataMasterModel();
            param_def.user_name = param.user_name;
            param_def.json_data = param.json_data;

            ConvertJSONresultToDataSet(ref msg, ref ret, ref dsData, param_def);
            if (!(ret.ToLower() == "error"))
            {
                iMaxSeq = get_max("epha_m_hazard_type");

                try
                {
                    using (TransactionScope scope = new TransactionScope())
                    {
                        ret = set_hazard_type(ref dsData, ref iMaxSeq);
                        if (ret != "true") throw new Exception(ret);

                        scope.Complete();
                    }
                }
                catch (Exception ex)
                {
                    ret = "error: " + ex.Message;
                }
            }

            return cls_json.SetJSONresult(refMsgSaveMaster(ret, msg, iMaxSeq.ToString()));
        }

        public string set_hazard_type(ref DataSet dsData, ref int seq_now)
        {
            string ret = "";
            dt = new DataTable();
            dt = ConvertDStoDT(dsData, "data");

            using (SqlConnection conn = new SqlConnection(ClassConnectionDb.ConnectionString()))
            {
                conn.Open();
                using (SqlTransaction transaction = conn.BeginTransaction())
                {
                    try
                    {
                        foreach (DataRow row in dt.Rows)
                        {
                            string action_type = row["action_type"].ToString();
                            List<SqlParameter> parameters = new List<SqlParameter>();
                            string sqlstr = "";

                            if (action_type == "insert")
                            {
                                sqlstr = "insert into EPHA_M_HAZARD_TYPE " +
                                         "(SEQ, ID, NAME, DESCRIPTIONS, ACTIVE_TYPE, CREATE_DATE, CREATE_BY, UPDATE_BY) " +
                                         "values (@SEQ, @ID, @NAME, @DESCRIPTIONS, @ACTIVE_TYPE, getdate(), @CREATE_BY, @UPDATE_BY)";

                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                                parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = row["NAME"] });
                                parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = row["DESCRIPTIONS"] });
                                parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = row["ACTIVE_TYPE"] });
                                parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = row["CREATE_BY"] });
                                parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] });

                                seq_now++;
                            }
                            else if (action_type == "update")
                            {
                                sqlstr = "update EPHA_M_HAZARD_TYPE set " +
                                         "NAME = @NAME, DESCRIPTIONS = @DESCRIPTIONS, ACTIVE_TYPE = @ACTIVE_TYPE, " +
                                         "UPDATE_DATE = getdate(), UPDATE_BY = @UPDATE_BY " +
                                         "where SEQ = @SEQ and ID = @ID";

                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] });
                                parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = row["NAME"] });
                                parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = row["DESCRIPTIONS"] });
                                parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = row["ACTIVE_TYPE"] });
                                parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] });
                            }
                            else if (action_type == "delete")
                            {
                                sqlstr = "delete from EPHA_M_HAZARD_TYPE where SEQ = @SEQ and ID = @ID";

                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] });
                            }

                            if (!string.IsNullOrEmpty(sqlstr))
                            {
                                ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                                if (ret != "true") throw new Exception(ret);
                            }
                        }

                        transaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        ret = "error: " + ex.Message;
                    }
                }
            }

            return ret;
        }


        public string get_master_hazard_riskfactors(LoadMasterPageModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = (param.user_name + "").Trim();

            int iMaxSeq = get_max("epha_m_hazard_riskfactors");

            sqlstr = @" select p.id, p.name 
                        from epha_m_hazard_type p 
                        where p.active_type = 1
                        order by p.id";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());
            dt.TableName = "hazard_type";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            sqlstr = @" select a.*, 'update' as action_type, 0 as action_change, case when a.seq = 0 then 1 else 0 end disable_page
                        from epha_m_hazard_riskfactors a 
                        order by seq ";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, new List<SqlParameter>());

            if (dt?.Rows.Count == 0)
            {
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = iMaxSeq;
                dt.Rows[0]["id"] = iMaxSeq;

                dt.Rows[0]["id_hazard_type"] = null;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();

                iMaxSeq += 1;
            }
            dt.AcceptChanges();

            dt.TableName = "data";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "seq", (iMaxSeq + 1).ToString());
            dtma.TableName = "max";

            dsData.Tables.Add(dtma.Copy()); dsData.AcceptChanges();

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;

        }
        public string set_master_hazard_riskfactors(SetMasterGuideWordsModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();

            int iMaxSeq = 0;
            string msg = "";
            string ret = "";
            string user_name = (param.user_name + "").Trim();
            string json_data = (param.json_data + "").Trim();

            SetDataMasterModel param_def = new SetDataMasterModel();
            param_def.user_name = param.user_name;
            param_def.json_data = param.json_data;

            ConvertJSONresultToDataSet(ref msg, ref ret, ref dsData, param_def);
            if (ret.ToLower() == "error") return cls_json.SetJSONresult(refMsgSaveMaster(ret, msg, ""));

            iMaxSeq = get_max("epha_m_hazard_riskfactors");

            try
            {
                using (TransactionScope scope = new TransactionScope())
                {
                    ret = set_hazard_riskfactors(ref dsData, ref iMaxSeq);
                    if (ret != "true") throw new Exception(ret);

                    scope.Complete();
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return cls_json.SetJSONresult(refMsgSaveMaster(ret, msg, iMaxSeq.ToString()));
        }

        public string set_hazard_riskfactors(ref DataSet dsData, ref int seq_now)
        {
            string ret = "";
            dt = new DataTable();
            dt = ConvertDStoDT(dsData, "data");

            using (SqlConnection conn = new SqlConnection(ClassConnectionDb.ConnectionString()))
            {
                conn.Open();
                using (SqlTransaction transaction = conn.BeginTransaction())
                {
                    try
                    {
                        foreach (DataRow row in dt.Rows)
                        {
                            string action_type = row["action_type"].ToString();
                            List<SqlParameter> parameters = new List<SqlParameter>();
                            string sqlstr = "";
                            if (!string.IsNullOrEmpty(action_type))
                            {
                                if (action_type == "insert")
                                {
                                    sqlstr = "insert into EPHA_M_HAZARD_RISKFACTORS " +
                                             "(SEQ, ID, ID_HAZARD_TYPE, HEALTH_HAZARDS, HAZARDS_RATING, STANDARD_TYPE_TEXT, STANDARD_VALUE, STANDARD_UNIT, STANDARD_DESC, ACTIVE_TYPE, " +
                                             "CREATE_DATE, CREATE_BY, UPDATE_BY) " +
                                             "values (@SEQ, @ID, @ID_HAZARD_TYPE, @HEALTH_HAZARDS, @HAZARDS_RATING, @STANDARD_TYPE_TEXT, @STANDARD_VALUE, @STANDARD_UNIT, @STANDARD_DESC, @ACTIVE_TYPE, " +
                                             "getdate(), @CREATE_BY, @UPDATE_BY)";

                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                                    parameters.Add(new SqlParameter("@ID_HAZARD_TYPE", SqlDbType.Int) { Value = row["ID_HAZARD_TYPE"] });
                                    parameters.Add(new SqlParameter("@HEALTH_HAZARDS", SqlDbType.NVarChar, 4000) { Value = row["HEALTH_HAZARDS"] });
                                    parameters.Add(new SqlParameter("@HAZARDS_RATING", SqlDbType.NVarChar, 4000) { Value = row["HAZARDS_RATING"] });
                                    parameters.Add(new SqlParameter("@STANDARD_TYPE_TEXT", SqlDbType.NVarChar, 4000) { Value = row["STANDARD_TYPE_TEXT"] });
                                    parameters.Add(new SqlParameter("@STANDARD_VALUE", SqlDbType.NVarChar, 4000) { Value = row["STANDARD_VALUE"] });
                                    parameters.Add(new SqlParameter("@STANDARD_UNIT", SqlDbType.NVarChar, 4000) { Value = row["STANDARD_UNIT"] });
                                    parameters.Add(new SqlParameter("@STANDARD_DESC", SqlDbType.NVarChar, 4000) { Value = row["STANDARD_DESC"] });
                                    parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = row["ACTIVE_TYPE"] });
                                    parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = row["CREATE_BY"] });
                                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] });

                                    seq_now++;
                                }
                                else if (action_type == "update")
                                {
                                    sqlstr = "update EPHA_M_HAZARD_RISKFACTORS set " +
                                             "ID_HAZARD_TYPE = @ID_HAZARD_TYPE, HEALTH_HAZARDS = @HEALTH_HAZARDS, HAZARDS_RATING = @HAZARDS_RATING, " +
                                             "STANDARD_TYPE_TEXT = @STANDARD_TYPE_TEXT, STANDARD_VALUE = @STANDARD_VALUE, STANDARD_UNIT = @STANDARD_UNIT, " +
                                             "STANDARD_DESC = @STANDARD_DESC, ACTIVE_TYPE = @ACTIVE_TYPE, UPDATE_DATE = getdate(), UPDATE_BY = @UPDATE_BY " +
                                             "where SEQ = @SEQ and ID = @ID";

                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] });
                                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] });
                                    parameters.Add(new SqlParameter("@ID_HAZARD_TYPE", SqlDbType.Int) { Value = row["ID_HAZARD_TYPE"] });
                                    parameters.Add(new SqlParameter("@HEALTH_HAZARDS", SqlDbType.NVarChar, 4000) { Value = row["HEALTH_HAZARDS"] });
                                    parameters.Add(new SqlParameter("@HAZARDS_RATING", SqlDbType.NVarChar, 4000) { Value = row["HAZARDS_RATING"] });
                                    parameters.Add(new SqlParameter("@STANDARD_TYPE_TEXT", SqlDbType.NVarChar, 4000) { Value = row["STANDARD_TYPE_TEXT"] });
                                    parameters.Add(new SqlParameter("@STANDARD_VALUE", SqlDbType.NVarChar, 4000) { Value = row["STANDARD_VALUE"] });
                                    parameters.Add(new SqlParameter("@STANDARD_UNIT", SqlDbType.NVarChar, 4000) { Value = row["STANDARD_UNIT"] });
                                    parameters.Add(new SqlParameter("@STANDARD_DESC", SqlDbType.NVarChar, 4000) { Value = row["STANDARD_DESC"] });
                                    parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = row["ACTIVE_TYPE"] });
                                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] });
                                }
                                else if (action_type == "delete")
                                {
                                    sqlstr = "delete from EPHA_M_HAZARD_RISKFACTORS where SEQ = @SEQ and ID = @ID";

                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] });
                                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] });
                                }

                                if (!string.IsNullOrEmpty(sqlstr))
                                {
                                    ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                                    if (ret != "true") throw new Exception(ret);
                                }
                            }
                        }

                        transaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        ret = "error: " + ex.Message;
                    }
                }
            }

            return ret;
        }


        #endregion HRA : Hazard Type & Hazard Riskfactors


        #region Department, Sections
        private void getMasterDepartmentSections(ref DataSet ds, string sections_name = "", string departments_name = "")
        {
            List<SqlParameter> parameters = new List<SqlParameter>();
            dt = new DataTable();

            #region Sections
            sqlstr = "usp_GetMasterSections";
            parameters = new List<SqlParameter>();
            if (sections_name != "") { parameters.Add(new SqlParameter("@sections_name", SqlDbType.VarChar) { Value = sections_name }); }
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters, "sections", true);
            if (dt?.Rows.Count > 0) { departments_name = dt.Rows[0]["departments"]?.ToString() ?? ""; }

            dt.TableName = "sections";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion Sections

            #region Departments 
            sqlstr = "usp_GetMasterDepartments";
            parameters = new List<SqlParameter>();
            if (departments_name != "") { parameters.Add(new SqlParameter("@departments_name", SqlDbType.VarChar) { Value = departments_name }); }
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters, "departments", true);

            dt.TableName = "departments";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion Departments

        }
        #endregion Department, Sections

        #region HRA : Group List
        public string get_master_group_list(LoadMasterPageBySectionModel param)
        {
            List<SqlParameter> parameters = new List<SqlParameter>();
            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = (param.user_name?.ToString() ?? "").Trim();

            int iMaxSeq = get_max("epha_m_group_list");

            sqlstr = "usp_get_epha_m_group_list";
            parameters = new List<SqlParameter>();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters, "data", true);

            if (dt?.Rows.Count == 0)
            {
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = iMaxSeq;
                dt.Rows[0]["id"] = iMaxSeq;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();

                iMaxSeq += 1;
            }
            dt.AcceptChanges();

            dt.TableName = "data";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "seq", (iMaxSeq + 1).ToString());
            dtma.TableName = "max";

            dsData.Tables.Add(dtma.Copy()); dsData.AcceptChanges();

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;

        }
        public string set_master_group_list(SetMasterGuideWordsModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();

            int iMaxSeq = 0;
            string msg = "";
            string ret = "";
            string user_name = (param.user_name + "").Trim();
            string json_data = (param.json_data + "").Trim();

            SetDataMasterModel param_def = new SetDataMasterModel();
            param_def.user_name = param.user_name;
            param_def.json_data = param.json_data;

            ConvertJSONresultToDataSet(ref msg, ref ret, ref dsData, param_def);
            if (ret.ToLower() == "error") { goto Next_Line_Convert; }
            iMaxSeq = get_max("epha_m_group_list");


            using (var conn = new SqlConnection(ClassConnectionDb.ConnectionString()))
            {
                conn.Open();
                using (var transaction = conn.BeginTransaction())
                {
                    try
                    {
                        // Data
                        set_group_list(ref dsData, conn, transaction, ref iMaxSeq);

                        transaction.Commit();
                        ret = "true";
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        ret = "error: " + ex.Message;
                    }
                }
            }

        Next_Line_Convert:;

            return cls_json.SetJSONresult(refMsgSaveMaster(ret, msg, (iMaxSeq.ToString() + "").ToString()));

        }
        public string set_group_list(ref DataSet dsData, SqlConnection conn, SqlTransaction trans, ref int seq_now)
        {
            string ret = "";
            List<SqlParameter> parameters = new List<SqlParameter>();
            DataTable dt = ConvertDStoDT(dsData, "data");

            foreach (DataRow row in dt.Rows)
            {
                string action_type = row["action_type"]?.ToString() ?? "";

                parameters = new List<SqlParameter>();

                if (action_type == "insert")
                {
                    sqlstr = "INSERT INTO EPHA_M_GROUP_LIST " +
                             "(SEQ, ID, NAME, DESCRIPTIONS, ACTIVE_TYPE, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                             "VALUES (@SEQ, @ID, @NAME, @DESCRIPTIONS, @ACTIVE_TYPE, GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                    parameters.Add(new SqlParameter("@NAME", SqlDbType.VarChar, 4000) { Value = row["NAME"] });
                    parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.VarChar, 4000) { Value = row["DESCRIPTIONS"] });
                    parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = row["ACTIVE_TYPE"] });
                    parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.VarChar, 50) { Value = row["CREATE_BY"] });
                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = row["UPDATE_BY"] });

                    seq_now += 1;
                }
                else if (action_type == "update")
                {
                    sqlstr = "UPDATE EPHA_M_GROUP_LIST SET " +
                             "NAME = @NAME, DESCRIPTIONS = @DESCRIPTIONS, ACTIVE_TYPE = @ACTIVE_TYPE " +
                             "WHERE SEQ = @SEQ AND ID = @ID";

                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] });
                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] });
                    parameters.Add(new SqlParameter("@NAME", SqlDbType.VarChar, 4000) { Value = row["NAME"] });
                    parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.VarChar, 4000) { Value = row["DESCRIPTIONS"] });
                    parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = row["ACTIVE_TYPE"] });
                }
                else if (action_type == "delete")
                {
                    sqlstr = "DELETE FROM EPHA_M_GROUP_LIST WHERE SEQ = @SEQ AND ID = @ID";

                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] });
                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] });
                }

                if (action_type != "")
                {
                    ClassConnectionDb cls_conn = new ClassConnectionDb();
                    ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, trans);
                    if (ret != "true") { break; }
                }
            }

            return ret;
        }

        #endregion HRA : Group List

        #region HRA : Worker Group
        public string get_master_worker_group(LoadMasterPageBySectionModel param)
        {
            List<SqlParameter> parameters = new List<SqlParameter>();
            dtma = new DataTable();
            dsData = new DataSet();
            string user_name = (param.user_name?.ToString() ?? "").Trim();
            string id_sections = (param.id_sections?.ToString() ?? "").Trim();
            string id_group_list = "";
            string key_group_list = "";

            getMasterDepartmentSections(ref dsData);
            if (dsData.Tables["sections"]?.Rows.Count > 0)
            {
                id_sections = dsData.Tables["sections"]?.Rows[0]?.ToString() ?? "";
            }

            sqlstr = "usp_GetMasterGroupList";
            parameters = new List<SqlParameter>();
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters, "group_list", true);
            if (dt?.Rows.Count > 0)
            {
                id_group_list = dt.Rows[0]["id"]?.ToString() ?? "";
                key_group_list = dt.Rows[0]["name"]?.ToString() ?? "";
            }
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            int iMaxSeq = get_max("epha_m_worker_group");

            sqlstr = "usp_get_epha_m_worker_group";
            parameters = new List<SqlParameter>();
            //parameters.Add(new SqlParameter("@sections_name", SqlDbType.VarChar) { Value = "CMCS" });
            dt = new DataTable();
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters, "data", true);

            if (dt?.Rows.Count == 0)
            {
                if (dsData.Tables["group_list"]?.Rows.Count > 0)
                {
                    id_group_list = dsData.Tables["group_list"]?.Rows[0]["id"]?.ToString() ?? "";
                    key_group_list = dsData.Tables["group_list"]?.Rows[0]["name"]?.ToString() ?? "";
                    for (int i = 0; i < dsData.Tables["group_list"]?.Rows.Count; i++)
                    {
                        //กรณีที่เป็นใบงานใหม่
                        dt.Rows.Add(dt.NewRow());
                        dt.Rows[i]["seq"] = iMaxSeq;
                        dt.Rows[i]["id"] = iMaxSeq;
                        dt.Rows[i]["id_sections_group"] = null;
                        dt.Rows[i]["id_sections"] = id_sections;
                        dt.Rows[i]["id_group_list"] = id_group_list;
                        dt.Rows[i]["key_group_list"] = key_group_list;

                        dt.Rows[i]["create_by"] = user_name;
                        dt.Rows[i]["action_type"] = "insert";
                        dt.Rows[i]["action_change"] = 0;
                        dt.AcceptChanges();
                        iMaxSeq += 1;
                    }
                }
                else
                {
                    //กรณีที่เป็นใบงานใหม่
                    dt.Rows.Add(dt.NewRow());
                    dt.Rows[0]["seq"] = iMaxSeq;
                    dt.Rows[0]["id"] = iMaxSeq;
                    dt.Rows[0]["id_sections_group"] = null;
                    dt.Rows[0]["id_sections"] = id_sections;
                    dt.Rows[0]["id_group_list"] = id_group_list;
                    dt.Rows[0]["key_group_list"] = key_group_list;

                    dt.Rows[0]["create_by"] = user_name;
                    dt.Rows[0]["action_type"] = "insert";
                    dt.Rows[0]["action_change"] = 0;
                    dt.AcceptChanges();
                    iMaxSeq += 1;
                }
            }
            dt.AcceptChanges();

            dt.TableName = "data";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            clshazop = new ClassHazop();
            clshazop.set_max_id(ref dtma, "seq", (iMaxSeq + 1).ToString());
            dtma.TableName = "max";

            dsData.Tables.Add(dtma.Copy()); dsData.AcceptChanges();

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;

        }
        public string set_master_worker_group(SetMasterGuideWordsModel param)
        {
            dtma = new DataTable();
            dsData = new DataSet();

            int iMaxSeq = 0;
            string msg = "";
            string ret = "";
            string user_name = (param.user_name + "").Trim();
            string json_data = (param.json_data + "").Trim();

            SetDataMasterModel param_def = new SetDataMasterModel();
            param_def.user_name = param.user_name;
            param_def.json_data = param.json_data;

            ConvertJSONresultToDataSet(ref msg, ref ret, ref dsData, param_def);
            if (ret.ToLower() == "error") { goto Next_Line_Convert; }
            iMaxSeq = get_max("epha_m_worker_group");


            using (var conn = new SqlConnection(ClassConnectionDb.ConnectionString()))
            {
                conn.Open();
                using (var transaction = conn.BeginTransaction())
                {
                    try
                    {
                        // Data
                        set_worker_group(ref dsData, conn, transaction, ref iMaxSeq);
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

        Next_Line_Convert:;

            return cls_json.SetJSONresult(refMsgSaveMaster(ret, msg, (iMaxSeq.ToString() + "").ToString()));

        }
        public string set_worker_group(ref DataSet dsData, SqlConnection conn, SqlTransaction trans, ref int seq_now)
        {
            string ret = "";
            List<SqlParameter> parameters = new List<SqlParameter>();
            DataTable dt = ConvertDStoDT(dsData, "data");

            foreach (DataRow row in dt.Rows)
            {
                string action_type = row["action_type"]?.ToString() ?? "";

                parameters = new List<SqlParameter>();

                if (action_type == "insert")
                {
                    sqlstr = "INSERT INTO EPHA_M_WORKER_GROUP " +
                             "(ID_SECTIONS, ID_GROUP_LIST, KEY_GROUP_LIST, SEQ, ID, NAME, DESCRIPTIONS, ACTIVE_TYPE, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                             "VALUES (@ID_SECTIONS, @ID_GROUP_LIST, @KEY_GROUP_LIST, @SEQ, @ID, @NAME, @DESCRIPTIONS, @ACTIVE_TYPE, GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                    parameters.Add(new SqlParameter("@ID_SECTIONS", SqlDbType.VarChar, 4000) { Value = row["ID_SECTIONS"] });
                    parameters.Add(new SqlParameter("@ID_GROUP_LIST", SqlDbType.Int) { Value = row["ID_GROUP_LIST"] });
                    parameters.Add(new SqlParameter("@KEY_GROUP_LIST", SqlDbType.Int) { Value = row["KEY_GROUP_LIST"] });
                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_now });
                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_now });
                    parameters.Add(new SqlParameter("@NAME", SqlDbType.VarChar, 4000) { Value = row["NAME"] });
                    parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.VarChar, 4000) { Value = row["DESCRIPTIONS"] });
                    parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = row["ACTIVE_TYPE"] });
                    parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.VarChar, 50) { Value = row["CREATE_BY"] });
                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = row["UPDATE_BY"] });

                    seq_now += 1;
                }
                else if (action_type == "update")
                {
                    sqlstr = "UPDATE EPHA_M_WORKER_GROUP SET " +
                             "NAME = @NAME, DESCRIPTIONS = @DESCRIPTIONS, ACTIVE_TYPE = @ACTIVE_TYPE " +
                             "WHERE SEQ = @SEQ AND ID = @ID";

                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] });
                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] });
                    parameters.Add(new SqlParameter("@NAME", SqlDbType.VarChar, 4000) { Value = row["NAME"] });
                    parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.VarChar, 4000) { Value = row["DESCRIPTIONS"] });
                    parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = row["ACTIVE_TYPE"] });
                }
                else if (action_type == "delete")
                {
                    sqlstr = "DELETE FROM EPHA_M_WORKER_GROUP WHERE SEQ = @SEQ AND ID = @ID";

                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] });
                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] });
                }

                if (action_type != "")
                {
                    ClassConnectionDb cls_conn = new ClassConnectionDb();
                    ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, trans);
                    if (ret != "true") { break; }
                }
            }

            return ret;
        }

        #endregion HRA : Worker Group

    }
}
