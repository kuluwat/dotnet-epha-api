

using dotnet_epha_api.Class;
using dotnet6_epha_api.Class;
using Model;
using Newtonsoft.Json;
using System.Data;
using System.Data.SqlClient;
using System.Transactions;

namespace Class
{
    public class ClassHazopSet
    {
        string sqlstr = "";
        string jsper = "";
        string ret = "true";
        ClassFunctions cls = new ClassFunctions();
        ClassJSON cls_json = new ClassJSON();

        ClassEmail clsmail = new ClassEmail();

        DataSet dsData = new DataSet();
        DataSet dt = new DataSet();
        DataSet dtcopy = new DataSet();
        DataSet dtcheck = new DataSet();

        string[] sMonth = ("JAN,FEB,MAR,APR,MAY,JUN,JUL,AUG,SEP,OCT,NOV,DEC").Split(',');
        private bool IsAuthorized(string userName)
        {
            // ตรวจสอบว่ามี user_name หรือไม่
            return !string.IsNullOrEmpty(userName);
        }
        private object ConvertToDBNull(object value)
        {
            return value == null || value == DBNull.Value ? DBNull.Value : value;
        }
        private object ConvertToIntOrDBNull(object value)
        {
            try
            {
                return int.TryParse(value.ToString(), out int result) ? (object)result : DBNull.Value;
            }
            catch { return DBNull.Value; }
        }

        //private object ConvertToDateTimeOrDBNull(object value)
        //{
        //    return DateTime.TryParse(value.ToString(), out DateTime result) ? (object)result : DBNull.Value;
        //}
        private object ConvertToDateTimeOrDBNull(object value)
        {
            // ตรวจสอบว่า value เป็น null หรือไม่ ถ้าใช่ให้คืนค่า DBNull.Value
            if (value == null)
            {
                return DBNull.Value;
            }

            string dateString = value.ToString().Trim().ToLower();

            // ตรวจสอบว่าค่าที่ส่งมาคือ getdate() หรือไม่
            if (dateString == "getdate()")
            {
                // คืนค่า DateTime ปัจจุบัน
                return DateTime.Now;
            }

            // กำหนดรูปแบบวันที่ที่ต้องการให้ตรงกับ YYYYMMDD และ YYYY-MM-DD
            string[] formats = { "yyyyMMdd", "yyyy-MM-dd" };

            // พยายามแปลง string ให้เป็น DateTime ตามรูปแบบที่กำหนด
            if (DateTime.TryParseExact(dateString, formats, null, System.Globalization.DateTimeStyles.None, out DateTime result))
            {
                return result;
            }
            else
            {
                // ถ้าแปลงไม่สำเร็จให้คืนค่า DBNull.Value
                return DBNull.Value;
            }
        }



        private void add_columns_in_table(string user_name, ref DataSet? dsData, string tableName, List<string> requiredColumns)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (string.IsNullOrEmpty(user_name)) { throw new UnauthorizedAccessException("User is not authorized to perform this action."); }

            // ตรวจสอบว่ามี DataTable ที่ชื่อ "tableName" ใน DataSet หรือไม่
            if (dsData.Tables.Contains(tableName))
            {
                DataTable table = dsData?.Tables[tableName] ?? new DataTable();

                // ตรวจสอบแต่ละคอลัมน์ในรายการ requiredColumns
                foreach (var column in requiredColumns)
                {
                    if (!table.Columns.Contains(column))
                    {
                        // ถ้าไม่มีคอลัมน์นี้ ให้เพิ่มเข้าไป
                        table.Columns.Add(column, typeof(string));
                    }
                }
            }
            else
            {
                // ถ้าไม่มี DataTable ที่ชื่อ "tableName"
                // สร้าง DataTable ใหม่และเพิ่มเข้าไปใน DataSet
                DataTable newTable = new DataTable(tableName);
                foreach (var column in requiredColumns)
                {
                    newTable.Columns.Add(column, typeof(string));
                }
                dsData.Tables.Add(newTable);
            }
        }
        #region function
        private static void findRevisionText(string pha_sub_software, string rev_no, ref string rev_text, ref string rev_desc
            , Boolean bReteMOC, Boolean bRetTA2, Boolean bTA2ApproveAll, Boolean bComplate, Boolean bTA2ApproveMOC = false)
        {
            string sub_software = (pha_sub_software == "whatif" ? "What's If" : pha_sub_software.ToUpper());
            int iRev = Convert.ToInt32(rev_no == "" || rev_no == null ? 0 : rev_no);

            string[] sAtoZ = ("A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,W,S,Y,Z").Split(',');
            rev_text = (sAtoZ[iRev]).ToString();

            if (bTA2ApproveMOC == true) { rev_desc = "Issued for Approval"; }
            else if (bReteMOC == true) { rev_desc = sub_software + " Approved"; }
            else if (bRetTA2 == true) { rev_desc = "Issued for Approval"; }
            else if (bTA2ApproveAll == true) { rev_desc = sub_software + " Approved"; }
            else if (bComplate == true) { rev_text = "0"; rev_desc = "Issued for Final"; }
            else { rev_desc = "Issued for Review"; }


            //OPEX
            //Step Running Revision Remark-ตัวอย่าง Remark - ตัวอย่าง2
            //Create ครั้งแรกเป็น Status Draft                          Revision = -0 -
            //กด Task Register                                      Revision += 1    1   A = Issued for Review
            //Genarate Full Report                                  Revision += 1    2   B = Issued for Review
            //กรณีที่ Submit หลังทำ Conduct(Export Report to eMOC)       
            //--> ต้องกด Submit TA2 Review and Approve for MOC ก่อน 
            //                                                      Revision += 1    3   C = Issued for Approval
            //--> เมื่อ eMOC แจ้งมาว่า TA2 Approve แล้วให้กดปุ่ม Complete for e-MOC  
            //                                                      Revision += 1    5   D = Issued for Approval
            // ??? ขาด step approver send back นะ
            //เมื่อ Review Fullowup และ Complate items All             Revision += 1    6   0 = Issued for Final


            //CAPEX
            //Create ครั้งแรกเป็น Status Draft                          Revision = -0 -
            //กด Task Register                                      Revision += 1    1   A = Issued for Review
            //Genarate Full Report                                  Revision += 1    2   B = Issued for Review
            //Submit to TA2(รอบแรก)                                 Revision += 1    3   C = Issued for Approval
            //TA2 Approve(คนสุดท้าย)                                  Revision += 1    4   D = HAZOP Approved
            //เมื่อ Review Fullowup และ Complate items All             Revision += 1    5   0 = Issued for Final


        }
        private void getJsontoDataSet(string jsper, ref DataSet _dsData, string tableName, string userName)
        {
            // ตรวจสอบสิทธิ์ในฟังก์ชันนี้อีกครั้งเพื่อความปลอดภัย
            if (string.IsNullOrEmpty(userName))
            {
                throw new UnauthorizedAccessException("User is not authorized to perform this action.");
            }

            if (jsper.Trim() != "")
            {
                try
                {
                    ClassJSON _cls_json = new ClassJSON();
                    DataTable _dt = new DataTable();

                    _dt = _cls_json.ConvertJSONresult(jsper);
                    if (_dt != null)
                    {
                        _dt.TableName = tableName;
                        _dsData.Tables.Add(_dt.Copy());
                        _dsData.AcceptChanges();
                    }
                }
                catch (Exception ex)
                {
                    // จัดการข้อผิดพลาดที่อาจเกิดขึ้นจากการแปลง JSON
                    throw new Exception("Error in converting JSON to DataTable: " + ex.Message);
                }
            }
        }

        public string importfile_data_jsea(uploadFile uploadFile, string folder)
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
                    if (file_sub_software != "") { folder = (file_sub_software ?? ""); }

                    if (string.IsNullOrEmpty(folder)) { msg_error = "Invalid folder."; }
                    else
                    {
                        msg_error = ClassFile.copy_file_data_to_server(ref _file_name, ref _file_download_name, ref _file_fullpath_name
                        , files, folder, "import", file_doc, true, false);

                        if (string.IsNullOrEmpty(file_seq) || string.IsNullOrEmpty(_file_name) || string.IsNullOrEmpty(_file_download_name) || string.IsNullOrEmpty(_file_fullpath_name))
                        { msg_error = "Invalid folder."; }
                        else
                        {
                            #region set data worksheet / import file data to database 
                            // ตรวจสอบว่าไฟล์ที่ได้มามีอยู่จริงหรือไม่
                            if (!File.Exists(_file_fullpath_name))
                            {
                                msg_error = $"The specified file does not exist.{_file_fullpath_name}";
                            }
                            else
                            {
                                if (string.IsNullOrEmpty(ClassFile.check_file_on_server(folder, _file_fullpath_name)))
                                {
                                    ClassExcel classExcel = new ClassExcel();
                                    msg_error = classExcel.import_excel_jsea_worksheet(file_seq, _file_fullpath_name, ref _dsData);
                                }
                                else { msg_error = "The file is not within the allowed directory."; }
                            }
                            #endregion set data worksheet / import file data to database 
                        }
                    }
                }
            }
            else { msg_error = "No Data."; }


            if (dtdef != null)
            {
                ClassFile.AddRowToDataTable(ref dtdef, _file_name, _file_fullpath_name, msg_error);
                if (dsData != null) _dsData.Tables.Add(dtdef.Copy());
            }

            return JsonConvert.SerializeObject(_dsData, Formatting.Indented);
        }

        public string uploadfile_data(uploadFile uploadFile, string folder)
        {
            string msg_error = "";
            DataSet _dsData = new DataSet();
            DataTable dtdef = (DataTable)ClassFile.DatatableFile();
            string _file_name = "";
            string _file_download_name = "";
            string _file_fullpath_name = "";

            if (dtdef != null)
            {
                if (uploadFile != null)
                {
                    IFormFileCollection files = uploadFile?.file_obj;
                    if (files?.Count > 0)
                    {
                        var file_seq = uploadFile?.file_seq ?? "";
                        var file_name = uploadFile?.file_name ?? "";
                        var file_part = uploadFile?.file_part ?? "";
                        var file_doc = uploadFile?.file_doc ?? "";
                        var file_sub_software = uploadFile?.sub_software ?? "";
                        if (file_sub_software != "") { folder = (file_sub_software ?? ""); }

                        msg_error = ClassFile.copy_file_data_to_server(ref _file_name, ref _file_download_name, ref _file_fullpath_name
                        , files, folder, file_part, file_doc, false, false);

                    }
                }
            }
            else { msg_error = "No Data."; }

            if (dtdef != null)
            {
                ClassFile.AddRowToDataTable(ref dtdef, _file_name, _file_fullpath_name, msg_error);
                if (dsData != null) _dsData.Tables.Add(dtdef.Copy());
            }

            return JsonConvert.SerializeObject(_dsData, Formatting.Indented);
        }

        #endregion function

        #region save data
        private string get_pha_no(string pha_sub_software, string year)
        {
            if (string.IsNullOrEmpty(pha_sub_software)) { return "No data."; }
            if (string.IsNullOrEmpty(year)) { return "No data."; }

            List<SqlParameter> parameters = new List<SqlParameter>();

            //hazop format : HAZOP-2013-1000002
            DataTable _dt = new DataTable();

            string gen_doc = "" + pha_sub_software.ToUpper() + "-" + year.ToUpper() + "-";


            sqlstr = @" select @gen_doc + right('0000000' + trim(str(coalesce(max(replace(upper(pha_no),@gen_doc,'')+1),1))),7) as pha_no ";
            sqlstr += @" from epha_t_header where lower(pha_sub_software) = lower(@pha_sub_software) and year = @year";

            if (!string.IsNullOrEmpty(gen_doc))
            {
                parameters.Add(new SqlParameter("@gen_doc", SqlDbType.VarChar, 50) { Value = gen_doc.ToUpper() ?? "" });
            }
            if (!string.IsNullOrEmpty(pha_sub_software))
            {
                parameters.Add(new SqlParameter("@pha_sub_software", SqlDbType.VarChar, 50) { Value = pha_sub_software.ToUpper() });
            }
            if (!string.IsNullOrEmpty(year))
            {
                parameters.Add(new SqlParameter("@year", SqlDbType.VarChar, 50) { Value = year.ToUpper() });
            }

            _dt = new DataTable();
            _dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);

            return (_dt?.Rows[0]["pha_no"].ToString() + "");
        }
        private int get_max_version(string seq)
        {
            if (string.IsNullOrEmpty(seq)) { return 0; }

            List<SqlParameter> parameters = new List<SqlParameter>();

            sqlstr = @"  select isnull(max(a.pha_version),1) as pha_version,  isnull(max(a.pha_version)+1,1) as pha_version_max from epha_t_header a where a.seq = @seq  ";

            parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });

            DataTable _dt = new DataTable();
            _dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
            if (_dt?.Rows.Count > 0)
            {
                return Convert.ToInt32(_dt.Rows[0]["pha_version_max"].ToString() + "");
            }
            return 0;
        }
        private int get_max(string table_name)
        {
            if (string.IsNullOrEmpty(table_name)) { return 0; }
            DataTable _dt = new DataTable();
            List<SqlParameter> parameters = new List<SqlParameter>();

            sqlstr = $@" select coalesce(max(id),0)+1 as id from {table_name}";

            _dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);

            return Convert.ToInt32(_dt.Rows[0]["id"].ToString() + "");
        }
        private void ConvertJSONresultToDataSet(string user_name, ref string msg, ref string ret, ref DataSet dsData, SetDataWorkflowModel param, string pha_status, string pha_sub_software
            )
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name)) { ret = "User is not authorized to perform this action."; return; }

            if (param == null) { msg = "No Data."; ret = "Error"; return; }
            if (string.IsNullOrEmpty(pha_status)) { msg = "No Data."; ret = "Error"; return; }
            if (string.IsNullOrEmpty(pha_sub_software)) { msg = "No Data."; ret = "Error"; return; }

            DataTable dt = new DataTable();
            cls_json = new ClassJSON();

            dsData = new DataSet();

            if (dsData != null)
            {
                #region ConvertJSONresult
                jsper = param?.json_header ?? "";
                if (jsper.Trim() == "") { msg = "No Data."; ret = "Error"; return; }
                try
                {
                    dt = new DataTable();
                    dt = cls_json.ConvertJSONresult(jsper);

                    dt.TableName = "header";
                    dsData?.Tables.Add(dt.Copy()); dsData?.AcceptChanges();
                }
                catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }


                jsper = param?.json_general ?? "";
                if (jsper.Trim() == "") { msg = "No Data."; ret = "Error"; return; }
                try
                {
                    dt = new DataTable();
                    dt = cls_json.ConvertJSONresult(jsper);
                    if (dt != null)
                    {
                        dt.TableName = "general";
                        dsData?.Tables.Add(dt.Copy()); dsData?.AcceptChanges();
                    }
                }
                catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }

                string sub_expense_type = "";
                try
                {
                    sub_expense_type = (dsData?.Tables["general"]?.Rows[0]["sub_expense_type"] + "");
                }
                catch { }


                jsper = param?.json_session ?? "";
                if (jsper.Trim() == "") { msg = "No Data."; ret = "Error"; }
                try
                {
                    dt = new DataTable();
                    dt = cls_json.ConvertJSONresult(jsper);
                    if (dt != null)
                    {
                        dt.TableName = "session";
                        dsData?.Tables.Add(dt.Copy()); dsData?.AcceptChanges();
                    }
                }
                catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; }

                jsper = param?.json_memberteam ?? "";
                try
                {
                    dt = new DataTable();
                    dt = cls_json.ConvertJSONresult(jsper);
                    if (dt != null)
                    {
                        dt.TableName = "memberteam";
                        dsData?.Tables.Add(dt.Copy()); dsData?.AcceptChanges();
                    }
                }
                catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; }

                jsper = param?.json_approver ?? "";
                try
                {
                    dt = new DataTable();
                    dt = cls_json.ConvertJSONresult(jsper);
                    if (dt != null)
                    {
                        dt.TableName = "approver";
                        dsData?.Tables.Add(dt.Copy()); dsData?.AcceptChanges();
                    }
                }
                catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; }

                jsper = param?.json_drawing ?? "";
                try
                {
                    dt = new DataTable();
                    dt = cls_json.ConvertJSONresult(jsper);
                    if (dt != null)
                    {
                        dt.TableName = "drawing";
                        dsData?.Tables.Add(dt.Copy()); dsData?.AcceptChanges();
                    }
                }
                catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; }

                if (!(pha_sub_software == "hra"))
                {
                    jsper = param?.json_functional_audition ?? "";
                    try
                    {
                        if (jsper.Trim() != "")
                        {
                            dt = new DataTable();
                            dt = cls_json.ConvertJSONresult(jsper);
                            if (dt != null)
                            {
                                dt.TableName = "functional_audition";
                                dsData?.Tables.Add(dt.Copy()); dsData?.AcceptChanges();
                            }
                        }
                    }
                    catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; }

                    jsper = param?.json_ram_level ?? "";
                    try
                    {
                        dt = new DataTable();
                        dt = cls_json.ConvertJSONresult(jsper);
                        if (dt != null)
                        {
                            dt.TableName = "ram_level";
                            dsData?.Tables.Add(dt.Copy()); dsData?.AcceptChanges();
                        }
                    }
                    catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; }

                    jsper = param?.json_ram_master ?? "";
                    try
                    {
                        dt = new DataTable();
                        dt = cls_json.ConvertJSONresult(jsper);
                        if (dt != null)
                        {
                            dt.TableName = "ram_master";
                            dsData?.Tables.Add(dt.Copy()); dsData?.AcceptChanges();
                        }
                    }
                    catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; }
                }

                jsper = param?.json_flow_action ?? "";
                try
                {
                    dt = new DataTable();
                    dt = cls_json.ConvertJSONresult(jsper);
                    if (dt != null)
                    {
                        dt.TableName = "flow_action";
                        dsData?.Tables.Add(dt.Copy()); dsData?.AcceptChanges();
                    }
                }
                catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; }

                if (dsData != null)
                {
                    if (param != null)
                    {
                        if (pha_sub_software == "hazop")
                        { ConvertJSONresultToDataSetHAZOP(user_name, ref msg, ref ret, ref dsData, param, pha_status, pha_sub_software); }
                        if (ret == "Error") { return; }
                    }
                }

                if (dsData != null)
                {
                    if (param != null)
                    {
                        if (pha_sub_software == "whatif")
                        { ConvertJSONresultToDataSetWhatif(user_name, ref msg, ref ret, ref dsData, param, pha_status, pha_sub_software); }
                        if (ret == "Error") { return; }
                    }
                }

                if (dsData != null)
                {
                    if (param != null)
                    {
                        if (pha_sub_software == "jsea")
                        { ConvertJSONresultToDataSetJSEA(user_name, ref msg, ref ret, ref dsData, param, pha_status, pha_sub_software); }
                    }
                }

                if (dsData != null)
                {
                    if (param != null)
                    {
                        if (pha_sub_software == "hra")
                        { ConvertJSONresultToDataSetHRA(user_name, ref msg, ref ret, ref dsData, param, pha_status, pha_sub_software); }
                        if (ret == "Error") { return; }
                    }
                }

                #endregion ConvertJSONresult


            }
        }
        private void ConvertJSONresultToDataSetHAZOP(string user_name, ref string msg, ref string ret, ref DataSet dsData, SetDataWorkflowModel param, string pha_status, string pha_sub_software)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name)) { ret = "Error"; return; }

            DataTable dt = new DataTable();
            cls_json = new ClassJSON();

            jsper = param.json_relatedpeople_outsider + "";
            try
            {
                if (jsper.Trim() != "")
                {
                    dt = new DataTable();
                    dt = cls_json.ConvertJSONresult(jsper);
                    if (dt != null)
                    {
                        dt.TableName = "relatedpeople_outsider";
                        dsData?.Tables.Add(dt.Copy()); dsData?.AcceptChanges();
                    }
                }
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }


            jsper = param.json_node + "";
            try
            {
                if (jsper.Trim() != "")
                {
                    dt = new DataTable();
                    dt = cls_json.ConvertJSONresult(jsper);
                    if (dt != null)
                    {
                        dt.TableName = "node";
                        dsData?.Tables.Add(dt.Copy()); dsData?.AcceptChanges();
                    }
                }
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }

            jsper = param.json_nodedrawing + "";
            try
            {
                dt = new DataTable();
                dt = cls_json.ConvertJSONresult(jsper);
                if (dt != null)
                {
                    dt.TableName = "nodedrawing";
                    dsData?.Tables.Add(dt.Copy()); dsData?.AcceptChanges();
                }
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }

            jsper = param.json_nodeguidwords + "";
            try
            {
                dt = new DataTable();
                dt = cls_json.ConvertJSONresult(jsper);
                if (dt != null)
                {
                    dt.TableName = "nodeguidwords";
                    dsData?.Tables.Add(dt.Copy()); dsData?.AcceptChanges();
                }
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }

            jsper = param.json_nodeworksheet + "";
            try
            {
                dt = new DataTable();
                dt = cls_json.ConvertJSONresult(jsper);
                if (dt != null)
                {
                    dt.TableName = "nodeworksheet";
                    dsData?.Tables.Add(dt.Copy()); dsData?.AcceptChanges();
                }
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }


        }
        private void ConvertJSONresultToDataSetWhatif(string user_name, ref string msg, ref string ret, ref DataSet dsData, SetDataWorkflowModel param, string pha_status, string pha_sub_software)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name)) { ret = "Error"; return; }

            DataTable dt = new DataTable();
            cls_json = new ClassJSON();


            //json_relatedpeople, json_list, json_listdrawing, json_listworksheet

            jsper = param.json_relatedpeople + "";
            try
            {
                if (jsper.Trim() != "")
                {
                    dt = new DataTable();
                    dt = cls_json.ConvertJSONresult(jsper);
                    if (dt != null)
                    {
                        dt.TableName = "relatedpeople";
                        dsData?.Tables.Add(dt.Copy()); dsData?.AcceptChanges();
                    }
                }
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }

            jsper = param.json_relatedpeople_outsider + "";
            try
            {
                if (jsper.Trim() != "")
                {
                    dt = new DataTable();
                    dt = cls_json.ConvertJSONresult(jsper);
                    if (dt != null)
                    {
                        dt.TableName = "relatedpeople_outsider";
                        dsData?.Tables.Add(dt.Copy()); dsData?.AcceptChanges();
                    }
                }
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }


            jsper = param.json_list + "";
            try
            {
                if (jsper.Trim() != "")
                {
                    dt = new DataTable();
                    dt = cls_json.ConvertJSONresult(jsper);
                    if (dt != null)
                    {
                        dt.TableName = "list";
                        dsData?.Tables.Add(dt.Copy()); dsData?.AcceptChanges();
                    }
                }
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }

            jsper = param.json_listdrawing + "";
            try
            {
                dt = new DataTable();
                dt = cls_json.ConvertJSONresult(jsper);
                if (dt != null)
                {
                    dt.TableName = "listdrawing";
                    dsData?.Tables.Add(dt.Copy()); dsData?.AcceptChanges();
                }
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }

            jsper = param.json_listworksheet + "";
            try
            {
                dt = new DataTable();
                dt = cls_json.ConvertJSONresult(jsper);
                if (dt != null)
                {
                    dt.TableName = "listworksheet";
                    dsData?.Tables.Add(dt.Copy()); dsData?.AcceptChanges();
                }
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }


        }
        private void ConvertJSONresultToDataSetJSEA(string user_name, ref string msg, ref string ret, ref DataSet dsData, SetDataWorkflowModel param, string pha_status, string pha_sub_software)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name)) { ret = "Error"; return; }

            DataTable dt = new DataTable();
            cls_json = new ClassJSON();

            if (true)
            {

                jsper = param.json_relatedpeople + "";
                try
                {
                    if (jsper.Trim() != "")
                    {
                        dt = new DataTable();
                        dt = cls_json.ConvertJSONresult(jsper);
                        if (dt != null)
                        {
                            dt.TableName = "relatedpeople";
                            dsData?.Tables.Add(dt.Copy()); dsData?.AcceptChanges();
                        }
                    }
                }
                catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }

                jsper = param.json_relatedpeople_outsider + "";
                try
                {
                    if (jsper.Trim() != "")
                    {
                        dt = new DataTable();
                        dt = cls_json.ConvertJSONresult(jsper);
                        if (dt != null)
                        {
                            dt.TableName = "relatedpeople_outsider";
                            dsData?.Tables.Add(dt.Copy()); dsData?.AcceptChanges();
                        }
                    }
                }
                catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }


                jsper = param.json_tasks_worksheet + "";
                try
                {
                    if (jsper.Trim() != "")
                    {
                        dt = new DataTable();
                        dt = cls_json.ConvertJSONresult(jsper);
                        if (dt != null)
                        {
                            dt.TableName = "tasks_worksheet";
                            dsData?.Tables.Add(dt.Copy()); dsData?.AcceptChanges();
                        }
                    }
                }
                catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }


            }

        }
        private void ConvertJSONresultToDataSetHRA(string user_name, ref string msg, ref string ret, ref DataSet dsData, SetDataWorkflowModel param, string pha_status, string pha_sub_software)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name)) { ret = "Error"; return; }

            DataTable dt = new DataTable();
            cls_json = new ClassJSON();

            if (true)
            {

                jsper = param.json_subareas + "";
                try
                {
                    if (jsper.Trim() != "")
                    {
                        dt = new DataTable();
                        dt = cls_json.ConvertJSONresult(jsper);
                        if (dt != null)
                        {
                            dt.TableName = "subareas";
                            dsData?.Tables.Add(dt.Copy()); dsData?.AcceptChanges();
                        }
                    }
                }
                catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; }

                jsper = param.json_hazard + "";
                try
                {
                    if (jsper.Trim() != "")
                    {
                        dt = new DataTable();
                        dt = cls_json.ConvertJSONresult(jsper);
                        if (dt != null)
                        {
                            dt.TableName = "hazard";
                            dsData?.Tables.Add(dt.Copy()); dsData?.AcceptChanges();
                        }
                    }
                }
                catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; }

                jsper = param.json_tasks + "";
                try
                {
                    if (jsper.Trim() != "")
                    {
                        dt = new DataTable();
                        dt = cls_json.ConvertJSONresult(jsper);
                        if (dt != null)
                        {
                            dt.TableName = "tasks";
                            dsData?.Tables.Add(dt.Copy()); dsData?.AcceptChanges();
                        }
                    }
                }
                catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; }

                jsper = param.json_descriptions + "";
                try
                {
                    if (jsper.Trim() != "")
                    {
                        dt = new DataTable();
                        dt = cls_json.ConvertJSONresult(jsper);
                        if (dt != null)
                        {
                            dt.TableName = "descriptions";
                            dsData?.Tables.Add(dt.Copy()); dsData?.AcceptChanges();
                        }
                    }
                }
                catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; }

                jsper = param.json_workers + "";
                try
                {
                    if (jsper.Trim() != "")
                    {
                        dt = new DataTable();
                        dt = cls_json.ConvertJSONresult(jsper);
                        if (dt != null)
                        {
                            dt.TableName = "workers";
                            dsData?.Tables.Add(dt.Copy()); dsData?.AcceptChanges();
                        }
                    }
                }
                catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; }

                jsper = param.json_worksheet + "";
                try
                {
                    if (jsper.Trim() != "")
                    {
                        dt = new DataTable();
                        dt = cls_json.ConvertJSONresult(jsper);
                        if (dt != null)
                        {
                            dt.TableName = "worksheet";
                            dsData?.Tables.Add(dt.Copy()); dsData?.AcceptChanges();
                        }
                    }
                }
                catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; }

                jsper = param.json_recommendations + "";
                try
                {
                    if (jsper.Trim() != "")
                    {
                        dt = new DataTable();
                        dt = cls_json.ConvertJSONresult(jsper);
                        if (dt != null)
                        {
                            dt.TableName = "recommendations";
                            dsData?.Tables.Add(dt.Copy()); dsData?.AcceptChanges();
                        }
                    }
                }
                catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; }

            }

        }

        public string keep_version(string user_name, ref string seq, ref string version, string pha_status_new, string pha_sub_software
            , Boolean bReteMOC, Boolean bRetTA2, Boolean bTA2ApproveAll, Boolean bComplate
            , Boolean bTA2ApproveMOC = false)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; ; }
            if (string.IsNullOrEmpty(seq)) { return "Invalid Seq."; }

            string sqlstr_field = "";
            List<SqlParameter> parameters = new List<SqlParameter>();

            DataTable dtColumnsHeader = new DataTable();
            DataTable dtColumns = new DataTable();
            //get data tabel all 
            string sTable = "epha_t_general,epha_t_functional_audition,epha_t_session,epha_t_member_team,epha_t_approver,epha_t_approver_ta3,epha_t_relatedpeople,epha_t_relatedpeople_outsider,epha_t_drawing";
            sTable += ",epha_t_recommendations,epha_t_recom_setting,epha_t_recom_follow";
            if (pha_sub_software == "hazop") { sTable += ",epha_t_node,epha_t_node_drawing,epha_t_node_guide_words,epha_t_node_worksheet"; }
            if (pha_sub_software == "whatif") { sTable += ",epha_t_list,epha_t_list_drawing,epha_t_list_worksheet"; }
            if (pha_sub_software == "jsea") { sTable += ",epha_t_tasks_worksheet"; }
            if (pha_sub_software == "hra") { sTable += ",epha_t_table1_hazard,epha_t_table1_subareas,epha_t_table2_tasks,epha_t_table2_workers,epha_t_table2_descriptions,epha_t_table3_worksheet"; }
            string[] xsplitTable = (sTable).Split(',');
            string seq_session_active = "";

            if (true)
            {
                if (bTA2ApproveMOC && (pha_status_new == "11"))
                {
                    cls = new ClassFunctions();
                    sqlstr = @"select a.pha_no,  max(b.seq) seq, max(b.id_pha) as id_pha from epha_t_header a inner join  epha_t_session b on a.id = b.id_pha 
                       where a.seq = @seq group by a.pha_no ";

                    parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });

                    DataTable dtSession = new DataTable();
                    dtSession = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
                    if (dtSession?.Rows.Count > 0) { seq_session_active = (dtSession?.Rows[0]["seq"] + ""); }
                }

                #region query copy data  
                cls = new ClassFunctions();
                sqlstr = @" select lower(table_name) as table_name,lower(column_name) as column_name from  information_schema.columns
                        where lower(table_name) in ( 'epha_t_header') 
                        and lower(column_name) not in ('seq','id','pha_status','pha_version','pha_version_desc','pha_version_text','next_version','create_date','update_date')
                        order by table_name ";

                parameters = new List<SqlParameter>();
                dtColumnsHeader = new DataTable();
                dtColumnsHeader = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);

                cls = new ClassFunctions();
                sqlstr = @" select lower(table_name) as table_name,lower(column_name) as column_name from  information_schema.columns
                        where lower(column_name) not in ('id_pha','create_date','update_date') ";

                sqlstr += " and lower(table_name) in (";
                for (int i = 0; i < xsplitTable.Length; i++)
                {
                    if (i > 0) { sqlstr += ", "; }
                    string paramName = "@sTable" + i;

                    sqlstr += paramName;
                    parameters.Add(new SqlParameter(paramName, SqlDbType.VarChar, 200) { Value = xsplitTable[i].Trim() });
                }
                sqlstr += ") order by table_name";

                dtColumns = new DataTable();
                dtColumns = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
                #endregion query copy data 
            }
            if (dtColumnsHeader?.Rows.Count == 0) { return "No Data."; }
            if (dtColumns?.Rows.Count == 0) { return "No Data."; }

            if (dtColumnsHeader?.Rows.Count > 0)
            {
                if (dtColumns?.Rows.Count > 0)
                {

                    //action type = insert,update,delete,old_data 
                    string year_now = DateTime.Now.ToString("yyyy");
                    string seq_header_max = get_max("epha_t_header").ToString() ?? "0";
                    string version_max = get_max_version(seq ?? "").ToString() ?? "0";


                    string ret = "true";
                    //กรณีที่มีมากกว่า 0 ให้ keep version เดิมและ new version ใหม่  
                    //update seq_header_now to id_pha  
                    cls = new ClassFunctions();
                    using (var conn = new SqlConnection(ClassConnectionDb.ConnectionString()))
                    {
                        conn.Open();

                        using (var transaction = conn.BeginTransaction())
                        {
                            try
                            {
                                //table epha_t_header
                                if (true)
                                {
                                    parameters = new List<SqlParameter>();

                                    string rev_no = (version_max).ToString(); string rev_text = ""; string rev_desc = "";
                                    findRevisionText(pha_sub_software, rev_no, ref rev_text, ref rev_desc, bReteMOC, bRetTA2, bTA2ApproveAll, bComplate, bTA2ApproveMOC);

                                    //header
                                    string table = "epha_t_header";
                                    string columns = "";
                                    DataRow[] dr = dtColumnsHeader.AsEnumerable().Where(row => row.Field<string>("table_name") == table).ToArray();

                                    if (dr?.Length > 0)
                                    {
                                        for (int c = 0; c < dr?.Length; c++)
                                        {
                                            if (c > 0) { columns += ","; }
                                            columns += dr[c]["column_name"] + "";
                                        }
                                        sqlstr_field = $"INSERT INTO {table} (seq, id, pha_status, pha_version, pha_version_text, pha_version_desc, next_version, create_date, update_date, {columns})" +
                                                   $" SELECT @seq_header_max AS seq, @seq_header_max AS id, @pha_status_new AS pha_status" +
                                                   $" , @version_max AS pha_version, @rev_text AS pha_version_text, @rev_desc AS pha_version_desc" +
                                                   $" , 'false' AS next_version, GETDATE() AS create_date, NULL AS update_date, {columns}" +
                                                   $" FROM {table} WHERE seq = @seq";

                                        parameters.Add(new SqlParameter("@seq_header_max", seq_header_max));
                                        parameters.Add(new SqlParameter("@pha_status_new", pha_status_new));
                                        parameters.Add(new SqlParameter("@version_max", version_max)); // version_max เป็น string
                                        parameters.Add(new SqlParameter("@rev_text", ConvertToDBNull(rev_text)));
                                        parameters.Add(new SqlParameter("@rev_desc", ConvertToDBNull(rev_desc)));
                                        parameters.Add(new SqlParameter("@seq", seq));

                                        ClassConnectionDb cls_conn = new ClassConnectionDb();
                                        ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr_field, parameters, conn, transaction);

                                        if (ret != "true") { goto Next_Line; }
                                    }


                                }

                                //table Other
                                if (true)
                                {
                                    for (int t = 0; t < xsplitTable.Length; t++)
                                    {
                                        parameters = new List<SqlParameter>();
                                        string table = xsplitTable[t].ToString();
                                        string columns = "";
                                        DataRow[] dr = dtColumns.AsEnumerable().Where(row => row.Field<string>("table_name") == table).ToArray();

                                        if (dr?.Length > 0)
                                        {
                                            for (int c = 0; c < dr.Length; c++)
                                            {
                                                if (c > 0) { columns += ","; }
                                                columns += dr[c]["column_name"] + "";
                                            }
                                            sqlstr_field = $" insert into {table} (id_pha,create_date,update_date,{columns}) " +
                                                           $" select @seq_header_max as id_pha,getdate() as create_date, null as update_date,{columns} from {table}  " +
                                                           $" where id_pha = @seq";

                                            parameters.Add(new SqlParameter("@seq_header_max", seq_header_max));
                                            parameters.Add(new SqlParameter("@seq", seq));

                                            ClassConnectionDb cls_conn = new ClassConnectionDb();
                                            ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr_field, parameters, conn, transaction);

                                            if (ret != "true") { goto Next_Line; }
                                        }
                                    }
                                }
                                if (ret == "true")
                                {
                                    //update seq_header_max now to seq
                                    seq = seq_header_max;

                                    //เนื่องจากใช้ seq, id เดียวกันกับ id_pha
                                    if (true)
                                    {
                                        parameters = new List<SqlParameter>();
                                        sqlstr_field = @" update EPHA_T_GENERAL set seq = id_pha,  id = id_pha where id_pha = @seq";

                                        parameters.Add(new SqlParameter("@seq", seq));
                                        ClassConnectionDb cls_conn = new ClassConnectionDb();
                                        ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr_field, parameters, conn, transaction);

                                        if (ret != "true") { goto Next_Line; }
                                    }

                                    //bTA2ApproveMOC --> ต้องกด Submit TA2 Review and Approve for MOC 
                                    // action_to_approve_moc -> 1 : ส่งไป eMOC, 2 : Approve & Review Complate 
                                    if (bTA2ApproveMOC && (pha_status_new == "11"))
                                    {
                                        parameters = new List<SqlParameter>();
                                        sqlstr_field = "update epha_t_session set date_to_approve_moc = getdate() , action_to_approve_moc = 1  where id_pha = @seq and id = @seq_session_active";

                                        parameters.Add(new SqlParameter("@seq_session_active", seq_session_active));
                                        parameters.Add(new SqlParameter("@seq", seq));

                                        ClassConnectionDb cls_conn = new ClassConnectionDb();
                                        ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr_field, parameters, conn, transaction);

                                        if (ret != "true") { goto Next_Line; }
                                    }
                                }

                            Next_Line:;

                                if (ret == "true")
                                {
                                    transaction.Commit();
                                }
                                else
                                {
                                    throw new Exception(ret);
                                }
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
            }

            return "No Data.";
        }

        public string update_revision_table_now(string user_name, string seq_header, string pha_no, string version, string pha_status, string expense_type, string pha_sub_software)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            if (string.IsNullOrEmpty(seq_header)) { return "No data."; }
            //if (string.IsNullOrEmpty(pha_no)) { return "No data."; }
            //if (string.IsNullOrEmpty(version)) { return "No data."; }
            if (string.IsNullOrEmpty(pha_status)) { return "No data."; }
            //if (string.IsNullOrEmpty(expense_type)) { return "No data."; }
            if (string.IsNullOrEmpty(pha_sub_software)) { return "No data."; }

            string ret = "true";

            #region update Revision 
            Boolean bReteMOC = false; Boolean bRetTA2 = false; Boolean bTA2ApproveAll = false; Boolean bComplate = false;
            string rev_no = version.ToString(); string rev_text = ""; string rev_desc = "";

            expense_type = expense_type ?? "opex";
            bReteMOC = (expense_type.ToLower() == "opex");
            bRetTA2 = (expense_type.ToLower() == "capex");

            bTA2ApproveAll = (pha_status == "13");
            bComplate = (pha_status == "91");

            findRevisionText(pha_sub_software, rev_no, ref rev_text, ref rev_desc, bReteMOC, bRetTA2, bTA2ApproveAll, bComplate);

            #endregion update Revision

            #region update pha status text 

            List<SqlParameter> parameters = new List<SqlParameter>();

            string sqlstr = "UPDATE epha_t_header SET ";
            sqlstr += "PHA_VERSION_TEXT = @rev_text, ";
            sqlstr += "PHA_VERSION_DESC = @rev_desc ";
            sqlstr += "WHERE SEQ = @seq_header ";

            parameters.Add(ClassConnectionDb.CreateSqlParameter("@rev_text", SqlDbType.VarChar, ConvertToDBNull(rev_text), 200));
            parameters.Add(ClassConnectionDb.CreateSqlParameter("@rev_desc", SqlDbType.VarChar, ConvertToDBNull(rev_desc), 200));
            parameters.Add(ClassConnectionDb.CreateSqlParameter("@seq_header", SqlDbType.Int, seq_header));

            if (!string.IsNullOrEmpty(pha_no))
            {
                sqlstr += " AND PHA_NO = @pha_no";
                parameters.Add(ClassConnectionDb.CreateSqlParameter("@pha_no", SqlDbType.VarChar, pha_no, 200));
            }

            try
            {
                ret = ClassConnectionDb.ExecuteNonQuerySQL(sqlstr, parameters);
            }
            catch (Exception ex)
            {
                // Log exception (consider using a logging framework)
                Console.WriteLine(ex.Message);
                ret = "error: " + ex.Message;
            }

            #endregion update pha status 

            return ret;
        }

        public string update_status_table_approver_sendback(string user_name, string seq_header)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            //update pha status
            if (string.IsNullOrEmpty(seq_header)) { return "No data."; }
            string ret = "true";

            List<SqlParameter> parameters = new List<SqlParameter>();

            sqlstr = "update epha_t_approver set comment = null, action_status = nul, action_review = 0  where action_review not in (2)";
            sqlstr += " and id_pha = @seq_header";

            parameters.Add(ClassConnectionDb.CreateSqlParameter("@seq_header", SqlDbType.Int, seq_header));

            try
            {
                ret = ClassConnectionDb.ExecuteNonQuerySQL(sqlstr, parameters);
            }
            catch (Exception ex)
            {
                // Log exception (consider using a logging framework)
                Console.WriteLine(ex.Message);
                ret = "error: " + ex.Message;
            }
            return ret;
        }
        public string update_status_table_now(string user_name, string seq_header, string pha_no, string pha_status)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }
            if (string.IsNullOrEmpty(seq_header)) { return "No data."; }
            if (string.IsNullOrEmpty(pha_status)) { return "No data."; }

            string ret = "true";
            List<SqlParameter> parameters = new List<SqlParameter>();

            sqlstr = "update  epha_t_header set ";
            sqlstr += " PHA_STATUS = @pha_status";
            sqlstr += " where SEQ = @seq_header ";

            parameters.Add(ClassConnectionDb.CreateSqlParameter("@pha_status", SqlDbType.Int, pha_status));
            parameters.Add(ClassConnectionDb.CreateSqlParameter("@seq_header", SqlDbType.Int, seq_header));

            if (!string.IsNullOrEmpty(pha_no))
            {
                sqlstr += " AND PHA_NO = @pha_no";
                parameters.Add(ClassConnectionDb.CreateSqlParameter("@pha_no", SqlDbType.VarChar, pha_no, 200));
            }

            try
            {
                ret = ClassConnectionDb.ExecuteNonQuerySQL(sqlstr, parameters);
            }
            catch (Exception ex)
            {
                // Log exception (consider using a logging framework)
                Console.WriteLine(ex.Message);
                ret = "error: " + ex.Message;
            }
            return ret;
        }

        public string copy_document_file_responder_to_reviewer(string user_name, string seq_header, string sub_software)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; ; }
            if (string.IsNullOrEmpty(seq_header)) { return "No data."; }
            if (string.IsNullOrEmpty(sub_software)) { return "No data."; }

            List<SqlParameter> parameters = new List<SqlParameter>();
            parameters.Add(ClassConnectionDb.CreateSqlParameter("@seq_header", SqlDbType.Int, seq_header));

            string module_name = "FollowUp";

            sqlstr = @" select id_pha, id_worksheet, no, document_file_name as document_file_name_def, document_file_path as document_file_path_def
                        , 'review_followup' as document_module 
                        , null as seq, null as id, null as document_file_name, null as document_file_path
                        from EPHA_T_DRAWING_WORKSHEET where lower(document_module) = lower('review_followup')
                        and id_pha = @seq_header";

            DataTable dtDocReviewer = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);

            if (dtDocReviewer?.Rows.Count > 0) { return "false"; }

            sqlstr = @" select a.*, document_file_name as document_file_name_def, document_file_path as document_file_path_def
                        from EPHA_T_DRAWING_WORKSHEET a where lower(document_module) = lower('followup')
                        and id_pha = @seq_header";

            DataTable dtDocResponder = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
            if (dtDocResponder?.Rows.Count == 0) { return "false"; }

            // กำหนดให้ copy file มาตั้งต้น
            // D:\dotnet6-epha-api\dotnet6-epha-api\wwwroot\AttachedFileTemp\FollowUp
            ////https://localhost:7098/api/AttachedFileTemp/FollowUp/HAZOP Report 202311281602.pdf
            //string _DownloadPath = @"/AttachedFileTemp/" + module_name + @"/";
            //string _FolderTemplate = MapPathFiles(@"/wwwroot/AttachedFileTemp/" + module_name + @"/");
            string _DownloadPath = ""; string _FolderTemplate = "";
            var fileHandler = new FileHandler();
            fileHandler.HandleFilesAttachedFileTemp(module_name, ref _DownloadPath, ref _FolderTemplate);

            DataTable dtNewFile = new DataTable();
            if (dtDocResponder != null) { dtNewFile = dtDocResponder.Clone(); }
            else { return "false"; }

            int iseq_max = Convert.ToInt32(get_max("EPHA_T_DRAWING_WORKSHEET").ToString());

            Boolean bUpdate = false;
            foreach (DataRow row in dtDocResponder.Rows)
            {
                //HAZOP Report 202311281602.pdf	
                string document_file_name_def = (row["document_file_name_def"] + "").ToString();

                //https://localhost:7098/api/AttachedFileTemp/FollowUp/HAZOP Report 202311281602.pdf
                string document_file_path_def = (row["document_file_path_def"] + "").ToString();

                //xxxxx\AttachedFileTemp\FollowUp\HAZOP Report 202311281602.pdf
                string file_name_def = ((document_file_path_def + "").ToUpper()).Split(_DownloadPath.ToUpper())[1].ToString();

                string[] xsplit_surname = file_name_def.Split(".");
                string file_name_new = (sub_software + " " + module_name + " " + "-" + iseq_max + DateTime.Now.ToString("yyyyMMddHHmm") + "." + xsplit_surname[xsplit_surname.Length - 1].ToString()).ToUpper(); ;

                string source_file = _FolderTemplate + file_name_def;
                string dest_file = _FolderTemplate + file_name_new;

                if (File.Exists(source_file))
                {
                    System.IO.File.Copy(source_file, dest_file);
                    if (true)
                    {
                        row["seq"] = iseq_max;
                        row["id"] = iseq_max;
                        row["document_file_name"] = document_file_name_def;
                        row["document_file_path"] = (xsplit_surname[0] + "").ToString() + _DownloadPath + file_name_new;
                        row["document_module"] = module_name;

                        dtNewFile.ImportRow(row); dtNewFile.AcceptChanges();
                        bUpdate = true;

                        iseq_max++;
                    }

                }

            }

            if (bUpdate)
            {
                DataTable dt = dtNewFile.Copy();
                dt.AcceptChanges();
                cls = new ClassFunctions();

                using (var conn = new SqlConnection(ClassConnectionDb.ConnectionString()))
                {
                    conn.Open();

                    using (var transaction = conn.BeginTransaction())
                    {
                        try
                        {
                            if (dt.Rows.Count > 0)
                            {
                                for (int i = 0; i < dt?.Rows.Count; i++)
                                {
                                    #region insert
                                    sqlstr = "INSERT INTO EPHA_T_DRAWING_WORKSHEET (" +
                                        "SEQ,ID,ID_PHA,ID_WORKSHEET,NO,DOCUMENT_NAME,DOCUMENT_NO,DOCUMENT_FILE_NAME,DOCUMENT_FILE_PATH,DOCUMENT_FILE_SIZE,DESCRIPTIONS,DOCUMENT_MODULE" +
                                        ",CREATE_DATE,UPDATE_DATE,CREATE_BY,UPDATE_BY" +
                                        ") VALUES (@SEQ, @ID, @ID_PHA, @ID_WORKSHEET, @NO, @DOCUMENT_NAME, @DOCUMENT_NO, @DOCUMENT_FILE_NAME, @DOCUMENT_FILE_PATH, @DOCUMENT_FILE_SIZE, @DESCRIPTIONS, @DOCUMENT_MODULE, GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";
                                    #endregion insert 

                                    parameters = new List<SqlParameter>();
                                    parameters.Add(ClassConnectionDb.CreateSqlParameter("@SEQ", SqlDbType.Int, dt.Rows[i]["SEQ"] ?? ""));
                                    parameters.Add(ClassConnectionDb.CreateSqlParameter("@ID", SqlDbType.Int, dt.Rows[i]["ID"] ?? ""));
                                    parameters.Add(ClassConnectionDb.CreateSqlParameter("@ID_PHA", SqlDbType.Int, dt.Rows[i]["ID_PHA"] ?? ""));
                                    parameters.Add(ClassConnectionDb.CreateSqlParameter("@ID_WORKSHEET", SqlDbType.Int, dt.Rows[i]["ID_WORKSHEET"] ?? ""));
                                    parameters.Add(ClassConnectionDb.CreateSqlParameter("@NO", SqlDbType.Int, dt.Rows[i]["NO"] ?? ""));
                                    parameters.Add(ClassConnectionDb.CreateSqlParameter("@DOCUMENT_NAME", SqlDbType.VarChar, dt.Rows[i]["DOCUMENT_NAME"] ?? "", 4000));
                                    parameters.Add(ClassConnectionDb.CreateSqlParameter("@DOCUMENT_NO", SqlDbType.VarChar, dt.Rows[i]["DOCUMENT_NO"] ?? "", 4000));
                                    parameters.Add(ClassConnectionDb.CreateSqlParameter("@DOCUMENT_FILE_NAME", SqlDbType.VarChar, dt.Rows[i]["DOCUMENT_FILE_NAME"] ?? "", 4000));
                                    parameters.Add(ClassConnectionDb.CreateSqlParameter("@DOCUMENT_FILE_PATH", SqlDbType.VarChar, dt.Rows[i]["DOCUMENT_FILE_PATH"] ?? "", 4000));
                                    //parameters.Add(ClassConnectionDb.CreateSqlParameter("@DOCUMENT_FILE_SIZE", SqlDbType.Int, dt.Rows[i]["DOCUMENT_FILE_SIZE"] ?? ""));
                                    parameters.Add(new SqlParameter("@DOCUMENT_FILE_SIZE", SqlDbType.Int)
                                    {
                                        Value = dt.Rows[i].Table.Columns.Contains("DOCUMENT_FILE_SIZE") ? ConvertToDBNull(dt.Rows[i]["DOCUMENT_FILE_SIZE"]) : DBNull.Value
                                    });
                                    parameters.Add(ClassConnectionDb.CreateSqlParameter("@DESCRIPTIONS", SqlDbType.VarChar, dt.Rows[i]["DESCRIPTIONS"] ?? "", 4000));
                                    parameters.Add(ClassConnectionDb.CreateSqlParameter("@DOCUMENT_MODULE", SqlDbType.VarChar, dt.Rows[i]["DOCUMENT_MODULE"] ?? "", 4000));
                                    parameters.Add(ClassConnectionDb.CreateSqlParameter("@CREATE_BY", SqlDbType.VarChar, dt.Rows[i]["CREATE_BY"] ?? "", 50));
                                    parameters.Add(ClassConnectionDb.CreateSqlParameter("@UPDATE_BY", SqlDbType.VarChar, dt.Rows[i]["UPDATE_BY"] ?? "", 50));

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
                                throw new Exception(ret);
                            }
                        }
                        catch (Exception ex)
                        {
                            transaction.Rollback();
                            // Log exception (consider using a logging framework)
                            Console.WriteLine(ex.Message);
                            ret = "error: " + ex.Message;
                        }
                    }
                }
            }



            return ret;
        }

        #endregion save data

        #region set page worksheet details 

        public string set_header(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet? dsData, ref string seq_header_now, ref string version_now, Boolean submit_generate)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            List<SqlParameter> parameters = new List<SqlParameter>();
            string ret = "true";
            try
            {
                DataTable dt = dsData?.Tables["header"]?.Copy() ?? new DataTable();
                if (dt?.Rows?.Count == 0) { return "true"; }

                ret = UpdateDataHeader(user_name, conn, transaction, dt, ref seq_header_now, ref version_now, submit_generate);
                if (ret != "true") throw new Exception(ret);
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }
            return ret;
        }
        public string UpdateDataHeader(string user_name, SqlConnection conn, SqlTransaction transaction, DataTable? dt, ref string seq_header_now, ref string version_now, Boolean submit_generate)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";
            try
            {
                for (int i = 0; i < dt?.Rows.Count; i++)
                {
                    List<SqlParameter> parameters = new List<SqlParameter>();
                    string action_type = dt?.Rows[i]["action_type"].ToString() ?? "";
                    string pha_version = dt?.Rows[i]["pha_version"].ToString() ?? "";

                    #region change version  
                    if (pha_version == "0" && action_type != "delete" && submit_generate)
                    {
                        pha_version = "1";
                        dt.Rows[i]["PHA_VERSION_TEXT"] = "A";
                        dt.Rows[i]["PHA_VERSION_DESC"] = "Issued for Review";
                    }
                    version_now = pha_version;
                    #endregion change version 

                    string sqlstr = "";

                    if (action_type == "insert")
                    {
                        #region insert
                        sqlstr = "INSERT INTO epha_t_header(SEQ, ID, YEAR, PHA_NO, PHA_VERSION, PHA_VERSION_TEXT, PHA_VERSION_DESC, PHA_STATUS, PHA_REQUEST_BY, PHA_SUB_SOFTWARE, " +
                                 "REQUEST_APPROVER, APPROVER_USER_NAME, APPROVER_USER_DISPLAYNAME, APPROVE_ACTION_TYPE, APPROVE_STATUS, APPROVE_COMMENT, " +
                                 "REQUEST_USER_NAME, REQUEST_USER_DISPLAYNAME, SAFETY_CRITICAL_EQUIPMENT_SHOW, FLOW_MAIL_TO_MEMBER, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                 "VALUES (@SEQ, @ID, @YEAR, @PHA_NO, @PHA_VERSION, @PHA_VERSION_TEXT, @PHA_VERSION_DESC, @PHA_STATUS, @PHA_REQUEST_BY, @PHA_SUB_SOFTWARE, " +
                                 "@REQUEST_APPROVER, @APPROVER_USER_NAME, @APPROVER_USER_DISPLAYNAME, @APPROVE_ACTION_TYPE, @APPROVE_STATUS, @APPROVE_COMMENT, " +
                                 "@REQUEST_USER_NAME, @REQUEST_USER_DISPLAYNAME, @SAFETY_CRITICAL_EQUIPMENT_SHOW, @FLOW_MAIL_TO_MEMBER, GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_header_now) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_header_now) });
                        parameters.Add(new SqlParameter("@YEAR", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["YEAR"]) });
                        parameters.Add(new SqlParameter("@PHA_NO", SqlDbType.VarChar, 200) { Value = ConvertToDBNull(dt.Rows[i]["PHA_NO"]) });
                        parameters.Add(new SqlParameter("@PHA_VERSION", SqlDbType.Int) { Value = ConvertToIntOrDBNull(pha_version) });
                        parameters.Add(new SqlParameter("@PHA_VERSION_TEXT", SqlDbType.VarChar, 200) { Value = ConvertToDBNull(dt.Rows[i]["PHA_VERSION_TEXT"]) });
                        parameters.Add(new SqlParameter("@PHA_VERSION_DESC", SqlDbType.VarChar, 200) { Value = ConvertToDBNull(dt.Rows[i]["PHA_VERSION_DESC"]) });
                        parameters.Add(new SqlParameter("@PHA_STATUS", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["PHA_STATUS"]) });
                        parameters.Add(new SqlParameter("@PHA_REQUEST_BY", SqlDbType.VarChar, 200) { Value = ConvertToDBNull(dt.Rows[i]["PHA_REQUEST_BY"]) });
                        parameters.Add(new SqlParameter("@PHA_SUB_SOFTWARE", SqlDbType.VarChar, 200) { Value = ConvertToDBNull(dt.Rows[i]["PHA_SUB_SOFTWARE"]) });
                        parameters.Add(new SqlParameter("@REQUEST_APPROVER", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["REQUEST_APPROVER"]) });
                        parameters.Add(new SqlParameter("@APPROVER_USER_NAME", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["APPROVER_USER_NAME"]) });
                        parameters.Add(new SqlParameter("@APPROVER_USER_DISPLAYNAME", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["APPROVER_USER_DISPLAYNAME"]) });
                        parameters.Add(new SqlParameter("@APPROVE_ACTION_TYPE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["APPROVE_ACTION_TYPE"]) });
                        parameters.Add(new SqlParameter("@APPROVE_STATUS", SqlDbType.VarChar, 200) { Value = ConvertToDBNull(dt.Rows[i]["APPROVE_STATUS"]) });
                        parameters.Add(new SqlParameter("@APPROVE_COMMENT", SqlDbType.VarChar, 200) { Value = ConvertToDBNull(dt.Rows[i]["APPROVE_COMMENT"]) });
                        parameters.Add(new SqlParameter("@REQUEST_USER_NAME", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["REQUEST_USER_NAME"]) });
                        parameters.Add(new SqlParameter("@REQUEST_USER_DISPLAYNAME", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["REQUEST_USER_DISPLAYNAME"]) });
                        parameters.Add(new SqlParameter("@SAFETY_CRITICAL_EQUIPMENT_SHOW", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SAFETY_CRITICAL_EQUIPMENT_SHOW"]) });
                        parameters.Add(new SqlParameter("@FLOW_MAIL_TO_MEMBER", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["FLOW_MAIL_TO_MEMBER"]) });
                        parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["CREATE_BY"]) });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });

                        #endregion insert
                    }
                    else if (action_type == "update")
                    {
                        seq_header_now = dt?.Rows[i]["SEQ"].ToString() ?? "";

                        #region update
                        sqlstr = "UPDATE epha_t_header SET PHA_VERSION = @PHA_VERSION, PHA_VERSION_TEXT = @PHA_VERSION_TEXT, PHA_VERSION_DESC = @PHA_VERSION_DESC, " +
                                 "PHA_STATUS = @PHA_STATUS, PHA_REQUEST_BY = @PHA_REQUEST_BY, PHA_SUB_SOFTWARE = @PHA_SUB_SOFTWARE, " +
                                 "REQUEST_APPROVER = @REQUEST_APPROVER, APPROVER_USER_NAME = @APPROVER_USER_NAME, APPROVER_USER_DISPLAYNAME = @APPROVER_USER_DISPLAYNAME, " +
                                 "APPROVE_ACTION_TYPE = @APPROVE_ACTION_TYPE, APPROVE_STATUS = @APPROVE_STATUS, APPROVE_COMMENT = @APPROVE_COMMENT, " +
                                 "REQUEST_USER_NAME = @REQUEST_USER_NAME, REQUEST_USER_DISPLAYNAME = @REQUEST_USER_DISPLAYNAME, " +
                                 "SAFETY_CRITICAL_EQUIPMENT_SHOW = @SAFETY_CRITICAL_EQUIPMENT_SHOW, FLOW_MAIL_TO_MEMBER = @FLOW_MAIL_TO_MEMBER, " +
                                 "UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY WHERE SEQ = @SEQ AND ID = @ID AND YEAR = @YEAR AND PHA_NO = @PHA_NO";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@YEAR", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["YEAR"]) });
                        parameters.Add(new SqlParameter("@PHA_NO", SqlDbType.VarChar, 200) { Value = ConvertToDBNull(dt.Rows[i]["PHA_NO"]) });
                        parameters.Add(new SqlParameter("@PHA_VERSION", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["PHA_VERSION"]) });
                        parameters.Add(new SqlParameter("@PHA_VERSION_TEXT", SqlDbType.VarChar, 200) { Value = ConvertToDBNull(dt.Rows[i]["PHA_VERSION_TEXT"]) });
                        parameters.Add(new SqlParameter("@PHA_VERSION_DESC", SqlDbType.VarChar, 200) { Value = ConvertToDBNull(dt.Rows[i]["PHA_VERSION_DESC"]) });
                        parameters.Add(new SqlParameter("@PHA_STATUS", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["PHA_STATUS"]) });
                        parameters.Add(new SqlParameter("@PHA_REQUEST_BY", SqlDbType.VarChar, 200) { Value = ConvertToDBNull(dt.Rows[i]["PHA_REQUEST_BY"]) });
                        parameters.Add(new SqlParameter("@PHA_SUB_SOFTWARE", SqlDbType.VarChar, 200) { Value = ConvertToDBNull(dt.Rows[i]["PHA_SUB_SOFTWARE"]) });
                        parameters.Add(new SqlParameter("@REQUEST_APPROVER", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["REQUEST_APPROVER"]) });
                        parameters.Add(new SqlParameter("@APPROVER_USER_NAME", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["APPROVER_USER_NAME"]) });
                        parameters.Add(new SqlParameter("@APPROVER_USER_DISPLAYNAME", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["APPROVER_USER_DISPLAYNAME"]) });
                        parameters.Add(new SqlParameter("@APPROVE_ACTION_TYPE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["APPROVE_ACTION_TYPE"]) });
                        parameters.Add(new SqlParameter("@APPROVE_STATUS", SqlDbType.VarChar, 200) { Value = ConvertToDBNull(dt.Rows[i]["APPROVE_STATUS"]) });
                        parameters.Add(new SqlParameter("@APPROVE_COMMENT", SqlDbType.VarChar, 200) { Value = ConvertToDBNull(dt.Rows[i]["APPROVE_COMMENT"]) });
                        parameters.Add(new SqlParameter("@REQUEST_USER_NAME", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["REQUEST_USER_NAME"]) });
                        parameters.Add(new SqlParameter("@REQUEST_USER_DISPLAYNAME", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["REQUEST_USER_DISPLAYNAME"]) });
                        parameters.Add(new SqlParameter("@SAFETY_CRITICAL_EQUIPMENT_SHOW", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SAFETY_CRITICAL_EQUIPMENT_SHOW"]) });
                        parameters.Add(new SqlParameter("@FLOW_MAIL_TO_MEMBER", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["FLOW_MAIL_TO_MEMBER"]) });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });

                        #endregion update
                    }
                    else if (action_type == "delete")
                    {
                        #region delete
                        sqlstr = "DELETE FROM epha_t_header WHERE SEQ = @SEQ AND ID = @ID AND YEAR = @YEAR AND PHA_NO = @PHA_NO";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@YEAR", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["YEAR"]) });
                        parameters.Add(new SqlParameter("@PHA_NO", SqlDbType.VarChar, 200) { Value = ConvertToDBNull(dt.Rows[i]["PHA_NO"]) });

                        #endregion delete
                    }

                    if (action_type != "")
                    {
                        ClassConnectionDb cls_conn = new ClassConnectionDb();
                        ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                        if (ret != "true") break;
                    }
                }

            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return ret;
        }

        public string set_parti(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet? dsData, string seq_header_now, DataSet dsDataOld, string flow_action = "")
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            List<SqlParameter> parameters = new List<SqlParameter>();
            DataTable dt = new DataTable();
            DataTable dtMainDelete = new DataTable();
            dtMainDelete.Columns.Add("SEQ", typeof(string));
            dtMainDelete.Columns.Add("ID", typeof(string));
            dtMainDelete.Columns.Add("ID_PHA", typeof(string));
            dtMainDelete.Columns.Add("ID_SESSION", typeof(string));

            if (flow_action == "change_approver")
            {
                //กรณีที่ change approver ให้เครียร์ ACTION_REVIEW, DATE_REVIEW, COMMENT, APPROVER_ACTION_TYPE, APPROVER_TYPE, ACTION_STATUS
                var requiredColumns = new List<string> { "ACTION_REVIEW", "DATE_REVIEW", "COMMENT", "APPROVER_ACTION_TYPE", "APPROVER_TYPE", "ACTION_STATUS" };
                add_columns_in_table(user_name, ref dsData, "approver", requiredColumns);
            }

            string ret = "true";
            string _module_name = (dsData?.Tables["header"]?.Rows[0]["pha_sub_software"] + "");//sub-software -> hazop, jsea, whatif, followup, review_followup 

            if (dsData?.Tables.Count > 0)
            {
                ret = UpdateGeneralData(user_name, conn, transaction, dsData, seq_header_now);
                if (ret != "true") { return ret; }
            }
            if (dsData?.Tables.Count > 0)
            {
                ret = UpdateFunctionalAuditionData(user_name, conn, transaction, dsData, seq_header_now);
                if (ret != "true") { return ret; }
            }
            if (dsData?.Tables.Count > 0)
            {
                ret = UpdateSessionData(user_name, conn, transaction, dsData, seq_header_now);
                if (ret != "true") { return ret; }
            }
            if (dsData?.Tables.Count > 0)
            {
                ret = UpdateMemberTeamData(user_name, conn, transaction, dsData, seq_header_now);
                if (ret != "true") { return ret; }
            }
            if (dsData?.Tables.Count > 0)
            {
                ret = UpdateApproverData(user_name, conn, transaction, dsData, seq_header_now, flow_action);
                if (ret != "true") { return ret; }
            }
            if (dsData?.Tables.Count > 0)
            {
                ret = UpdateRelatedPeopleData(user_name, conn, transaction, dsData, seq_header_now);
                if (ret != "true") { return ret; }
            }
            if (dsData?.Tables.Count > 0)
            {
                ret = UpdateRelatedPeopleOutsiderData(user_name, conn, transaction, dsData, seq_header_now);
                if (ret != "true") { return ret; }
            }

            if (dsData?.Tables.Count > 0)
            {
                ret = UpdateDrawingData(user_name, conn, transaction, dsData, seq_header_now, _module_name);
                if (ret != "true") { return ret; }
            }

            return ret;
        }
        public string UpdateGeneralData(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet? dsData, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";
            try
            {
                DataTable dt = dsData?.Tables["general"]?.Copy() ?? new DataTable();
                if (dt?.Rows?.Count == 0) { return "true"; }

                #region update data general
                for (int i = 0; i < dt?.Rows.Count; i++)
                {
                    List<SqlParameter> parameters = new List<SqlParameter>();
                    string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                    string sqlstr = "";

                    if (action_type == "insert")
                    {
                        #region insert
                        sqlstr = "INSERT INTO EPHA_T_GENERAL (" +
                                 "SEQ, ID, ID_PHA, ID_RAM, EXPENSE_TYPE, SUB_EXPENSE_TYPE, REFERENCE_MOC, " +
                                 "ID_AREA, ID_APU, ID_BUSINESS_UNIT, ID_UNIT_NO, OTHER_AREA, OTHER_APU, OTHER_BUSINESS_UNIT, OTHER_UNIT_NO, OTHER_FUNCTIONAL_LOCATION, FUNCTIONAL_LOCATION, " +
                                 "ID_REQUEST_TYPE, PHA_REQUEST_NAME, TARGET_START_DATE, TARGET_END_DATE, ACTUAL_START_DATE, ACTUAL_END_DATE, " +
                                 "MANDATORY_NOTE, DESCRIPTIONS, WORK_SCOPE, " +
                                 "ID_DEPARTMENT, ID_DEPARTMENTS, ID_SECTIONS, ID_COMPANY, ID_TOC, ID_TAGID, INPUT_TYPE_EXCEL, TYPES_OF_HAZARD, FILE_UPLOAD_SIZE, FILE_UPLOAD_NAME, FILE_UPLOAD_PATH, " +
                                 "CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) VALUES (" +
                                 "@SEQ, @ID, @ID_PHA, @ID_RAM, @EXPENSE_TYPE, @SUB_EXPENSE_TYPE, @REFERENCE_MOC, " +
                                 "@ID_AREA, @ID_APU, @ID_BUSINESS_UNIT, @ID_UNIT_NO, @OTHER_AREA, @OTHER_APU, @OTHER_BUSINESS_UNIT, @OTHER_UNIT_NO, @OTHER_FUNCTIONAL_LOCATION, @FUNCTIONAL_LOCATION, " +
                                 "@ID_REQUEST_TYPE, @PHA_REQUEST_NAME, @TARGET_START_DATE, @TARGET_END_DATE, @ACTUAL_START_DATE, @ACTUAL_END_DATE, " +
                                 "@MANDATORY_NOTE, @DESCRIPTIONS, @WORK_SCOPE, " +
                                 "@ID_DEPARTMENT, @ID_DEPARTMENTS, @ID_SECTIONS, @ID_COMPANY, @ID_TOC, @ID_TAGID, @INPUT_TYPE_EXCEL, @TYPES_OF_HAZARD, @FILE_UPLOAD_SIZE, @FILE_UPLOAD_NAME, @FILE_UPLOAD_PATH, " +
                                 "GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_header_now) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_header_now) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_header_now) });
                        parameters.Add(new SqlParameter("@ID_RAM", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_RAM"]) });
                        parameters.Add(new SqlParameter("@EXPENSE_TYPE", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["EXPENSE_TYPE"]) });
                        parameters.Add(new SqlParameter("@SUB_EXPENSE_TYPE", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["SUB_EXPENSE_TYPE"]) });
                        parameters.Add(new SqlParameter("@REFERENCE_MOC", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["REFERENCE_MOC"]) });

                        parameters.Add(new SqlParameter("@ID_AREA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_AREA"]) });
                        parameters.Add(new SqlParameter("@ID_APU", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_APU"]) });
                        parameters.Add(new SqlParameter("@ID_BUSINESS_UNIT", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_BUSINESS_UNIT"]) });
                        parameters.Add(new SqlParameter("@ID_UNIT_NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_UNIT_NO"]) });

                        parameters.Add(new SqlParameter("@OTHER_AREA", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["OTHER_AREA"]) });
                        parameters.Add(new SqlParameter("@OTHER_APU", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["OTHER_APU"]) });
                        parameters.Add(new SqlParameter("@OTHER_BUSINESS_UNIT", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["OTHER_BUSINESS_UNIT"]) });
                        parameters.Add(new SqlParameter("@OTHER_UNIT_NO", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["OTHER_UNIT_NO"]) });
                        parameters.Add(new SqlParameter("@OTHER_FUNCTIONAL_LOCATION", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["OTHER_FUNCTIONAL_LOCATION"]) });

                        parameters.Add(new SqlParameter("@FUNCTIONAL_LOCATION", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["FUNCTIONAL_LOCATION"]) });
                        parameters.Add(new SqlParameter("@ID_REQUEST_TYPE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_REQUEST_TYPE"]) });
                        parameters.Add(new SqlParameter("@PHA_REQUEST_NAME", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["PHA_REQUEST_NAME"]) });
                        parameters.Add(new SqlParameter("@TARGET_START_DATE", SqlDbType.Date) { Value = ConvertToDBNull(dt.Rows[i]["TARGET_START_DATE"]) });
                        parameters.Add(new SqlParameter("@TARGET_END_DATE", SqlDbType.Date) { Value = ConvertToDBNull(dt.Rows[i]["TARGET_END_DATE"]) });
                        parameters.Add(new SqlParameter("@ACTUAL_START_DATE", SqlDbType.Date) { Value = ConvertToDBNull(dt.Rows[i]["ACTUAL_START_DATE"]) });
                        parameters.Add(new SqlParameter("@ACTUAL_END_DATE", SqlDbType.Date) { Value = ConvertToDBNull(dt.Rows[i]["ACTUAL_END_DATE"]) });

                        parameters.Add(new SqlParameter("@MANDATORY_NOTE", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["MANDATORY_NOTE"]) });
                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["DESCRIPTIONS"]) });
                        parameters.Add(new SqlParameter("@WORK_SCOPE", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["WORK_SCOPE"]) });

                        parameters.Add(new SqlParameter("@ID_DEPARTMENT", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_DEPARTMENT"]) });
                        parameters.Add(new SqlParameter("@ID_DEPARTMENTS", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["ID_DEPARTMENTS"]) });
                        parameters.Add(new SqlParameter("@ID_SECTIONS", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["ID_SECTIONS"]) });
                        parameters.Add(new SqlParameter("@ID_COMPANY", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_COMPANY"]) });
                        parameters.Add(new SqlParameter("@ID_TOC", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_TOC"]) });
                        parameters.Add(new SqlParameter("@ID_TAGID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_TAGID"]) });
                        parameters.Add(new SqlParameter("@INPUT_TYPE_EXCEL", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["INPUT_TYPE_EXCEL"]) });
                        parameters.Add(new SqlParameter("@TYPES_OF_HAZARD", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["TYPES_OF_HAZARD"]) });
                        parameters.Add(new SqlParameter("@FILE_UPLOAD_SIZE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["FILE_UPLOAD_SIZE"]) });
                        parameters.Add(new SqlParameter("@FILE_UPLOAD_NAME", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["FILE_UPLOAD_NAME"]) });
                        parameters.Add(new SqlParameter("@FILE_UPLOAD_PATH", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["FILE_UPLOAD_PATH"]) });

                        parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["CREATE_BY"]) });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });

                        #endregion insert
                    }
                    else if (action_type == "update")
                    {
                        #region update
                        sqlstr = "UPDATE EPHA_T_GENERAL SET " +
                                 "ID_RAM = @ID_RAM, EXPENSE_TYPE = @EXPENSE_TYPE, SUB_EXPENSE_TYPE = @SUB_EXPENSE_TYPE, REFERENCE_MOC = @REFERENCE_MOC, " +
                                 "ID_AREA = @ID_AREA, ID_APU = @ID_APU, ID_BUSINESS_UNIT = @ID_BUSINESS_UNIT, ID_UNIT_NO = @ID_UNIT_NO, " +
                                 "OTHER_AREA = @OTHER_AREA, OTHER_APU = @OTHER_APU, OTHER_BUSINESS_UNIT = @OTHER_BUSINESS_UNIT, OTHER_UNIT_NO = @OTHER_UNIT_NO, OTHER_FUNCTIONAL_LOCATION = @OTHER_FUNCTIONAL_LOCATION, " +
                                 "FUNCTIONAL_LOCATION = @FUNCTIONAL_LOCATION, ID_REQUEST_TYPE = @ID_REQUEST_TYPE, PHA_REQUEST_NAME = @PHA_REQUEST_NAME, " +
                                 "TARGET_START_DATE = @TARGET_START_DATE, TARGET_END_DATE = @TARGET_END_DATE, ACTUAL_START_DATE = @ACTUAL_START_DATE, ACTUAL_END_DATE = @ACTUAL_END_DATE, " +
                                 "MANDATORY_NOTE = @MANDATORY_NOTE, DESCRIPTIONS = @DESCRIPTIONS, WORK_SCOPE = @WORK_SCOPE, " +
                                 "ID_DEPARTMENT = @ID_DEPARTMENT, ID_DEPARTMENTS = @ID_DEPARTMENTS, ID_SECTIONS = @ID_SECTIONS, ID_COMPANY = @ID_COMPANY, ID_TOC = @ID_TOC, " +
                                 "ID_TAGID = @ID_TAGID, INPUT_TYPE_EXCEL = @INPUT_TYPE_EXCEL, TYPES_OF_HAZARD = @TYPES_OF_HAZARD, FILE_UPLOAD_SIZE = @FILE_UPLOAD_SIZE, " +
                                 "FILE_UPLOAD_NAME = @FILE_UPLOAD_NAME, FILE_UPLOAD_PATH = @FILE_UPLOAD_PATH, UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY WHERE SEQ = @SEQ AND ID = @ID";

                        parameters.Add(new SqlParameter("@ID_RAM", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_RAM"]) });
                        parameters.Add(new SqlParameter("@EXPENSE_TYPE", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["EXPENSE_TYPE"]) });
                        parameters.Add(new SqlParameter("@SUB_EXPENSE_TYPE", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["SUB_EXPENSE_TYPE"]) });
                        parameters.Add(new SqlParameter("@REFERENCE_MOC", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["REFERENCE_MOC"]) });

                        parameters.Add(new SqlParameter("@ID_AREA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_AREA"]) });
                        parameters.Add(new SqlParameter("@ID_APU", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_APU"]) });
                        parameters.Add(new SqlParameter("@ID_BUSINESS_UNIT", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_BUSINESS_UNIT"]) });
                        parameters.Add(new SqlParameter("@ID_UNIT_NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_UNIT_NO"]) });

                        parameters.Add(new SqlParameter("@OTHER_AREA", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["OTHER_AREA"]) });
                        parameters.Add(new SqlParameter("@OTHER_APU", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["OTHER_APU"]) });
                        parameters.Add(new SqlParameter("@OTHER_BUSINESS_UNIT", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["OTHER_BUSINESS_UNIT"]) });
                        parameters.Add(new SqlParameter("@OTHER_UNIT_NO", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["OTHER_UNIT_NO"]) });
                        parameters.Add(new SqlParameter("@OTHER_FUNCTIONAL_LOCATION", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["OTHER_FUNCTIONAL_LOCATION"]) });

                        parameters.Add(new SqlParameter("@FUNCTIONAL_LOCATION", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["FUNCTIONAL_LOCATION"]) });
                        parameters.Add(new SqlParameter("@ID_REQUEST_TYPE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_REQUEST_TYPE"]) });
                        parameters.Add(new SqlParameter("@PHA_REQUEST_NAME", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["PHA_REQUEST_NAME"]) });
                        parameters.Add(new SqlParameter("@TARGET_START_DATE", SqlDbType.Date) { Value = ConvertToDBNull(dt.Rows[i]["TARGET_START_DATE"]) });
                        parameters.Add(new SqlParameter("@TARGET_END_DATE", SqlDbType.Date) { Value = ConvertToDBNull(dt.Rows[i]["TARGET_END_DATE"]) });
                        parameters.Add(new SqlParameter("@ACTUAL_START_DATE", SqlDbType.Date) { Value = ConvertToDBNull(dt.Rows[i]["ACTUAL_START_DATE"]) });
                        parameters.Add(new SqlParameter("@ACTUAL_END_DATE", SqlDbType.Date) { Value = ConvertToDBNull(dt.Rows[i]["ACTUAL_END_DATE"]) });

                        parameters.Add(new SqlParameter("@MANDATORY_NOTE", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["MANDATORY_NOTE"]) });
                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["DESCRIPTIONS"]) });
                        parameters.Add(new SqlParameter("@WORK_SCOPE", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["WORK_SCOPE"]) });

                        parameters.Add(new SqlParameter("@ID_DEPARTMENT", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_DEPARTMENT"]) });
                        parameters.Add(new SqlParameter("@ID_DEPARTMENTS", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["ID_DEPARTMENTS"]) });
                        parameters.Add(new SqlParameter("@ID_SECTIONS", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["ID_SECTIONS"]) });
                        parameters.Add(new SqlParameter("@ID_COMPANY", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_COMPANY"]) });
                        parameters.Add(new SqlParameter("@ID_TOC", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_TOC"]) });
                        parameters.Add(new SqlParameter("@ID_TAGID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_TAGID"]) });
                        parameters.Add(new SqlParameter("@INPUT_TYPE_EXCEL", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["INPUT_TYPE_EXCEL"]) });
                        parameters.Add(new SqlParameter("@TYPES_OF_HAZARD", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["TYPES_OF_HAZARD"]) });
                        parameters.Add(new SqlParameter("@FILE_UPLOAD_SIZE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["FILE_UPLOAD_SIZE"]) });
                        parameters.Add(new SqlParameter("@FILE_UPLOAD_NAME", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["FILE_UPLOAD_NAME"]) });
                        parameters.Add(new SqlParameter("@FILE_UPLOAD_PATH", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["FILE_UPLOAD_PATH"]) });

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });

                        #endregion update
                    }
                    else if (action_type == "delete")
                    {
                        #region delete
                        sqlstr = "DELETE FROM EPHA_T_GENERAL WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_PHA"]) });
                        #endregion delete
                    }

                    if (action_type != "")
                    {
                        ClassConnectionDb cls_conn = new ClassConnectionDb();
                        ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                        if (ret != "true") break;
                    }
                }
                #endregion update data general

            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return ret;
        }
        public string UpdateFunctionalAuditionData(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet? dsData, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";

            try
            {
                DataTable dt = dsData?.Tables["functional_audition"]?.Copy() ?? new DataTable();
                if (dt?.Rows?.Count == 0) { return "true"; }

                for (int i = 0; i < dt?.Rows.Count; i++)
                {
                    List<SqlParameter> parameters = new List<SqlParameter>();
                    string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                    string seq_functional_audition = (dt.Rows[i]["seq"] + "").ToString();
                    string sqlstr = "";

                    if (action_type == "insert")
                    {
                        #region insert
                        sqlstr = "INSERT INTO EPHA_T_FUNCTIONAL_AUDITION (" +
                                 "SEQ, ID, ID_PHA, FUNCTIONAL_LOCATION, " +
                                 "CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                 "VALUES (@SEQ, @ID, @ID_PHA, @FUNCTIONAL_LOCATION, " +
                                 "GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_functional_audition) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_functional_audition) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_header_now) });
                        parameters.Add(new SqlParameter("@FUNCTIONAL_LOCATION", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["FUNCTIONAL_LOCATION"]) });
                        parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["CREATE_BY"]) });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });
                        #endregion insert
                    }
                    else if (action_type == "update")
                    {
                        #region update
                        sqlstr = "UPDATE EPHA_T_FUNCTIONAL_AUDITION SET " +
                                 "FUNCTIONAL_LOCATION = @FUNCTIONAL_LOCATION, " +
                                 "UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                 "WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA";

                        parameters.Add(new SqlParameter("@FUNCTIONAL_LOCATION", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["FUNCTIONAL_LOCATION"]) });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });
                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_PHA"]) });
                        #endregion update
                    }
                    else if (action_type == "delete")
                    {
                        #region delete
                        sqlstr = "DELETE FROM EPHA_T_FUNCTIONAL_AUDITION WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_PHA"]) });
                        #endregion delete
                    }

                    if (action_type != "")
                    {
                        ClassConnectionDb cls_conn = new ClassConnectionDb();
                        ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                        if (ret != "true") { break; }
                    }
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return ret;
        }
        public string UpdateSessionData(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet? dsData, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";

            try
            {
                DataTable dt = dsData?.Tables["session"]?.Copy() ?? new DataTable();
                if (dt?.Rows?.Count == 0) { return "true"; }

                for (int i = 0; i < dt?.Rows.Count; i++)
                {
                    List<SqlParameter> parameters = new List<SqlParameter>();
                    string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                    string sqlstr = "";

                    if (action_type == "insert")
                    {
                        #region insert
                        sqlstr = "INSERT INTO EPHA_T_SESSION (" +
                                 "SEQ, ID, ID_PHA, NO, MEETING_DATE, MEETING_START_TIME, MEETING_END_TIME, " +
                                 "NOTE_TO_APPROVER, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                 "VALUES (@SEQ, @ID, @ID_PHA, @NO, @MEETING_DATE, @MEETING_START_TIME, @MEETING_END_TIME, " +
                                 "@NOTE_TO_APPROVER, GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_header_now) });
                        parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["NO"]) });
                        parameters.Add(new SqlParameter("@MEETING_DATE", SqlDbType.Date) { Value = ConvertToDBNull(dt.Rows[i]["MEETING_DATE"]) });
                        parameters.Add(new SqlParameter("@MEETING_START_TIME", SqlDbType.VarChar, 100) { Value = ConvertToDBNull(dt.Rows[i]["MEETING_START_TIME"]) });
                        parameters.Add(new SqlParameter("@MEETING_END_TIME", SqlDbType.VarChar, 100) { Value = ConvertToDBNull(dt.Rows[i]["MEETING_END_TIME"]) });
                        parameters.Add(new SqlParameter("@NOTE_TO_APPROVER", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["NOTE_TO_APPROVER"]) });
                        parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["CREATE_BY"]) });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });
                        #endregion insert
                    }
                    else if (action_type == "update")
                    {
                        #region update
                        sqlstr = "UPDATE EPHA_T_SESSION SET " +
                                 "MEETING_DATE = @MEETING_DATE, MEETING_START_TIME = @MEETING_START_TIME, " +
                                 "MEETING_END_TIME = @MEETING_END_TIME, NOTE_TO_APPROVER = @NOTE_TO_APPROVER, " +
                                 "UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                 "WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA";

                        parameters.Add(new SqlParameter("@MEETING_DATE", SqlDbType.Date) { Value = ConvertToDBNull(dt.Rows[i]["MEETING_DATE"]) });
                        parameters.Add(new SqlParameter("@MEETING_START_TIME", SqlDbType.VarChar, 100) { Value = ConvertToDBNull(dt.Rows[i]["MEETING_START_TIME"]) });
                        parameters.Add(new SqlParameter("@MEETING_END_TIME", SqlDbType.VarChar, 100) { Value = ConvertToDBNull(dt.Rows[i]["MEETING_END_TIME"]) });
                        parameters.Add(new SqlParameter("@NOTE_TO_APPROVER", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["NOTE_TO_APPROVER"]) });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });
                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_PHA"]) });
                        #endregion update
                    }
                    else if (action_type == "delete")
                    {
                        #region delete
                        sqlstr = "DELETE FROM EPHA_T_SESSION WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_PHA"]) });
                        #endregion delete
                    }

                    if (action_type != "")
                    {
                        ClassConnectionDb cls_conn = new ClassConnectionDb();
                        ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                        if (ret != "true") { break; }
                    }
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return ret;
        }
        public string UpdateMemberTeamData(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet? dsData, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";

            try
            {
                DataTable dt = dsData?.Tables["memberteam"]?.Copy() ?? new DataTable();
                if (dt?.Rows?.Count == 0) { return "true"; }

                for (int i = 0; i < dt?.Rows.Count; i++)
                {
                    string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                    List<SqlParameter> parameters = new List<SqlParameter>();
                    string sqlstr = "";

                    if (action_type == "insert")
                    {
                        #region insert
                        sqlstr = "INSERT INTO EPHA_T_MEMBER_TEAM (" +
                                 "SEQ, ID, ID_SESSION, ID_PHA, NO, USER_NAME, USER_DISPLAYNAME, " +
                                 "CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                 "VALUES (@SEQ, @ID, @ID_SESSION, @ID_PHA, @NO, @USER_NAME, @USER_DISPLAYNAME, " +
                                 "GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_SESSION", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_SESSION"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_header_now) });
                        parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["NO"]) });
                        parameters.Add(new SqlParameter("@USER_NAME", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["USER_NAME"]) });
                        parameters.Add(new SqlParameter("@USER_DISPLAYNAME", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["USER_DISPLAYNAME"]) });
                        parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["CREATE_BY"]) });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });
                        #endregion insert
                    }
                    else if (action_type == "update")
                    {
                        #region update
                        sqlstr = "UPDATE EPHA_T_MEMBER_TEAM SET " +
                                 "NO = @NO, USER_NAME = @USER_NAME, USER_DISPLAYNAME = @USER_DISPLAYNAME, " +
                                 "UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                 "WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA AND ID_SESSION = @ID_SESSION";

                        parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["NO"]) });
                        parameters.Add(new SqlParameter("@USER_NAME", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["USER_NAME"]) });
                        parameters.Add(new SqlParameter("@USER_DISPLAYNAME", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["USER_DISPLAYNAME"]) });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });
                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_PHA"]) });
                        parameters.Add(new SqlParameter("@ID_SESSION", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_SESSION"]) });
                        #endregion update
                    }
                    else if (action_type == "delete")
                    {
                        #region delete
                        sqlstr = "DELETE FROM EPHA_T_MEMBER_TEAM WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA AND ID_SESSION = @ID_SESSION";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_PHA"]) });
                        parameters.Add(new SqlParameter("@ID_SESSION", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_SESSION"]) });
                        #endregion delete
                    }

                    if (action_type != "")
                    {
                        ClassConnectionDb cls_conn = new ClassConnectionDb();
                        ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                        if (ret != "true") { break; }
                    }
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return ret;
        }
        public string UpdateApproverData(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet? dsData, string seq_header_now, string flow_action)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";

            try
            {
                DataTable dt = dsData?.Tables["approver"]?.Copy() ?? new DataTable();
                if (dt?.Rows?.Count == 0) { return "true"; }

                for (int i = 0; i < dt?.Rows.Count; i++)
                {
                    string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                    List<SqlParameter> parameters = new List<SqlParameter>();
                    string sqlstr = "";

                    if (action_type == "insert")
                    {
                        #region insert
                        sqlstr = "INSERT INTO EPHA_T_APPROVER (" +
                                 "SEQ, ID, ID_SESSION, ID_PHA, NO, USER_NAME, USER_DISPLAYNAME, APPROVER_ACTION_TYPE, APPROVER_TYPE, " +
                                 "CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                 "VALUES (@SEQ, @ID, @ID_SESSION, @ID_PHA, @NO, @USER_NAME, @USER_DISPLAYNAME, @APPROVER_ACTION_TYPE, @APPROVER_TYPE, " +
                                 "GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_SESSION", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_SESSION"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_header_now) });
                        parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["NO"]) });
                        parameters.Add(new SqlParameter("@USER_NAME", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["USER_NAME"]) });
                        parameters.Add(new SqlParameter("@USER_DISPLAYNAME", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["USER_DISPLAYNAME"]) });
                        parameters.Add(new SqlParameter("@APPROVER_ACTION_TYPE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["APPROVER_ACTION_TYPE"]) });
                        parameters.Add(new SqlParameter("@APPROVER_TYPE", SqlDbType.VarChar, 100) { Value = ConvertToDBNull(dt.Rows[i]["APPROVER_TYPE"]) });
                        parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["CREATE_BY"]) });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });
                        #endregion insert
                    }
                    else if (action_type == "update")
                    {
                        #region update
                        sqlstr = "UPDATE EPHA_T_APPROVER SET " +
                                 "NO = @NO, USER_NAME = @USER_NAME, USER_DISPLAYNAME = @USER_DISPLAYNAME, ";

                        if (flow_action == "change_approver")
                        {
                            sqlstr += "APPROVER_ACTION_TYPE = NULL, APPROVER_TYPE = NULL, ACTION_REVIEW = NULL, " +
                                      "DATE_REVIEW = NULL, COMMENT = NULL, ACTION_STATUS = NULL, ";
                        }
                        else
                        {
                            sqlstr += "APPROVER_ACTION_TYPE = @APPROVER_ACTION_TYPE, APPROVER_TYPE = @APPROVER_TYPE, ";
                        }

                        sqlstr += "UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                  "WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA AND ID_SESSION = @ID_SESSION";

                        parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["NO"]) });
                        parameters.Add(new SqlParameter("@USER_NAME", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["USER_NAME"]) });
                        parameters.Add(new SqlParameter("@USER_DISPLAYNAME", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["USER_DISPLAYNAME"]) });

                        if (flow_action != "change_approver")
                        {
                            parameters.Add(new SqlParameter("@APPROVER_ACTION_TYPE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["APPROVER_ACTION_TYPE"]) });
                            parameters.Add(new SqlParameter("@APPROVER_TYPE", SqlDbType.VarChar, 100) { Value = ConvertToDBNull(dt.Rows[i]["APPROVER_TYPE"]) });
                        }

                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });
                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_PHA"]) });
                        parameters.Add(new SqlParameter("@ID_SESSION", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_SESSION"]) });
                        #endregion update
                    }
                    else if (action_type == "delete")
                    {
                        #region delete
                        sqlstr = "DELETE FROM EPHA_T_APPROVER WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA AND ID_SESSION = @ID_SESSION";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_PHA"]) });
                        parameters.Add(new SqlParameter("@ID_SESSION", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_SESSION"]) });
                        #endregion delete
                    }

                    if (action_type != "")
                    {
                        ClassConnectionDb cls_conn = new ClassConnectionDb();
                        ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                        if (ret != "true") { break; }
                    }
                }

            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return ret;
        }
        public string UpdateRelatedPeopleData(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet? dsData, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";

            try
            {
                DataTable dt = dsData?.Tables["relatedpeople"]?.Copy() ?? new DataTable();
                if (dt?.Rows?.Count == 0) { return "true"; }

                for (int i = 0; i < dt?.Rows.Count; i++)
                {
                    string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                    List<SqlParameter> parameters = new List<SqlParameter>();
                    string sqlstr = "";

                    if (action_type == "insert")
                    {
                        #region insert
                        sqlstr = "INSERT INTO EPHA_T_RELATEDPEOPLE (" +
                                 "SEQ, ID, ID_SESSION, ID_PHA, NO, USER_NAME, USER_DISPLAYNAME, USER_TITLE, USER_TYPE, APPROVER_TYPE, ACTION_STATUS, ACTION_REVIEW, DATE_REVIEW, COMMENT, " +
                                 "CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                 "VALUES (@SEQ, @ID, @ID_SESSION, @ID_PHA, @NO, @USER_NAME, @USER_DISPLAYNAME, @USER_TITLE, @USER_TYPE, @APPROVER_TYPE, @ACTION_STATUS, @ACTION_REVIEW, @DATE_REVIEW, @COMMENT, " +
                                 "GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_SESSION", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_SESSION"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_header_now) });
                        parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["NO"]) });
                        parameters.Add(new SqlParameter("@USER_NAME", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["USER_NAME"]) });
                        parameters.Add(new SqlParameter("@USER_DISPLAYNAME", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["USER_DISPLAYNAME"]) });
                        parameters.Add(new SqlParameter("@USER_TITLE", SqlDbType.VarChar, 100) { Value = ConvertToDBNull(dt.Rows[i]["USER_TITLE"]) });
                        parameters.Add(new SqlParameter("@USER_TYPE", SqlDbType.VarChar, 100) { Value = ConvertToDBNull(dt.Rows[i]["USER_TYPE"]) });
                        parameters.Add(new SqlParameter("@APPROVER_TYPE", SqlDbType.VarChar, 100) { Value = ConvertToDBNull(dt.Rows[i]["APPROVER_TYPE"]) });
                        parameters.Add(new SqlParameter("@ACTION_STATUS", SqlDbType.VarChar, 100) { Value = ConvertToDBNull(dt.Rows[i]["ACTION_STATUS"]) });
                        parameters.Add(new SqlParameter("@ACTION_REVIEW", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ACTION_REVIEW"]) });
                        parameters.Add(new SqlParameter("@DATE_REVIEW", SqlDbType.DateTime) { Value = ConvertToDateTimeOrDBNull(dt.Rows[i]["DATE_REVIEW"]) });
                        parameters.Add(new SqlParameter("@COMMENT", SqlDbType.VarChar, 100) { Value = ConvertToDBNull(dt.Rows[i]["COMMENT"]) });
                        parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["CREATE_BY"]) });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });
                        #endregion insert
                    }
                    else if (action_type == "update")
                    {
                        #region update
                        sqlstr = "UPDATE EPHA_T_RELATEDPEOPLE SET " +
                                 "NO = @NO, USER_NAME = @USER_NAME, USER_DISPLAYNAME = @USER_DISPLAYNAME, USER_TITLE = @USER_TITLE, USER_TYPE = @USER_TYPE, " +
                                 "APPROVER_TYPE = @APPROVER_TYPE, ACTION_STATUS = @ACTION_STATUS, ACTION_REVIEW = @ACTION_REVIEW, DATE_REVIEW = @DATE_REVIEW, COMMENT = @COMMENT, " +
                                 "UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                 "WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA AND ID_SESSION = @ID_SESSION";

                        parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["NO"]) });
                        parameters.Add(new SqlParameter("@USER_NAME", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["USER_NAME"]) });
                        parameters.Add(new SqlParameter("@USER_DISPLAYNAME", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["USER_DISPLAYNAME"]) });
                        parameters.Add(new SqlParameter("@USER_TITLE", SqlDbType.VarChar, 100) { Value = ConvertToDBNull(dt.Rows[i]["USER_TITLE"]) });
                        parameters.Add(new SqlParameter("@USER_TYPE", SqlDbType.VarChar, 100) { Value = ConvertToDBNull(dt.Rows[i]["USER_TYPE"]) });
                        parameters.Add(new SqlParameter("@APPROVER_TYPE", SqlDbType.VarChar, 100) { Value = ConvertToDBNull(dt.Rows[i]["APPROVER_TYPE"]) });
                        parameters.Add(new SqlParameter("@ACTION_STATUS", SqlDbType.VarChar, 100) { Value = ConvertToDBNull(dt.Rows[i]["ACTION_STATUS"]) });
                        parameters.Add(new SqlParameter("@ACTION_REVIEW", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ACTION_REVIEW"]) });
                        parameters.Add(new SqlParameter("@DATE_REVIEW", SqlDbType.DateTime) { Value = ConvertToDateTimeOrDBNull(dt.Rows[i]["DATE_REVIEW"]) });
                        parameters.Add(new SqlParameter("@COMMENT", SqlDbType.VarChar, 100) { Value = ConvertToDBNull(dt.Rows[i]["COMMENT"]) });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });
                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_PHA"]) });
                        parameters.Add(new SqlParameter("@ID_SESSION", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_SESSION"]) });
                        #endregion update
                    }
                    else if (action_type == "delete")
                    {
                        #region delete
                        sqlstr = "DELETE FROM EPHA_T_RELATEDPEOPLE WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA AND ID_SESSION = @ID_SESSION";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_PHA"]) });
                        parameters.Add(new SqlParameter("@ID_SESSION", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_SESSION"]) });
                        #endregion delete
                    }

                    if (action_type != "")
                    {
                        ClassConnectionDb cls_conn = new ClassConnectionDb();
                        ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                        if (ret != "true") { break; }
                    }
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return ret;
        }
        public string UpdateRelatedPeopleOutsiderData(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet? dsData, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";

            try
            {
                DataTable dt = dsData?.Tables["relatedpeople_outsider"]?.Copy() ?? new DataTable();
                if (dt?.Rows?.Count == 0) { return "true"; }

                for (int i = 0; i < dt?.Rows.Count; i++)
                {
                    string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                    List<SqlParameter> parameters = new List<SqlParameter>();
                    string sqlstr = "";

                    if (action_type == "insert")
                    {
                        #region insert
                        sqlstr = "INSERT INTO EPHA_T_RELATEDPEOPLE_OUTSIDER (" +
                                 "SEQ, ID, ID_SESSION, ID_PHA, NO, USER_NAME, USER_DISPLAYNAME, USER_TITLE, USER_TYPE, APPROVER_TYPE, ACTION_STATUS, ACTION_REVIEW, DATE_REVIEW, COMMENT, " +
                                 "CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                 "VALUES (@SEQ, @ID, @ID_SESSION, @ID_PHA, @NO, @USER_NAME, @USER_DISPLAYNAME, @USER_TITLE, @USER_TYPE, @APPROVER_TYPE, @ACTION_STATUS, @ACTION_REVIEW, @DATE_REVIEW, @COMMENT, " +
                                 "GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_SESSION", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_SESSION"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_header_now) });
                        parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["NO"]) });
                        parameters.Add(new SqlParameter("@USER_NAME", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["USER_NAME"]) });
                        parameters.Add(new SqlParameter("@USER_DISPLAYNAME", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["USER_DISPLAYNAME"]) });
                        parameters.Add(new SqlParameter("@USER_TITLE", SqlDbType.VarChar, 100) { Value = ConvertToDBNull(dt.Rows[i]["USER_TITLE"]) });
                        parameters.Add(new SqlParameter("@USER_TYPE", SqlDbType.VarChar, 100) { Value = ConvertToDBNull(dt.Rows[i]["USER_TYPE"]) });
                        parameters.Add(new SqlParameter("@APPROVER_TYPE", SqlDbType.VarChar, 100) { Value = ConvertToDBNull(dt.Rows[i]["APPROVER_TYPE"]) });
                        parameters.Add(new SqlParameter("@ACTION_STATUS", SqlDbType.VarChar, 100) { Value = ConvertToDBNull(dt.Rows[i]["ACTION_STATUS"]) });
                        parameters.Add(new SqlParameter("@ACTION_REVIEW", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ACTION_REVIEW"]) });
                        parameters.Add(new SqlParameter("@DATE_REVIEW", SqlDbType.DateTime) { Value = ConvertToDateTimeOrDBNull(dt.Rows[i]["DATE_REVIEW"]) });
                        parameters.Add(new SqlParameter("@COMMENT", SqlDbType.VarChar, 100) { Value = ConvertToDBNull(dt.Rows[i]["COMMENT"]) });
                        parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["CREATE_BY"]) });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });
                        #endregion insert
                    }
                    else if (action_type == "update")
                    {
                        #region update
                        sqlstr = "UPDATE EPHA_T_RELATEDPEOPLE_OUTSIDER SET " +
                                 "NO = @NO, USER_NAME = @USER_NAME, USER_DISPLAYNAME = @USER_DISPLAYNAME, USER_TITLE = @USER_TITLE, USER_TYPE = @USER_TYPE, " +
                                 "APPROVER_TYPE = @APPROVER_TYPE, ACTION_STATUS = @ACTION_STATUS, ACTION_REVIEW = @ACTION_REVIEW, DATE_REVIEW = @DATE_REVIEW, COMMENT = @COMMENT, " +
                                 "UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                 "WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA AND ID_SESSION = @ID_SESSION";

                        parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["NO"]) });
                        parameters.Add(new SqlParameter("@USER_NAME", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["USER_NAME"]) });
                        parameters.Add(new SqlParameter("@USER_DISPLAYNAME", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["USER_DISPLAYNAME"]) });
                        parameters.Add(new SqlParameter("@USER_TITLE", SqlDbType.VarChar, 100) { Value = ConvertToDBNull(dt.Rows[i]["USER_TITLE"]) });
                        parameters.Add(new SqlParameter("@USER_TYPE", SqlDbType.VarChar, 100) { Value = ConvertToDBNull(dt.Rows[i]["USER_TYPE"]) });
                        parameters.Add(new SqlParameter("@APPROVER_TYPE", SqlDbType.VarChar, 100) { Value = ConvertToDBNull(dt.Rows[i]["APPROVER_TYPE"]) });
                        parameters.Add(new SqlParameter("@ACTION_STATUS", SqlDbType.VarChar, 100) { Value = ConvertToDBNull(dt.Rows[i]["ACTION_STATUS"]) });
                        parameters.Add(new SqlParameter("@ACTION_REVIEW", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ACTION_REVIEW"]) });
                        parameters.Add(new SqlParameter("@DATE_REVIEW", SqlDbType.DateTime) { Value = ConvertToDateTimeOrDBNull(dt.Rows[i]["DATE_REVIEW"]) });
                        parameters.Add(new SqlParameter("@COMMENT", SqlDbType.VarChar, 100) { Value = ConvertToDBNull(dt.Rows[i]["COMMENT"]) });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });
                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_PHA"]) });
                        parameters.Add(new SqlParameter("@ID_SESSION", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_SESSION"]) });
                        #endregion update
                    }
                    else if (action_type == "delete")
                    {
                        #region delete
                        sqlstr = "DELETE FROM EPHA_T_RELATEDPEOPLE_OUTSIDER WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA AND ID_SESSION = @ID_SESSION";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_PHA"]) });
                        parameters.Add(new SqlParameter("@ID_SESSION", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_SESSION"]) });
                        #endregion delete
                    }

                    if (action_type != "")
                    {
                        ClassConnectionDb cls_conn = new ClassConnectionDb();
                        ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                        if (ret != "true") { break; }
                    }
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return ret;
        }
        public string UpdateDrawingData(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet? dsData, string seq_header_now, string _module_name)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";

            try
            {
                DataTable dt = dsData?.Tables["drawing"]?.Copy() ?? new DataTable();
                if (dt?.Rows?.Count == 0) { return "true"; }

                for (int i = 0; i < dt?.Rows.Count; i++)
                {
                    if (string.IsNullOrEmpty(dt.Rows[i]["DOCUMENT_MODULE"] + ""))
                    {
                        dt.Rows[i]["DOCUMENT_MODULE"] = _module_name;
                    }

                    string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                    List<SqlParameter> parameters = new List<SqlParameter>();
                    string sqlstr = "";

                    if (action_type == "insert")
                    {
                        #region insert
                        sqlstr = "INSERT INTO EPHA_T_DRAWING (" +
                                 "SEQ, ID, ID_PHA, NO, DOCUMENT_NAME, DOCUMENT_NO, DOCUMENT_FILE_NAME, DOCUMENT_FILE_PATH, DOCUMENT_FILE_SIZE, DESCRIPTIONS, DOCUMENT_MODULE, " +
                                 "CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                 "VALUES (@SEQ, @ID, @ID_PHA, @NO, @DOCUMENT_NAME, @DOCUMENT_NO, @DOCUMENT_FILE_NAME, @DOCUMENT_FILE_PATH, @DOCUMENT_FILE_SIZE, @DESCRIPTIONS, @DOCUMENT_MODULE, " +
                                 "GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_header_now) });
                        parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["NO"]) });
                        parameters.Add(new SqlParameter("@DOCUMENT_NAME", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["DOCUMENT_NAME"]) });
                        parameters.Add(new SqlParameter("@DOCUMENT_NO", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["DOCUMENT_NO"]) });
                        parameters.Add(new SqlParameter("@DOCUMENT_FILE_NAME", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["DOCUMENT_FILE_NAME"]) });
                        parameters.Add(new SqlParameter("@DOCUMENT_FILE_PATH", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["DOCUMENT_FILE_PATH"]) });
                        //parameters.Add(new SqlParameter("@DOCUMENT_FILE_SIZE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["DOCUMENT_FILE_SIZE"]) });
                        parameters.Add(new SqlParameter("@DOCUMENT_FILE_SIZE", SqlDbType.Int)
                        {
                            Value = dt.Rows[i].Table.Columns.Contains("DOCUMENT_FILE_SIZE") ? ConvertToDBNull(dt.Rows[i]["DOCUMENT_FILE_SIZE"]) : DBNull.Value
                        });
                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["DESCRIPTIONS"]) });
                        parameters.Add(new SqlParameter("@DOCUMENT_MODULE", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["DOCUMENT_MODULE"]) });
                        parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["CREATE_BY"]) });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });
                        #endregion insert
                    }
                    else if (action_type == "update")
                    {
                        #region update
                        sqlstr = "UPDATE EPHA_T_DRAWING SET " +
                                 "NO = @NO, DOCUMENT_NAME = @DOCUMENT_NAME, DOCUMENT_NO = @DOCUMENT_NO, DOCUMENT_FILE_NAME = @DOCUMENT_FILE_NAME, DOCUMENT_FILE_PATH = @DOCUMENT_FILE_PATH, " +
                                 "DOCUMENT_FILE_SIZE = @DOCUMENT_FILE_SIZE, DESCRIPTIONS = @DESCRIPTIONS, DOCUMENT_MODULE = @DOCUMENT_MODULE, " +
                                 "UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                 "WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA";

                        parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["NO"]) });
                        parameters.Add(new SqlParameter("@DOCUMENT_NAME", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["DOCUMENT_NAME"]) });
                        parameters.Add(new SqlParameter("@DOCUMENT_NO", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["DOCUMENT_NO"]) });
                        parameters.Add(new SqlParameter("@DOCUMENT_FILE_NAME", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["DOCUMENT_FILE_NAME"]) });
                        parameters.Add(new SqlParameter("@DOCUMENT_FILE_PATH", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["DOCUMENT_FILE_PATH"]) });
                        //parameters.Add(new SqlParameter("@DOCUMENT_FILE_SIZE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["DOCUMENT_FILE_SIZE"]) });
                        parameters.Add(new SqlParameter("@DOCUMENT_FILE_SIZE", SqlDbType.Int)
                        {
                            Value = dt.Rows[i].Table.Columns.Contains("DOCUMENT_FILE_SIZE") ? ConvertToDBNull(dt.Rows[i]["DOCUMENT_FILE_SIZE"]) : DBNull.Value
                        });
                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["DESCRIPTIONS"]) });
                        parameters.Add(new SqlParameter("@DOCUMENT_MODULE", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["DOCUMENT_MODULE"]) });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });
                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_PHA"]) });
                        #endregion update
                    }
                    else if (action_type == "delete")
                    {
                        #region delete
                        sqlstr = "DELETE FROM EPHA_T_DRAWING WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_PHA"]) });
                        #endregion delete
                    }

                    if (action_type != "")
                    {
                        ClassConnectionDb cls_conn = new ClassConnectionDb();
                        ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                        if (ret != "true") { break; }
                    }
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return ret;
        }
        public string set_hazop_partii(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet? dsData, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            List<SqlParameter> parameters = new List<SqlParameter>();
            string ret = "true";
            if (dsData?.Tables.Count > 0)
            {
                ret = UpdateNodeData(user_name, conn, transaction, dsData, seq_header_now);
                if (ret != "true") { return ret; }
            }
            if (dsData?.Tables.Count > 0)
            {
                ret = UpdateNodeDrawingData(user_name, conn, transaction, dsData, seq_header_now);
                if (ret != "true") { return ret; }
            }
            if (dsData?.Tables.Count > 0)
            {
                ret = UpdateNodeGuideWordsData(user_name, conn, transaction, dsData, seq_header_now);
                if (ret != "true") { return ret; }
            }

            return ret;

        }
        public string UpdateNodeData(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet? dsData, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";

            try
            {
                DataTable dt = dsData?.Tables["node"]?.Copy() ?? new DataTable();
                if (dt?.Rows?.Count == 0) { return "true"; }

                for (int i = 0; i < dt?.Rows.Count; i++)
                {
                    string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                    List<SqlParameter> parameters = new List<SqlParameter>();
                    string sqlstr = "";

                    if (action_type == "insert")
                    {
                        #region insert
                        sqlstr = "INSERT INTO EPHA_T_NODE (" +
                                 "SEQ, ID, ID_PHA, NO, NODE, DESIGN_INTENT, DESIGN_CONDITIONS, OPERATING_CONDITIONS, NODE_BOUNDARY, DESCRIPTIONS, " +
                                 "CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                 "VALUES (@SEQ, @ID, @ID_PHA, @NO, @NODE, @DESIGN_INTENT, @DESIGN_CONDITIONS, @OPERATING_CONDITIONS, @NODE_BOUNDARY, @DESCRIPTIONS, " +
                                 "GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_header_now) });
                        parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["NO"]) });
                        parameters.Add(new SqlParameter("@NODE", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["NODE"]) });
                        parameters.Add(new SqlParameter("@DESIGN_INTENT", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["DESIGN_INTENT"]) });
                        parameters.Add(new SqlParameter("@DESIGN_CONDITIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["DESIGN_CONDITIONS"]) });
                        parameters.Add(new SqlParameter("@OPERATING_CONDITIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["OPERATING_CONDITIONS"]) });
                        parameters.Add(new SqlParameter("@NODE_BOUNDARY", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["NODE_BOUNDARY"]) });
                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["DESCRIPTIONS"]) });
                        parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["CREATE_BY"]) });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });
                        #endregion insert
                    }
                    else if (action_type == "update")
                    {
                        #region update
                        sqlstr = "UPDATE EPHA_T_NODE SET " +
                                 "NO = @NO, NODE = @NODE, DESIGN_INTENT = @DESIGN_INTENT, DESIGN_CONDITIONS = @DESIGN_CONDITIONS, OPERATING_CONDITIONS = @OPERATING_CONDITIONS, " +
                                 "NODE_BOUNDARY = @NODE_BOUNDARY, DESCRIPTIONS = @DESCRIPTIONS, " +
                                 "UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                 "WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA";

                        parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["NO"]) });
                        parameters.Add(new SqlParameter("@NODE", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["NODE"]) });
                        parameters.Add(new SqlParameter("@DESIGN_INTENT", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["DESIGN_INTENT"]) });
                        parameters.Add(new SqlParameter("@DESIGN_CONDITIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["DESIGN_CONDITIONS"]) });
                        parameters.Add(new SqlParameter("@OPERATING_CONDITIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["OPERATING_CONDITIONS"]) });
                        parameters.Add(new SqlParameter("@NODE_BOUNDARY", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["NODE_BOUNDARY"]) });
                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["DESCRIPTIONS"]) });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });
                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_PHA"]) });
                        #endregion update
                    }
                    else if (action_type == "delete")
                    {
                        #region delete
                        sqlstr = "DELETE FROM EPHA_T_NODE WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_PHA"]) });
                        #endregion delete
                    }

                    if (action_type != "")
                    {
                        ClassConnectionDb cls_conn = new ClassConnectionDb();
                        ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                        if (ret != "true") { break; }
                    }
                }

            }
            catch (Exception ex)
            {
                transaction.Rollback();
                ret = "error: " + ex.Message;
            }

            return ret;
        }
        public string UpdateNodeDrawingData(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet? dsData, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";

            try
            {
                DataTable dt = dsData?.Tables["nodedrawing"]?.Copy() ?? new DataTable();
                if (dt?.Rows?.Count == 0) { return "true"; }

                for (int i = 0; i < dt?.Rows.Count; i++)
                {
                    string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                    List<SqlParameter> parameters = new List<SqlParameter>();
                    string sqlstr = "";

                    if (action_type == "insert")
                    {
                        #region insert
                        sqlstr = "INSERT INTO EPHA_T_NODE_DRAWING (" +
                                 "SEQ, ID, ID_PHA, ID_NODE, ID_DRAWING, NO, PAGE_START_FIRST, PAGE_END_FIRST, PAGE_START_SECOND, PAGE_END_SECOND, " +
                                 "PAGE_START_THIRD, PAGE_END_THIRD, DESCRIPTIONS, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                 "VALUES (@SEQ, @ID, @ID_PHA, @ID_NODE, @ID_DRAWING, @NO, @PAGE_START_FIRST, @PAGE_END_FIRST, @PAGE_START_SECOND, " +
                                 "@PAGE_END_SECOND, @PAGE_START_THIRD, @PAGE_END_THIRD, @DESCRIPTIONS, GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_header_now) });
                        parameters.Add(new SqlParameter("@ID_NODE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_NODE"]) });
                        parameters.Add(new SqlParameter("@ID_DRAWING", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_DRAWING"]) });
                        parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["NO"]) });
                        parameters.Add(new SqlParameter("@PAGE_START_FIRST", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["PAGE_START_FIRST"]) });
                        parameters.Add(new SqlParameter("@PAGE_END_FIRST", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["PAGE_END_FIRST"]) });
                        parameters.Add(new SqlParameter("@PAGE_START_SECOND", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["PAGE_START_SECOND"]) });
                        parameters.Add(new SqlParameter("@PAGE_END_SECOND", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["PAGE_END_SECOND"]) });
                        parameters.Add(new SqlParameter("@PAGE_START_THIRD", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["PAGE_START_THIRD"]) });
                        parameters.Add(new SqlParameter("@PAGE_END_THIRD", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["PAGE_END_THIRD"]) });
                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["DESCRIPTIONS"]) });
                        parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["CREATE_BY"]) });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });
                        #endregion insert
                    }
                    else if (action_type == "update")
                    {
                        #region update
                        sqlstr = "UPDATE EPHA_T_NODE_DRAWING SET " +
                                 "NO = @NO, ID_DRAWING = @ID_DRAWING, PAGE_START_FIRST = @PAGE_START_FIRST, PAGE_END_FIRST = @PAGE_END_FIRST, " +
                                 "PAGE_START_SECOND = @PAGE_START_SECOND, PAGE_END_SECOND = @PAGE_END_SECOND, PAGE_START_THIRD = @PAGE_START_THIRD, " +
                                 "PAGE_END_THIRD = @PAGE_END_THIRD, DESCRIPTIONS = @DESCRIPTIONS, UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                 "WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA AND ID_NODE = @ID_NODE";

                        parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["NO"]) });
                        parameters.Add(new SqlParameter("@ID_DRAWING", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_DRAWING"]) });
                        parameters.Add(new SqlParameter("@PAGE_START_FIRST", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["PAGE_START_FIRST"]) });
                        parameters.Add(new SqlParameter("@PAGE_END_FIRST", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["PAGE_END_FIRST"]) });
                        parameters.Add(new SqlParameter("@PAGE_START_SECOND", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["PAGE_START_SECOND"]) });
                        parameters.Add(new SqlParameter("@PAGE_END_SECOND", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["PAGE_END_SECOND"]) });
                        parameters.Add(new SqlParameter("@PAGE_START_THIRD", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["PAGE_START_THIRD"]) });
                        parameters.Add(new SqlParameter("@PAGE_END_THIRD", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["PAGE_END_THIRD"]) });
                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["DESCRIPTIONS"]) });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });
                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_PHA"]) });
                        parameters.Add(new SqlParameter("@ID_NODE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_NODE"]) });
                        #endregion update
                    }
                    else if (action_type == "delete")
                    {
                        #region delete
                        sqlstr = "DELETE FROM EPHA_T_NODE_DRAWING WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA AND ID_NODE = @ID_NODE";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_PHA"]) });
                        parameters.Add(new SqlParameter("@ID_NODE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_NODE"]) });
                        #endregion delete
                    }

                    if (action_type != "")
                    {
                        ClassConnectionDb cls_conn = new ClassConnectionDb();
                        ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                        if (ret != "true") { break; }
                    }
                }

            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return ret;
        }
        public string UpdateNodeGuideWordsData(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet? dsData, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";

            try
            {
                DataTable dt = dsData?.Tables["nodeguidwords"]?.Copy() ?? new DataTable();
                if (dt?.Rows?.Count == 0) { return "true"; }

                for (int i = 0; i < dt?.Rows.Count; i++)
                {
                    string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                    List<SqlParameter> parameters = new List<SqlParameter>();
                    string sqlstr = "";

                    if (action_type == "insert")
                    {
                        #region insert
                        sqlstr = "INSERT INTO EPHA_T_NODE_GUIDE_WORDS (" +
                                 "SEQ, ID, ID_PHA, ID_NODE, ID_GUIDE_WORD, NO, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                 "VALUES (@SEQ, @ID, @ID_PHA, @ID_NODE, @ID_GUIDE_WORD, @NO, GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_header_now) });
                        parameters.Add(new SqlParameter("@ID_NODE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_NODE"]) });
                        parameters.Add(new SqlParameter("@ID_GUIDE_WORD", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_GUIDE_WORD"]) });
                        parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["NO"]) });
                        parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["CREATE_BY"]) });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });
                        #endregion insert
                    }
                    else if (action_type == "update")
                    {
                        #region update
                        sqlstr = "UPDATE EPHA_T_NODE_GUIDE_WORDS SET " +
                                 "ID_GUIDE_WORD = @ID_GUIDE_WORD, NO = @NO, UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                 "WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA AND ID_NODE = @ID_NODE";

                        parameters.Add(new SqlParameter("@ID_GUIDE_WORD", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_GUIDE_WORD"]) });
                        parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["NO"]) });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });
                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_PHA"]) });
                        parameters.Add(new SqlParameter("@ID_NODE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_NODE"]) });
                        #endregion update
                    }
                    else if (action_type == "delete")
                    {
                        #region delete
                        sqlstr = "DELETE FROM EPHA_T_NODE_GUIDE_WORDS WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA AND ID_NODE = @ID_NODE";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_PHA"]) });
                        parameters.Add(new SqlParameter("@ID_NODE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_NODE"]) });
                        #endregion delete
                    }

                    if (action_type != "")
                    {
                        ClassConnectionDb cls_conn = new ClassConnectionDb();
                        ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                        if (ret != "true") { break; }
                    }
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return ret;
        }
        public string set_hazop_partiii(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet? dsData, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";
            List<SqlParameter> parameters = new List<SqlParameter>();
            if (dsData?.Tables.Count > 0)
            {
                ret = UpdateNodeWorksheetData(user_name, conn, transaction, dsData, seq_header_now);
                if (ret != "true") { return ret; }
            }

            return ret;

        }
        public string UpdateNodeWorksheetData(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet? dsData, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";

            try
            {
                DataTable dt = dsData?.Tables["nodeworksheet"]?.Copy() ?? new DataTable();
                if (dt?.Rows?.Count == 0) { return "true"; }

                for (int i = 0; i < dt?.Rows.Count; i++)
                {
                    string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                    List<SqlParameter> parameters = new List<SqlParameter>();
                    string sqlstr = "";

                    if (!string.IsNullOrEmpty(action_type))
                    {
                        if (action_type == "insert")
                        {
                            #region insert
                            sqlstr = "INSERT INTO EPHA_T_NODE_WORKSHEET (" +
                                     "SEQ, ID, ID_PHA, ROW_TYPE, ID_NODE, ID_GUIDE_WORD, SEQ_GUIDE_WORD, SEQ_CAUSES, SEQ_CONSEQUENCES, SEQ_CATEGORY, NO, CAUSES_NO, CAUSES, CONSEQUENCES_NO, CONSEQUENCES, CATEGORY_NO, CATEGORY_TYPE, RAM_BEFOR_SECURITY, RAM_BEFOR_LIKELIHOOD, RAM_BEFOR_RISK, MAJOR_ACCIDENT_EVENT, SAFETY_CRITICAL_EQUIPMENT, SAFETY_CRITICAL_EQUIPMENT_TAG, EXISTING_SAFEGUARDS, RAM_AFTER_SECURITY, RAM_AFTER_LIKELIHOOD, RAM_AFTER_RISK, FK_RECOMMENDATIONS, SEQ_RECOMMENDATIONS, RECOMMENDATIONS, RECOMMENDATIONS_NO, RESPONDER_USER_NAME, RESPONDER_USER_DISPLAYNAME, RAM_ACTION_SECURITY, RAM_ACTION_LIKELIHOOD, RAM_ACTION_RISK, ESTIMATED_START_DATE, ESTIMATED_END_DATE, ACTION_STATUS, IMPLEMENT, RESPONDER_ACTION_TYPE, RESPONDER_ACTION_DATE, REVIEWER_ACTION_TYPE, REVIEWER_ACTION_DATE, ACTION_PROJECT_TEAM, PROJECT_TEAM_TEXT, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                     "VALUES (@SEQ, @ID, @ID_PHA, @ROW_TYPE, @ID_NODE, @ID_GUIDE_WORD, @SEQ_GUIDE_WORD, @SEQ_CAUSES, @SEQ_CONSEQUENCES, @SEQ_CATEGORY, @NO, @CAUSES_NO, @CAUSES, @CONSEQUENCES_NO, @CONSEQUENCES, @CATEGORY_NO, @CATEGORY_TYPE, @RAM_BEFOR_SECURITY, @RAM_BEFOR_LIKELIHOOD, @RAM_BEFOR_RISK, @MAJOR_ACCIDENT_EVENT, @SAFETY_CRITICAL_EQUIPMENT, @SAFETY_CRITICAL_EQUIPMENT_TAG, @EXISTING_SAFEGUARDS, @RAM_AFTER_SECURITY, @RAM_AFTER_LIKELIHOOD, @RAM_AFTER_RISK, @FK_RECOMMENDATIONS, @SEQ_RECOMMENDATIONS, @RECOMMENDATIONS, @RECOMMENDATIONS_NO, @RESPONDER_USER_NAME, @RESPONDER_USER_DISPLAYNAME, @RAM_ACTION_SECURITY, @RAM_ACTION_LIKELIHOOD, @RAM_ACTION_RISK, @ESTIMATED_START_DATE, @ESTIMATED_END_DATE, @ACTION_STATUS, @IMPLEMENT, @RESPONDER_ACTION_TYPE, @RESPONDER_ACTION_DATE, @REVIEWER_ACTION_TYPE, @REVIEWER_ACTION_DATE, @ACTION_PROJECT_TEAM, @PROJECT_TEAM_TEXT, GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                            parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_header_now) });
                            parameters.Add(new SqlParameter("@ROW_TYPE", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["ROW_TYPE"]) });
                            parameters.Add(new SqlParameter("@ID_NODE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_NODE"]) });
                            parameters.Add(new SqlParameter("@ID_GUIDE_WORD", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_GUIDE_WORD"]) });
                            parameters.Add(new SqlParameter("@SEQ_GUIDE_WORD", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ_GUIDE_WORD"]) });
                            parameters.Add(new SqlParameter("@SEQ_CAUSES", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ_CAUSES"]) });
                            parameters.Add(new SqlParameter("@SEQ_CONSEQUENCES", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ_CONSEQUENCES"]) });
                            parameters.Add(new SqlParameter("@SEQ_CATEGORY", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ_CATEGORY"]) });
                            parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["NO"]) });
                            parameters.Add(new SqlParameter("@CAUSES_NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["CAUSES_NO"]) });
                            parameters.Add(new SqlParameter("@CAUSES", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["CAUSES"]) });
                            parameters.Add(new SqlParameter("@CONSEQUENCES_NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["CONSEQUENCES_NO"]) });
                            parameters.Add(new SqlParameter("@CONSEQUENCES", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["CONSEQUENCES"]) });
                            parameters.Add(new SqlParameter("@CATEGORY_NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["CATEGORY_NO"]) });
                            parameters.Add(new SqlParameter("@CATEGORY_TYPE", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["CATEGORY_TYPE"]) });
                            parameters.Add(new SqlParameter("@RAM_BEFOR_SECURITY", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["RAM_BEFOR_SECURITY"]) });
                            parameters.Add(new SqlParameter("@RAM_BEFOR_LIKELIHOOD", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["RAM_BEFOR_LIKELIHOOD"]) });
                            parameters.Add(new SqlParameter("@RAM_BEFOR_RISK", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["RAM_BEFOR_RISK"]) });
                            parameters.Add(new SqlParameter("@MAJOR_ACCIDENT_EVENT", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["MAJOR_ACCIDENT_EVENT"]) });
                            parameters.Add(new SqlParameter("@SAFETY_CRITICAL_EQUIPMENT", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["SAFETY_CRITICAL_EQUIPMENT"]) });
                            parameters.Add(new SqlParameter("@SAFETY_CRITICAL_EQUIPMENT_TAG", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["SAFETY_CRITICAL_EQUIPMENT_TAG"]) });
                            parameters.Add(new SqlParameter("@EXISTING_SAFEGUARDS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["EXISTING_SAFEGUARDS"]) });
                            parameters.Add(new SqlParameter("@RAM_AFTER_SECURITY", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["RAM_AFTER_SECURITY"]) });
                            parameters.Add(new SqlParameter("@RAM_AFTER_LIKELIHOOD", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["RAM_AFTER_LIKELIHOOD"]) });
                            parameters.Add(new SqlParameter("@RAM_AFTER_RISK", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["RAM_AFTER_RISK"]) });
                            parameters.Add(new SqlParameter("@FK_RECOMMENDATIONS", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["FK_RECOMMENDATIONS"]) });
                            parameters.Add(new SqlParameter("@SEQ_RECOMMENDATIONS", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ_RECOMMENDATIONS"]) });
                            parameters.Add(new SqlParameter("@RECOMMENDATIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["RECOMMENDATIONS"]) });
                            parameters.Add(new SqlParameter("@RECOMMENDATIONS_NO", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["RECOMMENDATIONS_NO"]) });
                            parameters.Add(new SqlParameter("@RESPONDER_USER_NAME", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["RESPONDER_USER_NAME"]) });
                            parameters.Add(new SqlParameter("@RESPONDER_USER_DISPLAYNAME", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["RESPONDER_USER_DISPLAYNAME"]) });
                            parameters.Add(new SqlParameter("@RAM_ACTION_SECURITY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["RAM_ACTION_SECURITY"]) });
                            parameters.Add(new SqlParameter("@RAM_ACTION_LIKELIHOOD", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["RAM_ACTION_LIKELIHOOD"]) });
                            parameters.Add(new SqlParameter("@RAM_ACTION_RISK", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["RAM_ACTION_RISK"]) });
                            parameters.Add(new SqlParameter("@ESTIMATED_START_DATE", SqlDbType.Date) { Value = ConvertToDBNull(dt.Rows[i]["ESTIMATED_START_DATE"]) });
                            parameters.Add(new SqlParameter("@ESTIMATED_END_DATE", SqlDbType.Date) { Value = ConvertToDBNull(dt.Rows[i]["ESTIMATED_END_DATE"]) });
                            parameters.Add(new SqlParameter("@ACTION_STATUS", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["ACTION_STATUS"]) });
                            parameters.Add(new SqlParameter("@IMPLEMENT", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["IMPLEMENT"]) });
                            parameters.Add(new SqlParameter("@RESPONDER_ACTION_TYPE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["RESPONDER_ACTION_TYPE"]) });
                            parameters.Add(new SqlParameter("@RESPONDER_ACTION_DATE", SqlDbType.Date) { Value = ConvertToDBNull(dt.Rows[i]["RESPONDER_ACTION_DATE"]) });
                            parameters.Add(new SqlParameter("@REVIEWER_ACTION_TYPE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["REVIEWER_ACTION_TYPE"]) });
                            parameters.Add(new SqlParameter("@REVIEWER_ACTION_DATE", SqlDbType.Date) { Value = ConvertToDBNull(dt.Rows[i]["REVIEWER_ACTION_DATE"]) });
                            parameters.Add(new SqlParameter("@ACTION_PROJECT_TEAM", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ACTION_PROJECT_TEAM"]) });
                            parameters.Add(new SqlParameter("@PROJECT_TEAM_TEXT", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["PROJECT_TEAM_TEXT"]) });
                            parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["CREATE_BY"]) });
                            parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });
                            #endregion insert
                        }
                        else if (action_type == "update")
                        {
                            #region update
                            sqlstr = "UPDATE EPHA_T_NODE_WORKSHEET SET " +
                                     "ROW_TYPE = @ROW_TYPE, ID_NODE = @ID_NODE, ID_GUIDE_WORD = @ID_GUIDE_WORD, SEQ_GUIDE_WORD = @SEQ_GUIDE_WORD, SEQ_CAUSES = @SEQ_CAUSES, SEQ_CONSEQUENCES = @SEQ_CONSEQUENCES, SEQ_CATEGORY = @SEQ_CATEGORY, NO = @NO, CAUSES_NO = @CAUSES_NO, CAUSES = @CAUSES, CONSEQUENCES_NO = @CONSEQUENCES_NO, CONSEQUENCES = @CONSEQUENCES, CATEGORY_NO = @CATEGORY_NO, CATEGORY_TYPE = @CATEGORY_TYPE, RAM_BEFOR_SECURITY = @RAM_BEFOR_SECURITY, RAM_BEFOR_LIKELIHOOD = @RAM_BEFOR_LIKELIHOOD, RAM_BEFOR_RISK = @RAM_BEFOR_RISK, MAJOR_ACCIDENT_EVENT = @MAJOR_ACCIDENT_EVENT, SAFETY_CRITICAL_EQUIPMENT = @SAFETY_CRITICAL_EQUIPMENT, SAFETY_CRITICAL_EQUIPMENT_TAG = @SAFETY_CRITICAL_EQUIPMENT_TAG, EXISTING_SAFEGUARDS = @EXISTING_SAFEGUARDS, RAM_AFTER_SECURITY = @RAM_AFTER_SECURITY, RAM_AFTER_LIKELIHOOD = @RAM_AFTER_LIKELIHOOD, RAM_AFTER_RISK = @RAM_AFTER_RISK, FK_RECOMMENDATIONS = @FK_RECOMMENDATIONS, SEQ_RECOMMENDATIONS = @SEQ_RECOMMENDATIONS, RECOMMENDATIONS = @RECOMMENDATIONS, RECOMMENDATIONS_NO = @RECOMMENDATIONS_NO, RESPONDER_USER_NAME = @RESPONDER_USER_NAME, RESPONDER_USER_DISPLAYNAME = @RESPONDER_USER_DISPLAYNAME, RAM_ACTION_SECURITY = @RAM_ACTION_SECURITY, RAM_ACTION_LIKELIHOOD = @RAM_ACTION_LIKELIHOOD, RAM_ACTION_RISK = @RAM_ACTION_RISK, ESTIMATED_START_DATE = @ESTIMATED_START_DATE, ESTIMATED_END_DATE = @ESTIMATED_END_DATE, ACTION_STATUS = @ACTION_STATUS, IMPLEMENT = @IMPLEMENT, RESPONDER_ACTION_TYPE = @RESPONDER_ACTION_TYPE, RESPONDER_ACTION_DATE = @RESPONDER_ACTION_DATE, REVIEWER_ACTION_TYPE = @REVIEWER_ACTION_TYPE, REVIEWER_ACTION_DATE = @REVIEWER_ACTION_DATE, ACTION_PROJECT_TEAM = @ACTION_PROJECT_TEAM, PROJECT_TEAM_TEXT = @PROJECT_TEAM_TEXT, UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                     "WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA AND ID_NODE = @ID_NODE AND ID_GUIDE_WORD = @ID_GUIDE_WORD";

                            parameters.Add(new SqlParameter("@ROW_TYPE", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["ROW_TYPE"]) });
                            parameters.Add(new SqlParameter("@ID_NODE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_NODE"]) });
                            parameters.Add(new SqlParameter("@ID_GUIDE_WORD", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_GUIDE_WORD"]) });
                            parameters.Add(new SqlParameter("@SEQ_GUIDE_WORD", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ_GUIDE_WORD"]) });
                            parameters.Add(new SqlParameter("@SEQ_CAUSES", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ_CAUSES"]) });
                            parameters.Add(new SqlParameter("@SEQ_CONSEQUENCES", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ_CONSEQUENCES"]) });
                            parameters.Add(new SqlParameter("@SEQ_CATEGORY", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ_CATEGORY"]) });
                            parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["NO"]) });
                            parameters.Add(new SqlParameter("@CAUSES_NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["CAUSES_NO"]) });
                            parameters.Add(new SqlParameter("@CAUSES", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["CAUSES"]) });
                            parameters.Add(new SqlParameter("@CONSEQUENCES_NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["CONSEQUENCES_NO"]) });
                            parameters.Add(new SqlParameter("@CONSEQUENCES", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["CONSEQUENCES"]) });
                            parameters.Add(new SqlParameter("@CATEGORY_NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["CATEGORY_NO"]) });
                            parameters.Add(new SqlParameter("@CATEGORY_TYPE", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["CATEGORY_TYPE"]) });
                            parameters.Add(new SqlParameter("@RAM_BEFOR_SECURITY", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["RAM_BEFOR_SECURITY"]) });
                            parameters.Add(new SqlParameter("@RAM_BEFOR_LIKELIHOOD", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["RAM_BEFOR_LIKELIHOOD"]) });
                            parameters.Add(new SqlParameter("@RAM_BEFOR_RISK", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["RAM_BEFOR_RISK"]) });
                            parameters.Add(new SqlParameter("@MAJOR_ACCIDENT_EVENT", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["MAJOR_ACCIDENT_EVENT"]) });
                            parameters.Add(new SqlParameter("@SAFETY_CRITICAL_EQUIPMENT", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["SAFETY_CRITICAL_EQUIPMENT"]) });
                            parameters.Add(new SqlParameter("@SAFETY_CRITICAL_EQUIPMENT_TAG", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["SAFETY_CRITICAL_EQUIPMENT_TAG"]) });
                            parameters.Add(new SqlParameter("@EXISTING_SAFEGUARDS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["EXISTING_SAFEGUARDS"]) });
                            parameters.Add(new SqlParameter("@RAM_AFTER_SECURITY", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["RAM_AFTER_SECURITY"]) });
                            parameters.Add(new SqlParameter("@RAM_AFTER_LIKELIHOOD", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["RAM_AFTER_LIKELIHOOD"]) });
                            parameters.Add(new SqlParameter("@RAM_AFTER_RISK", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["RAM_AFTER_RISK"]) });
                            parameters.Add(new SqlParameter("@FK_RECOMMENDATIONS", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["FK_RECOMMENDATIONS"]) });
                            parameters.Add(new SqlParameter("@SEQ_RECOMMENDATIONS", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ_RECOMMENDATIONS"]) });
                            parameters.Add(new SqlParameter("@RECOMMENDATIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["RECOMMENDATIONS"]) });
                            parameters.Add(new SqlParameter("@RECOMMENDATIONS_NO", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["RECOMMENDATIONS_NO"]) });
                            parameters.Add(new SqlParameter("@RESPONDER_USER_NAME", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["RESPONDER_USER_NAME"]) });
                            parameters.Add(new SqlParameter("@RESPONDER_USER_DISPLAYNAME", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["RESPONDER_USER_DISPLAYNAME"]) });
                            parameters.Add(new SqlParameter("@RAM_ACTION_SECURITY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["RAM_ACTION_SECURITY"]) });
                            parameters.Add(new SqlParameter("@RAM_ACTION_LIKELIHOOD", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["RAM_ACTION_LIKELIHOOD"]) });
                            parameters.Add(new SqlParameter("@RAM_ACTION_RISK", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["RAM_ACTION_RISK"]) });
                            parameters.Add(new SqlParameter("@ESTIMATED_START_DATE", SqlDbType.Date) { Value = ConvertToDBNull(dt.Rows[i]["ESTIMATED_START_DATE"]) });
                            parameters.Add(new SqlParameter("@ESTIMATED_END_DATE", SqlDbType.Date) { Value = ConvertToDBNull(dt.Rows[i]["ESTIMATED_END_DATE"]) });
                            parameters.Add(new SqlParameter("@ACTION_STATUS", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["ACTION_STATUS"]) });
                            parameters.Add(new SqlParameter("@IMPLEMENT", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["IMPLEMENT"]) });
                            parameters.Add(new SqlParameter("@RESPONDER_ACTION_TYPE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["RESPONDER_ACTION_TYPE"]) });
                            parameters.Add(new SqlParameter("@RESPONDER_ACTION_DATE", SqlDbType.Date) { Value = ConvertToDBNull(dt.Rows[i]["RESPONDER_ACTION_DATE"]) });
                            parameters.Add(new SqlParameter("@REVIEWER_ACTION_TYPE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["REVIEWER_ACTION_TYPE"]) });
                            parameters.Add(new SqlParameter("@REVIEWER_ACTION_DATE", SqlDbType.Date) { Value = ConvertToDBNull(dt.Rows[i]["REVIEWER_ACTION_DATE"]) });
                            parameters.Add(new SqlParameter("@ACTION_PROJECT_TEAM", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ACTION_PROJECT_TEAM"]) });
                            parameters.Add(new SqlParameter("@PROJECT_TEAM_TEXT", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["PROJECT_TEAM_TEXT"]) });
                            parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });
                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                            parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_PHA"]) });
                            parameters.Add(new SqlParameter("@ID_NODE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_NODE"]) });
                            parameters.Add(new SqlParameter("@ID_GUIDE_WORD", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_GUIDE_WORD"]) });
                            #endregion update
                        }
                        else if (action_type == "delete")
                        {
                            #region delete
                            sqlstr = "DELETE FROM EPHA_T_NODE_WORKSHEET WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA AND ID_NODE = @ID_NODE AND ID_GUIDE_WORD = @ID_GUIDE_WORD";

                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                            parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_PHA"]) });
                            parameters.Add(new SqlParameter("@ID_NODE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_NODE"]) });
                            parameters.Add(new SqlParameter("@ID_GUIDE_WORD", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_GUIDE_WORD"]) });
                            #endregion delete
                        }

                        if (!string.IsNullOrEmpty(sqlstr))
                        {
                            ClassConnectionDb cls_conn = new ClassConnectionDb();
                            ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                            if (ret != "true") { break; }
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

        public string set_jsea_partii(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet? dsData, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";
            List<SqlParameter> parameters = new List<SqlParameter>();
            if (dsData?.Tables.Count > 0)
            {
                ret = UpdateListData(user_name, conn, transaction, dsData, seq_header_now);
                if (ret != "true") { return ret; }
            }
            if (dsData?.Tables.Count > 0)
            {
                ret = UpdateListDrawingData(user_name, conn, transaction, dsData, seq_header_now);
                if (ret != "true") { return ret; }
            }
            return ret;
        }
        public string UpdateListData(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet? dsData, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";

            try
            {
                DataTable dt = dsData?.Tables["list"]?.Copy() ?? new DataTable();
                if (dt?.Rows?.Count == 0) { return "true"; }

                for (int i = 0; i < dt?.Rows.Count; i++)
                {
                    List<SqlParameter> parameters = new List<SqlParameter>();
                    string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                    string sqlstr = "";

                    if (action_type == "insert")
                    {
                        #region insert
                        sqlstr = "INSERT INTO EPHA_T_LIST (" +
                                 "SEQ, ID, ID_PHA, NO, LIST, DESIGN_INTENT, DESIGN_CONDITIONS, OPERATING_CONDITIONS, LIST_BOUNDARY, DESCRIPTIONS, " +
                                 "CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                 "VALUES (@SEQ, @ID, @ID_PHA, @NO, @LIST, @DESIGN_INTENT, @DESIGN_CONDITIONS, @OPERATING_CONDITIONS, @LIST_BOUNDARY, @DESCRIPTIONS, " +
                                 "GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_header_now) });
                        parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["NO"]) });
                        parameters.Add(new SqlParameter("@LIST", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["LIST"]) });
                        parameters.Add(new SqlParameter("@DESIGN_INTENT", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["DESIGN_INTENT"]) });
                        parameters.Add(new SqlParameter("@DESIGN_CONDITIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["DESIGN_CONDITIONS"]) });
                        parameters.Add(new SqlParameter("@OPERATING_CONDITIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["OPERATING_CONDITIONS"]) });
                        parameters.Add(new SqlParameter("@LIST_BOUNDARY", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["LIST_BOUNDARY"]) });
                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["DESCRIPTIONS"]) });
                        parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["CREATE_BY"]) });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });
                        #endregion insert
                    }
                    else if (action_type == "update")
                    {
                        #region update
                        sqlstr = "UPDATE EPHA_T_LIST SET " +
                                 "NO = @NO, LIST = @LIST, DESIGN_INTENT = @DESIGN_INTENT, DESIGN_CONDITIONS = @DESIGN_CONDITIONS, OPERATING_CONDITIONS = @OPERATING_CONDITIONS, " +
                                 "LIST_BOUNDARY = @LIST_BOUNDARY, DESCRIPTIONS = @DESCRIPTIONS, UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                 "WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA";

                        parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["NO"]) });
                        parameters.Add(new SqlParameter("@LIST", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["LIST"]) });
                        parameters.Add(new SqlParameter("@DESIGN_INTENT", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["DESIGN_INTENT"]) });
                        parameters.Add(new SqlParameter("@DESIGN_CONDITIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["DESIGN_CONDITIONS"]) });
                        parameters.Add(new SqlParameter("@OPERATING_CONDITIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["OPERATING_CONDITIONS"]) });
                        parameters.Add(new SqlParameter("@LIST_BOUNDARY", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["LIST_BOUNDARY"]) });
                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["DESCRIPTIONS"]) });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });
                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_PHA"]) });
                        #endregion update
                    }
                    else if (action_type == "delete")
                    {
                        #region delete
                        sqlstr = "DELETE FROM EPHA_T_LIST WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_PHA"]) });
                        #endregion delete
                    }

                    if (action_type != "")
                    {
                        ClassConnectionDb cls_conn = new ClassConnectionDb();
                        ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                        if (ret != "true") { break; }
                    }
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return ret;
        }
        public string UpdateListDrawingData(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet? dsData, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";

            try
            {
                DataTable dt = dsData?.Tables["listdrawing"]?.Copy() ?? new DataTable();
                if (dt?.Rows?.Count == 0) { return "true"; }

                for (int i = 0; i < dt?.Rows.Count; i++)
                {
                    List<SqlParameter> parameters = new List<SqlParameter>();
                    string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                    string sqlstr = "";

                    if (action_type == "insert")
                    {
                        #region insert
                        sqlstr = "INSERT INTO EPHA_T_LIST_DRAWING (" +
                                 "SEQ, ID, ID_PHA, ID_LIST, ID_DRAWING, NO, PAGE_START_FIRST, PAGE_END_FIRST, PAGE_START_SECOND, PAGE_END_SECOND, " +
                                 "PAGE_START_THIRD, PAGE_END_THIRD, DESCRIPTIONS, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                 "VALUES (@SEQ, @ID, @ID_PHA, @ID_LIST, @ID_DRAWING, @NO, @PAGE_START_FIRST, @PAGE_END_FIRST, @PAGE_START_SECOND, @PAGE_END_SECOND, " +
                                 "@PAGE_START_THIRD, @PAGE_END_THIRD, @DESCRIPTIONS, GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_header_now) });
                        parameters.Add(new SqlParameter("@ID_LIST", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_LIST"]) });
                        parameters.Add(new SqlParameter("@ID_DRAWING", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_DRAWING"]) });
                        parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["NO"]) });
                        parameters.Add(new SqlParameter("@PAGE_START_FIRST", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["PAGE_START_FIRST"]) });
                        parameters.Add(new SqlParameter("@PAGE_END_FIRST", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["PAGE_END_FIRST"]) });
                        parameters.Add(new SqlParameter("@PAGE_START_SECOND", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["PAGE_START_SECOND"]) });
                        parameters.Add(new SqlParameter("@PAGE_END_SECOND", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["PAGE_END_SECOND"]) });
                        parameters.Add(new SqlParameter("@PAGE_START_THIRD", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["PAGE_START_THIRD"]) });
                        parameters.Add(new SqlParameter("@PAGE_END_THIRD", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["PAGE_END_THIRD"]) });
                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["DESCRIPTIONS"]) });
                        parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["CREATE_BY"]) });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });
                        #endregion insert
                    }
                    else if (action_type == "update")
                    {
                        #region update
                        sqlstr = "UPDATE EPHA_T_LIST_DRAWING SET " +
                                 "NO = @NO, ID_DRAWING = @ID_DRAWING, PAGE_START_FIRST = @PAGE_START_FIRST, PAGE_END_FIRST = @PAGE_END_FIRST, PAGE_START_SECOND = @PAGE_START_SECOND, " +
                                 "PAGE_END_SECOND = @PAGE_END_SECOND, PAGE_START_THIRD = @PAGE_START_THIRD, PAGE_END_THIRD = @PAGE_END_THIRD, DESCRIPTIONS = @DESCRIPTIONS, " +
                                 "UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                 "WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA AND ID_LIST = @ID_LIST";

                        parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["NO"]) });
                        parameters.Add(new SqlParameter("@ID_DRAWING", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_DRAWING"]) });
                        parameters.Add(new SqlParameter("@PAGE_START_FIRST", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["PAGE_START_FIRST"]) });
                        parameters.Add(new SqlParameter("@PAGE_END_FIRST", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["PAGE_END_FIRST"]) });
                        parameters.Add(new SqlParameter("@PAGE_START_SECOND", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["PAGE_START_SECOND"]) });
                        parameters.Add(new SqlParameter("@PAGE_END_SECOND", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["PAGE_END_SECOND"]) });
                        parameters.Add(new SqlParameter("@PAGE_START_THIRD", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["PAGE_START_THIRD"]) });
                        parameters.Add(new SqlParameter("@PAGE_END_THIRD", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["PAGE_END_THIRD"]) });
                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["DESCRIPTIONS"]) });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });
                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_PHA"]) });
                        parameters.Add(new SqlParameter("@ID_LIST", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_LIST"]) });
                        #endregion update
                    }
                    else if (action_type == "delete")
                    {
                        #region delete
                        sqlstr = "DELETE FROM EPHA_T_LIST_DRAWING WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA AND ID_LIST = @ID_LIST";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_PHA"]) });
                        parameters.Add(new SqlParameter("@ID_LIST", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_LIST"]) });
                        #endregion delete
                    }

                    if (action_type != "")
                    {
                        ClassConnectionDb cls_conn = new ClassConnectionDb();
                        ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                        if (ret != "true") { break; }
                    }
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return ret;
        }
        public string set_jsea_partiii(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet? dsData, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";
            List<SqlParameter> parameters = new List<SqlParameter>();
            if (dsData?.Tables.Count > 0)
            {
                ret = UpdateTasksWorksheetData(user_name, conn, transaction, dsData, seq_header_now);
                if (ret != "true") { return ret; }
            }
            return ret;
        }
        public string UpdateTasksWorksheetData(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet? dsData, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";

            try
            {
                DataTable dt = dsData?.Tables["tasks_worksheet"]?.Copy() ?? new DataTable();
                if (dt?.Rows?.Count == 0) { return "true"; }

                for (int i = 0; i < dt?.Rows.Count; i++)
                {
                    List<SqlParameter> parameters = new List<SqlParameter>();
                    string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                    string sqlstr = "";

                    if (!string.IsNullOrEmpty(action_type))
                    {
                        if (action_type == "insert")
                        {
                            #region insert
                            sqlstr = "INSERT INTO EPHA_T_TASKS_WORKSHEET (" +
                                     "SEQ, ID, ID_PHA, SEQ_WORKSTEP, SEQ_TASKDESC, SEQ_POTENTAILHAZARD, SEQ_POSSIBLECASE, SEQ_CATEGORY, " +
                                     "NO, ROW_TYPE, WORKSTEP_NO, WORKSTEP, TASKDESC_NO, TASKDESC, POTENTAILHAZARD_NO, POTENTAILHAZARD, POSSIBLECASE_NO, POSSIBLECASE, CATEGORY_NO, CATEGORY_TYPE, " +
                                     "RAM_BEFOR_SECURITY, RAM_BEFOR_LIKELIHOOD, RAM_BEFOR_RISK, MAJOR_ACCIDENT_EVENT, SAFETY_CRITICAL_EQUIPMENT, EXISTING_SAFEGUARDS, RAM_AFTER_SECURITY, RAM_AFTER_LIKELIHOOD, RAM_AFTER_RISK, " +
                                     "RECOMMENDATIONS_NO, RECOMMENDATIONS, SAFETY_CRITICAL_EQUIPMENT_TAG, RESPONDER_ACTION_BY, RESPONDER_USER_NAME, RESPONDER_USER_DISPLAYNAME, ACTION_STATUS, " +
                                     "RAM_ACTION_SECURITY, RAM_ACTION_LIKELIHOOD, RAM_ACTION_RISK, ESTIMATED_START_DATE, ESTIMATED_END_DATE, " +
                                     "CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                     "VALUES (@SEQ, @ID, @ID_PHA, @SEQ_WORKSTEP, @SEQ_TASKDESC, @SEQ_POTENTAILHAZARD, @SEQ_POSSIBLECASE, @SEQ_CATEGORY, " +
                                     "@NO, @ROW_TYPE, @WORKSTEP_NO, @WORKSTEP, @TASKDESC_NO, @TASKDESC, @POTENTAILHAZARD_NO, @POTENTAILHAZARD, @POSSIBLECASE_NO, @POSSIBLECASE, @CATEGORY_NO, @CATEGORY_TYPE, " +
                                     "@RAM_BEFOR_SECURITY, @RAM_BEFOR_LIKELIHOOD, @RAM_BEFOR_RISK, @MAJOR_ACCIDENT_EVENT, @SAFETY_CRITICAL_EQUIPMENT, @EXISTING_SAFEGUARDS, @RAM_AFTER_SECURITY, @RAM_AFTER_LIKELIHOOD, @RAM_AFTER_RISK, " +
                                     "@RECOMMENDATIONS_NO, @RECOMMENDATIONS, @SAFETY_CRITICAL_EQUIPMENT_TAG, @RESPONDER_ACTION_BY, @RESPONDER_USER_NAME, @RESPONDER_USER_DISPLAYNAME, @ACTION_STATUS, " +
                                     "@RAM_ACTION_SECURITY, @RAM_ACTION_LIKELIHOOD, @RAM_ACTION_RISK, @ESTIMATED_START_DATE, @ESTIMATED_END_DATE, " +
                                     "GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                            parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_header_now) });
                            parameters.Add(new SqlParameter("@SEQ_WORKSTEP", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ_WORKSTEP"]) });
                            parameters.Add(new SqlParameter("@SEQ_TASKDESC", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ_TASKDESC"]) });
                            parameters.Add(new SqlParameter("@SEQ_POTENTAILHAZARD", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ_POTENTAILHAZARD"]) });
                            parameters.Add(new SqlParameter("@SEQ_POSSIBLECASE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ_POSSIBLECASE"]) });
                            parameters.Add(new SqlParameter("@SEQ_CATEGORY", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ_CATEGORY"]) });
                            parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["NO"]) });
                            parameters.Add(new SqlParameter("@ROW_TYPE", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["ROW_TYPE"]) });
                            parameters.Add(new SqlParameter("@WORKSTEP_NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["WORKSTEP_NO"]) });
                            parameters.Add(new SqlParameter("@WORKSTEP", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["WORKSTEP"]) });
                            parameters.Add(new SqlParameter("@TASKDESC_NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["TASKDESC_NO"]) });
                            parameters.Add(new SqlParameter("@TASKDESC", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["TASKDESC"]) });
                            parameters.Add(new SqlParameter("@POTENTAILHAZARD_NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["POTENTAILHAZARD_NO"]) });
                            parameters.Add(new SqlParameter("@POTENTAILHAZARD", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["POTENTAILHAZARD"]) });
                            parameters.Add(new SqlParameter("@POSSIBLECASE_NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["POSSIBLECASE_NO"]) });
                            parameters.Add(new SqlParameter("@POSSIBLECASE", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["POSSIBLECASE"]) });
                            parameters.Add(new SqlParameter("@CATEGORY_NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["CATEGORY_NO"]) });
                            parameters.Add(new SqlParameter("@CATEGORY_TYPE", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["CATEGORY_TYPE"]) });
                            parameters.Add(new SqlParameter("@RAM_BEFOR_SECURITY", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["RAM_BEFOR_SECURITY"]) });
                            parameters.Add(new SqlParameter("@RAM_BEFOR_LIKELIHOOD", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["RAM_BEFOR_LIKELIHOOD"]) });
                            parameters.Add(new SqlParameter("@RAM_BEFOR_RISK", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["RAM_BEFOR_RISK"]) });
                            parameters.Add(new SqlParameter("@MAJOR_ACCIDENT_EVENT", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["MAJOR_ACCIDENT_EVENT"]) });
                            parameters.Add(new SqlParameter("@SAFETY_CRITICAL_EQUIPMENT", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["SAFETY_CRITICAL_EQUIPMENT"]) });
                            parameters.Add(new SqlParameter("@EXISTING_SAFEGUARDS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["EXISTING_SAFEGUARDS"]) });
                            parameters.Add(new SqlParameter("@RAM_AFTER_SECURITY", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["RAM_AFTER_SECURITY"]) });
                            parameters.Add(new SqlParameter("@RAM_AFTER_LIKELIHOOD", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["RAM_AFTER_LIKELIHOOD"]) });
                            parameters.Add(new SqlParameter("@RAM_AFTER_RISK", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["RAM_AFTER_RISK"]) });
                            parameters.Add(new SqlParameter("@RECOMMENDATIONS_NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["RECOMMENDATIONS_NO"]) });
                            parameters.Add(new SqlParameter("@RECOMMENDATIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["RECOMMENDATIONS"]) });
                            parameters.Add(new SqlParameter("@SAFETY_CRITICAL_EQUIPMENT_TAG", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["SAFETY_CRITICAL_EQUIPMENT_TAG"]) });
                            parameters.Add(new SqlParameter("@RESPONDER_ACTION_BY", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["RESPONDER_ACTION_BY"]) });
                            parameters.Add(new SqlParameter("@RESPONDER_USER_NAME", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["RESPONDER_USER_NAME"]) });
                            parameters.Add(new SqlParameter("@RESPONDER_USER_DISPLAYNAME", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["RESPONDER_USER_DISPLAYNAME"]) });
                            parameters.Add(new SqlParameter("@ACTION_STATUS", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["ACTION_STATUS"]) });
                            parameters.Add(new SqlParameter("@RAM_ACTION_SECURITY", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["RAM_ACTION_SECURITY"]) });
                            parameters.Add(new SqlParameter("@RAM_ACTION_LIKELIHOOD", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["RAM_ACTION_LIKELIHOOD"]) });
                            parameters.Add(new SqlParameter("@RAM_ACTION_RISK", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["RAM_ACTION_RISK"]) });
                            parameters.Add(new SqlParameter("@ESTIMATED_START_DATE", SqlDbType.DateTime) { Value = ConvertToDBNull(dt.Rows[i]["ESTIMATED_START_DATE"]) });
                            parameters.Add(new SqlParameter("@ESTIMATED_END_DATE", SqlDbType.DateTime) { Value = ConvertToDBNull(dt.Rows[i]["ESTIMATED_END_DATE"]) });
                            parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["CREATE_BY"]) });
                            parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });
                            #endregion insert
                        }
                        else if (action_type == "update")
                        {
                            #region update
                            sqlstr = "UPDATE EPHA_T_TASKS_WORKSHEET SET " +
                                     "SEQ_WORKSTEP = @SEQ_WORKSTEP, SEQ_TASKDESC = @SEQ_TASKDESC, SEQ_POTENTAILHAZARD = @SEQ_POTENTAILHAZARD, SEQ_POSSIBLECASE = @SEQ_POSSIBLECASE, SEQ_CATEGORY = @SEQ_CATEGORY, " +
                                     "NO = @NO, WORKSTEP_NO = @WORKSTEP_NO, WORKSTEP = @WORKSTEP, TASKDESC_NO = @TASKDESC_NO, TASKDESC = @TASKDESC, POTENTAILHAZARD_NO = @POTENTAILHAZARD_NO, POTENTAILHAZARD = @POTENTAILHAZARD, " +
                                     "POSSIBLECASE_NO = @POSSIBLECASE_NO, POSSIBLECASE = @POSSIBLECASE, CATEGORY_NO = @CATEGORY_NO, CATEGORY_TYPE = @CATEGORY_TYPE, RAM_BEFOR_SECURITY = @RAM_BEFOR_SECURITY, RAM_BEFOR_LIKELIHOOD = @RAM_BEFOR_LIKELIHOOD, " +
                                     "RAM_BEFOR_RISK = @RAM_BEFOR_RISK, MAJOR_ACCIDENT_EVENT = @MAJOR_ACCIDENT_EVENT, SAFETY_CRITICAL_EQUIPMENT = @SAFETY_CRITICAL_EQUIPMENT, EXISTING_SAFEGUARDS = @EXISTING_SAFEGUARDS, RAM_AFTER_SECURITY = @RAM_AFTER_SECURITY, " +
                                     "RAM_AFTER_LIKELIHOOD = @RAM_AFTER_LIKELIHOOD, RAM_AFTER_RISK = @RAM_AFTER_RISK, RECOMMENDATIONS_NO = @RECOMMENDATIONS_NO, RECOMMENDATIONS = @RECOMMENDATIONS, SAFETY_CRITICAL_EQUIPMENT_TAG = @SAFETY_CRITICAL_EQUIPMENT_TAG, " +
                                     "RESPONDER_ACTION_BY = @RESPONDER_ACTION_BY, RESPONDER_USER_NAME = @RESPONDER_USER_NAME, RESPONDER_USER_DISPLAYNAME = @RESPONDER_USER_DISPLAYNAME, ACTION_STATUS = @ACTION_STATUS, RAM_ACTION_SECURITY = @RAM_ACTION_SECURITY, " +
                                     "RAM_ACTION_LIKELIHOOD = @RAM_ACTION_LIKELIHOOD, RAM_ACTION_RISK = @RAM_ACTION_RISK, ESTIMATED_START_DATE = @ESTIMATED_START_DATE, ESTIMATED_END_DATE = @ESTIMATED_END_DATE, UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                     "WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA";

                            parameters.Add(new SqlParameter("@SEQ_WORKSTEP", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ_WORKSTEP"]) });
                            parameters.Add(new SqlParameter("@SEQ_TASKDESC", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ_TASKDESC"]) });
                            parameters.Add(new SqlParameter("@SEQ_POTENTAILHAZARD", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ_POTENTAILHAZARD"]) });
                            parameters.Add(new SqlParameter("@SEQ_POSSIBLECASE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ_POSSIBLECASE"]) });
                            parameters.Add(new SqlParameter("@SEQ_CATEGORY", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ_CATEGORY"]) });
                            parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["NO"]) });
                            parameters.Add(new SqlParameter("@WORKSTEP_NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["WORKSTEP_NO"]) });
                            parameters.Add(new SqlParameter("@WORKSTEP", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["WORKSTEP"]) });
                            parameters.Add(new SqlParameter("@TASKDESC_NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["TASKDESC_NO"]) });
                            parameters.Add(new SqlParameter("@TASKDESC", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["TASKDESC"]) });
                            parameters.Add(new SqlParameter("@POTENTAILHAZARD_NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["POTENTAILHAZARD_NO"]) });
                            parameters.Add(new SqlParameter("@POTENTAILHAZARD", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["POTENTAILHAZARD"]) });
                            parameters.Add(new SqlParameter("@POSSIBLECASE_NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["POSSIBLECASE_NO"]) });
                            parameters.Add(new SqlParameter("@POSSIBLECASE", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["POSSIBLECASE"]) });
                            parameters.Add(new SqlParameter("@CATEGORY_NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["CATEGORY_NO"]) });
                            parameters.Add(new SqlParameter("@CATEGORY_TYPE", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["CATEGORY_TYPE"]) });
                            parameters.Add(new SqlParameter("@RAM_BEFOR_SECURITY", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["RAM_BEFOR_SECURITY"]) });
                            parameters.Add(new SqlParameter("@RAM_BEFOR_LIKELIHOOD", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["RAM_BEFOR_LIKELIHOOD"]) });
                            parameters.Add(new SqlParameter("@RAM_BEFOR_RISK", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["RAM_BEFOR_RISK"]) });
                            parameters.Add(new SqlParameter("@MAJOR_ACCIDENT_EVENT", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["MAJOR_ACCIDENT_EVENT"]) });
                            parameters.Add(new SqlParameter("@SAFETY_CRITICAL_EQUIPMENT", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["SAFETY_CRITICAL_EQUIPMENT"]) });
                            parameters.Add(new SqlParameter("@EXISTING_SAFEGUARDS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["EXISTING_SAFEGUARDS"]) });
                            parameters.Add(new SqlParameter("@RAM_AFTER_SECURITY", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["RAM_AFTER_SECURITY"]) });
                            parameters.Add(new SqlParameter("@RAM_AFTER_LIKELIHOOD", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["RAM_AFTER_LIKELIHOOD"]) });
                            parameters.Add(new SqlParameter("@RAM_AFTER_RISK", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["RAM_AFTER_RISK"]) });
                            parameters.Add(new SqlParameter("@RECOMMENDATIONS_NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["RECOMMENDATIONS_NO"]) });
                            parameters.Add(new SqlParameter("@RECOMMENDATIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["RECOMMENDATIONS"]) });
                            parameters.Add(new SqlParameter("@SAFETY_CRITICAL_EQUIPMENT_TAG", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["SAFETY_CRITICAL_EQUIPMENT_TAG"]) });
                            parameters.Add(new SqlParameter("@RESPONDER_ACTION_BY", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["RESPONDER_ACTION_BY"]) });
                            parameters.Add(new SqlParameter("@RESPONDER_USER_NAME", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["RESPONDER_USER_NAME"]) });
                            parameters.Add(new SqlParameter("@RESPONDER_USER_DISPLAYNAME", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["RESPONDER_USER_DISPLAYNAME"]) });
                            parameters.Add(new SqlParameter("@ACTION_STATUS", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["ACTION_STATUS"]) });
                            parameters.Add(new SqlParameter("@RAM_ACTION_SECURITY", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["RAM_ACTION_SECURITY"]) });
                            parameters.Add(new SqlParameter("@RAM_ACTION_LIKELIHOOD", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["RAM_ACTION_LIKELIHOOD"]) });
                            parameters.Add(new SqlParameter("@RAM_ACTION_RISK", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(dt.Rows[i]["RAM_ACTION_RISK"]) });
                            parameters.Add(new SqlParameter("@ESTIMATED_START_DATE", SqlDbType.DateTime) { Value = ConvertToDBNull(dt.Rows[i]["ESTIMATED_START_DATE"]) });
                            parameters.Add(new SqlParameter("@ESTIMATED_END_DATE", SqlDbType.DateTime) { Value = ConvertToDBNull(dt.Rows[i]["ESTIMATED_END_DATE"]) });
                            parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });
                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                            parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_PHA"]) });
                            #endregion update
                        }
                        else if (action_type == "delete")
                        {
                            #region delete
                            sqlstr = "DELETE FROM EPHA_T_TASKS_WORKSHEET WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA";
                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                            parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_PHA"]) });
                            #endregion delete
                        }

                        if (!string.IsNullOrEmpty(sqlstr))
                        {
                            ClassConnectionDb cls_conn = new ClassConnectionDb();
                            ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                            if (ret != "true") throw new Exception(ret);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ret = ex.Message;
            }

            return ret;
        }

        public string set_whatif_partii(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet? dsData, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";
            List<SqlParameter> parameters = new List<SqlParameter>();
            if (dsData?.Tables.Count > 0)
            {
                ret = UpdateDataList(user_name, conn, transaction, dsData, seq_header_now);
                if (ret != "true") { return ret; }
            }
            if (dsData?.Tables.Count > 0)
            {
                ret = UpdateDataListDrawing(user_name, conn, transaction, dsData, seq_header_now);
                if (ret != "true") { return ret; }
            }
            return ret;
        }
        public string UpdateDataList(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet dsData, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";
            try
            {
                if (dsData?.Tables["list"] == null) return "true";

                DataTable dt = dsData?.Tables["list"]?.Copy() ?? new DataTable();

                foreach (DataRow row in dt.Rows)
                {
                    string action_type = row["action_type"]?.ToString() ?? "";
                    string sqlstr = string.Empty;
                    List<SqlParameter> parameters = new List<SqlParameter>();

                    if (!string.IsNullOrEmpty(action_type))
                    {
                        if (action_type == "insert")
                        {
                            #region insert
                            sqlstr = "INSERT INTO EPHA_T_LIST (" +
                                     "SEQ, ID, ID_PHA, NO, LIST, DESIGN_INTENT, DESIGN_CONDITIONS, OPERATING_CONDITIONS, LIST_BOUNDARY, DESCRIPTIONS, " +
                                     "CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                     "VALUES (@SEQ, @ID, @ID_PHA, @NO, @LIST, @DESIGN_INTENT, @DESIGN_CONDITIONS, @OPERATING_CONDITIONS, @LIST_BOUNDARY, @DESCRIPTIONS, " +
                                     "GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ"]) });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID"]) });
                            parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_header_now) });
                            parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["NO"]) });
                            parameters.Add(new SqlParameter("@LIST", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["LIST"]) });
                            parameters.Add(new SqlParameter("@DESIGN_INTENT", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["DESIGN_INTENT"]) });
                            parameters.Add(new SqlParameter("@DESIGN_CONDITIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["DESIGN_CONDITIONS"]) });
                            parameters.Add(new SqlParameter("@OPERATING_CONDITIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["OPERATING_CONDITIONS"]) });
                            parameters.Add(new SqlParameter("@LIST_BOUNDARY", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["LIST_BOUNDARY"]) });
                            parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["DESCRIPTIONS"]) });
                            parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(row["CREATE_BY"]) });
                            parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(row["UPDATE_BY"]) });
                            #endregion insert
                        }
                        else if (action_type == "update")
                        {
                            #region update
                            sqlstr = "UPDATE EPHA_T_LIST SET " +
                                     "NO = @NO, LIST = @LIST, DESIGN_INTENT = @DESIGN_INTENT, DESIGN_CONDITIONS = @DESIGN_CONDITIONS, " +
                                     "OPERATING_CONDITIONS = @OPERATING_CONDITIONS, LIST_BOUNDARY = @LIST_BOUNDARY, DESCRIPTIONS = @DESCRIPTIONS, " +
                                     "UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                     "WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA";

                            parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["NO"]) });
                            parameters.Add(new SqlParameter("@LIST", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["LIST"]) });
                            parameters.Add(new SqlParameter("@DESIGN_INTENT", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["DESIGN_INTENT"]) });
                            parameters.Add(new SqlParameter("@DESIGN_CONDITIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["DESIGN_CONDITIONS"]) });
                            parameters.Add(new SqlParameter("@OPERATING_CONDITIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["OPERATING_CONDITIONS"]) });
                            parameters.Add(new SqlParameter("@LIST_BOUNDARY", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["LIST_BOUNDARY"]) });
                            parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["DESCRIPTIONS"]) });
                            parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(row["UPDATE_BY"]) });
                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ"]) });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID"]) });
                            parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_PHA"]) });
                            #endregion update
                        }
                        else if (action_type == "delete")
                        {
                            #region delete
                            sqlstr = "DELETE FROM EPHA_T_LIST WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA";
                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ"]) });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID"]) });
                            parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_PHA"]) });
                            #endregion delete
                        }

                        if (!string.IsNullOrEmpty(sqlstr))
                        {
                            ClassConnectionDb cls_conn = new ClassConnectionDb();
                            ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                            if (ret != "true") throw new Exception(ret);
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                ret = ex.Message;
            }

            return ret;
        }
        public string UpdateDataListDrawing(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet dsData, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }
            string ret = "true";
            try
            {
                if (dsData?.Tables["listdrawing"] == null) return "true";

                DataTable dt = dsData.Tables["listdrawing"]?.Copy() ?? new DataTable();

                foreach (DataRow row in dt.Rows)
                {
                    string action_type = row["action_type"]?.ToString() ?? "";
                    string sqlstr = string.Empty;
                    List<SqlParameter> parameters = new List<SqlParameter>();

                    if (!string.IsNullOrEmpty(action_type))
                    {
                        if (action_type == "insert")
                        {
                            #region insert
                            sqlstr = "INSERT INTO EPHA_T_LIST_DRAWING (" +
                                     "SEQ, ID, ID_PHA, ID_LIST, ID_DRAWING, NO, PAGE_START_FIRST, PAGE_END_FIRST, " +
                                     "PAGE_START_SECOND, PAGE_END_SECOND, PAGE_START_THIRD, PAGE_END_THIRD, DESCRIPTIONS, " +
                                     "CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                     "VALUES (@SEQ, @ID, @ID_PHA, @ID_LIST, @ID_DRAWING, @NO, @PAGE_START_FIRST, @PAGE_END_FIRST, " +
                                     "@PAGE_START_SECOND, @PAGE_END_SECOND, @PAGE_START_THIRD, @PAGE_END_THIRD, @DESCRIPTIONS, " +
                                     "GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ"]) });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID"]) });
                            parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_header_now) });
                            parameters.Add(new SqlParameter("@ID_LIST", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_LIST"]) });
                            parameters.Add(new SqlParameter("@ID_DRAWING", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_DRAWING"]) });
                            parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["NO"]) });
                            parameters.Add(new SqlParameter("@PAGE_START_FIRST", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["PAGE_START_FIRST"]) });
                            parameters.Add(new SqlParameter("@PAGE_END_FIRST", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["PAGE_END_FIRST"]) });
                            parameters.Add(new SqlParameter("@PAGE_START_SECOND", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["PAGE_START_SECOND"]) });
                            parameters.Add(new SqlParameter("@PAGE_END_SECOND", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["PAGE_END_SECOND"]) });
                            parameters.Add(new SqlParameter("@PAGE_START_THIRD", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["PAGE_START_THIRD"]) });
                            parameters.Add(new SqlParameter("@PAGE_END_THIRD", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["PAGE_END_THIRD"]) });
                            parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["DESCRIPTIONS"]) });
                            parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(row["CREATE_BY"]) });
                            parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(row["UPDATE_BY"]) });
                            #endregion insert
                        }
                        else if (action_type == "update")
                        {
                            #region update
                            sqlstr = "UPDATE EPHA_T_LIST_DRAWING SET " +
                                     "NO = @NO, ID_DRAWING = @ID_DRAWING, PAGE_START_FIRST = @PAGE_START_FIRST, PAGE_END_FIRST = @PAGE_END_FIRST, " +
                                     "PAGE_START_SECOND = @PAGE_START_SECOND, PAGE_END_SECOND = @PAGE_END_SECOND, PAGE_START_THIRD = @PAGE_START_THIRD, " +
                                     "PAGE_END_THIRD = @PAGE_END_THIRD, DESCRIPTIONS = @DESCRIPTIONS, UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                     "WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA AND ID_LIST = @ID_LIST";

                            parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["NO"]) });
                            parameters.Add(new SqlParameter("@ID_DRAWING", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_DRAWING"]) });
                            parameters.Add(new SqlParameter("@PAGE_START_FIRST", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["PAGE_START_FIRST"]) });
                            parameters.Add(new SqlParameter("@PAGE_END_FIRST", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["PAGE_END_FIRST"]) });
                            parameters.Add(new SqlParameter("@PAGE_START_SECOND", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["PAGE_START_SECOND"]) });
                            parameters.Add(new SqlParameter("@PAGE_END_SECOND", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["PAGE_END_SECOND"]) });
                            parameters.Add(new SqlParameter("@PAGE_START_THIRD", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["PAGE_START_THIRD"]) });
                            parameters.Add(new SqlParameter("@PAGE_END_THIRD", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["PAGE_END_THIRD"]) });
                            parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["DESCRIPTIONS"]) });
                            parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(row["UPDATE_BY"]) });
                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ"]) });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID"]) });
                            parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_PHA"]) });
                            parameters.Add(new SqlParameter("@ID_LIST", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_LIST"]) });
                            #endregion update
                        }
                        else if (action_type == "delete")
                        {
                            #region delete
                            sqlstr = "DELETE FROM EPHA_T_LIST_DRAWING WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA AND ID_LIST = @ID_LIST";

                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ"]) });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID"]) });
                            parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_PHA"]) });
                            parameters.Add(new SqlParameter("@ID_LIST", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_LIST"]) });
                            #endregion delete
                        }

                        if (!string.IsNullOrEmpty(sqlstr))
                        {
                            ClassConnectionDb cls_conn = new ClassConnectionDb();
                            ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                            if (ret != "true") throw new Exception(ret);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ret = ex.Message;
            }

            return ret;
        }

        public string set_whatif_partiii(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet? dsData, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";
            List<SqlParameter> parameters = new List<SqlParameter>();
            if (dsData?.Tables.Count > 0)
            {
                ret = UpdateDataListWorksheet(user_name, conn, transaction, dsData, seq_header_now);
                if (ret != "true") { return ret; }
            }
            return ret;
        }
        public string UpdateDataListWorksheet(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet dsData, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }
            string ret = "true";

            try
            {
                if (dsData?.Tables["listworksheet"] == null) return "true";
                DataTable dt = dsData?.Tables["listworksheet"]?.Copy() ?? new DataTable();

                foreach (DataRow row in dt.Rows)
                {
                    string action_type = row["action_type"]?.ToString() ?? "";
                    string sqlstr = string.Empty;
                    List<SqlParameter> parameters = new List<SqlParameter>();

                    if (action_type == "insert")
                    {
                        #region insert
                        sqlstr = "INSERT INTO EPHA_T_LIST_WORKSHEET (" +
                                 "SEQ, ID, ID_PHA, ID_LIST, ROW_TYPE, " +
                                 "SEQ_LIST_SYSTEM, SEQ_LIST_SUB_SYSTEM, SEQ_CAUSES, SEQ_CONSEQUENCES, SEQ_CATEGORY, " +
                                 "NO, LIST_SYSTEM_NO, LIST_SYSTEM, LIST_SUB_SYSTEM_NO, LIST_SUB_SYSTEM, CAUSES_NO, CAUSES, CONSEQUENCES_NO, CONSEQUENCES, " +
                                 "CATEGORY_NO, CATEGORY_TYPE, RAM_BEFOR_SECURITY, RAM_BEFOR_LIKELIHOOD, RAM_BEFOR_RISK, MAJOR_ACCIDENT_EVENT, SAFETY_CRITICAL_EQUIPMENT, SAFETY_CRITICAL_EQUIPMENT_TAG, EXISTING_SAFEGUARDS, " +
                                 "RAM_AFTER_SECURITY, RAM_AFTER_LIKELIHOOD, RAM_AFTER_RISK, FK_RECOMMENDATIONS, SEQ_RECOMMENDATIONS, RECOMMENDATIONS_NO, RECOMMENDATIONS, RESPONDER_USER_NAME, RESPONDER_USER_DISPLAYNAME, " +
                                 "RAM_ACTION_SECURITY, RAM_ACTION_LIKELIHOOD, RAM_ACTION_RISK, ESTIMATED_START_DATE, ESTIMATED_END_DATE, ACTION_STATUS, " +
                                 "IMPLEMENT, RESPONDER_ACTION_TYPE, RESPONDER_ACTION_DATE, REVIEWER_ACTION_TYPE, REVIEWER_ACTION_DATE, " +
                                 "ACTION_PROJECT_TEAM, PROJECT_TEAM_TEXT, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                 "VALUES (" +
                                 "@SEQ, @ID, @ID_PHA, @ID_LIST, @ROW_TYPE, " +
                                 "@SEQ_LIST_SYSTEM, @SEQ_LIST_SUB_SYSTEM, @SEQ_CAUSES, @SEQ_CONSEQUENCES, @SEQ_CATEGORY, " +
                                 "@NO, @LIST_SYSTEM_NO, @LIST_SYSTEM, @LIST_SUB_SYSTEM_NO, @LIST_SUB_SYSTEM, @CAUSES_NO, @CAUSES, @CONSEQUENCES_NO, @CONSEQUENCES, " +
                                 "@CATEGORY_NO, @CATEGORY_TYPE, @RAM_BEFOR_SECURITY, @RAM_BEFOR_LIKELIHOOD, @RAM_BEFOR_RISK, @MAJOR_ACCIDENT_EVENT, @SAFETY_CRITICAL_EQUIPMENT, @SAFETY_CRITICAL_EQUIPMENT_TAG, @EXISTING_SAFEGUARDS, " +
                                 "@RAM_AFTER_SECURITY, @RAM_AFTER_LIKELIHOOD, @RAM_AFTER_RISK, @FK_RECOMMENDATIONS, @SEQ_RECOMMENDATIONS, @RECOMMENDATIONS_NO, @RECOMMENDATIONS, @RESPONDER_USER_NAME, @RESPONDER_USER_DISPLAYNAME, " +
                                 "@RAM_ACTION_SECURITY, @RAM_ACTION_LIKELIHOOD, @RAM_ACTION_RISK, @ESTIMATED_START_DATE, @ESTIMATED_END_DATE, @ACTION_STATUS, " +
                                 "@IMPLEMENT, @RESPONDER_ACTION_TYPE, @RESPONDER_ACTION_DATE, @REVIEWER_ACTION_TYPE, @REVIEWER_ACTION_DATE, " +
                                 "@ACTION_PROJECT_TEAM, @PROJECT_TEAM_TEXT, GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = seq_header_now });
                        parameters.Add(new SqlParameter("@ID_LIST", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_LIST"]) });
                        parameters.Add(new SqlParameter("@ROW_TYPE", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(row["ROW_TYPE"]) });
                        parameters.Add(new SqlParameter("@SEQ_LIST_SYSTEM", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ_LIST_SYSTEM"]) });
                        parameters.Add(new SqlParameter("@SEQ_LIST_SUB_SYSTEM", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ_LIST_SUB_SYSTEM"]) });
                        parameters.Add(new SqlParameter("@SEQ_CAUSES", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ_CAUSES"]) });
                        parameters.Add(new SqlParameter("@SEQ_CONSEQUENCES", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ_CONSEQUENCES"]) });
                        parameters.Add(new SqlParameter("@SEQ_CATEGORY", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ_CATEGORY"]) });
                        parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["NO"]) });
                        parameters.Add(new SqlParameter("@LIST_SYSTEM_NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["LIST_SYSTEM_NO"]) });
                        parameters.Add(new SqlParameter("@LIST_SYSTEM", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["LIST_SYSTEM"]) });
                        parameters.Add(new SqlParameter("@LIST_SUB_SYSTEM_NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["LIST_SUB_SYSTEM_NO"]) });
                        parameters.Add(new SqlParameter("@LIST_SUB_SYSTEM", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["LIST_SUB_SYSTEM"]) });
                        parameters.Add(new SqlParameter("@CAUSES_NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["CAUSES_NO"]) });
                        parameters.Add(new SqlParameter("@CAUSES", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["CAUSES"]) });
                        parameters.Add(new SqlParameter("@CONSEQUENCES_NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["CONSEQUENCES_NO"]) });
                        parameters.Add(new SqlParameter("@CONSEQUENCES", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["CONSEQUENCES"]) });
                        parameters.Add(new SqlParameter("@CATEGORY_NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["CATEGORY_NO"]) });
                        parameters.Add(new SqlParameter("@CATEGORY_TYPE", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["CATEGORY_TYPE"]) });
                        parameters.Add(new SqlParameter("@RAM_BEFOR_SECURITY", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(row["RAM_BEFOR_SECURITY"]) });
                        parameters.Add(new SqlParameter("@RAM_BEFOR_LIKELIHOOD", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(row["RAM_BEFOR_LIKELIHOOD"]) });
                        parameters.Add(new SqlParameter("@RAM_BEFOR_RISK", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(row["RAM_BEFOR_RISK"]) });
                        parameters.Add(new SqlParameter("@MAJOR_ACCIDENT_EVENT", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(row["MAJOR_ACCIDENT_EVENT"]) });
                        parameters.Add(new SqlParameter("@SAFETY_CRITICAL_EQUIPMENT", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(row["SAFETY_CRITICAL_EQUIPMENT"]) });
                        parameters.Add(new SqlParameter("@SAFETY_CRITICAL_EQUIPMENT_TAG", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["SAFETY_CRITICAL_EQUIPMENT_TAG"]) });
                        parameters.Add(new SqlParameter("@EXISTING_SAFEGUARDS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["EXISTING_SAFEGUARDS"]) });
                        parameters.Add(new SqlParameter("@RAM_AFTER_SECURITY", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(row["RAM_AFTER_SECURITY"]) });
                        parameters.Add(new SqlParameter("@RAM_AFTER_LIKELIHOOD", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(row["RAM_AFTER_LIKELIHOOD"]) });
                        parameters.Add(new SqlParameter("@RAM_AFTER_RISK", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(row["RAM_AFTER_RISK"]) });
                        parameters.Add(new SqlParameter("@FK_RECOMMENDATIONS", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["FK_RECOMMENDATIONS"]) });
                        parameters.Add(new SqlParameter("@SEQ_RECOMMENDATIONS", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ_RECOMMENDATIONS"]) });
                        parameters.Add(new SqlParameter("@RECOMMENDATIONS_NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["RECOMMENDATIONS_NO"]) });
                        parameters.Add(new SqlParameter("@RECOMMENDATIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["RECOMMENDATIONS"]) });
                        parameters.Add(new SqlParameter("@RESPONDER_USER_NAME", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(row["RESPONDER_USER_NAME"]) });
                        parameters.Add(new SqlParameter("@RESPONDER_USER_DISPLAYNAME", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["RESPONDER_USER_DISPLAYNAME"]) });
                        parameters.Add(new SqlParameter("@RAM_ACTION_SECURITY", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(row["RAM_ACTION_SECURITY"]) });
                        parameters.Add(new SqlParameter("@RAM_ACTION_LIKELIHOOD", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(row["RAM_ACTION_LIKELIHOOD"]) });
                        parameters.Add(new SqlParameter("@RAM_ACTION_RISK", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(row["RAM_ACTION_RISK"]) });
                        parameters.Add(new SqlParameter("@ESTIMATED_START_DATE", SqlDbType.Date) { Value = ConvertToDBNull(row["ESTIMATED_START_DATE"]) });
                        parameters.Add(new SqlParameter("@ESTIMATED_END_DATE", SqlDbType.Date) { Value = ConvertToDBNull(row["ESTIMATED_END_DATE"]) });
                        parameters.Add(new SqlParameter("@ACTION_STATUS", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(row["ACTION_STATUS"]) });
                        parameters.Add(new SqlParameter("@IMPLEMENT", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["IMPLEMENT"]) });
                        parameters.Add(new SqlParameter("@RESPONDER_ACTION_TYPE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["RESPONDER_ACTION_TYPE"]) });
                        parameters.Add(new SqlParameter("@RESPONDER_ACTION_DATE", SqlDbType.Date) { Value = ConvertToDBNull(row["RESPONDER_ACTION_DATE"]) });
                        parameters.Add(new SqlParameter("@REVIEWER_ACTION_TYPE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["REVIEWER_ACTION_TYPE"]) });
                        parameters.Add(new SqlParameter("@REVIEWER_ACTION_DATE", SqlDbType.Date) { Value = ConvertToDBNull(row["REVIEWER_ACTION_DATE"]) });
                        parameters.Add(new SqlParameter("@ACTION_PROJECT_TEAM", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ACTION_PROJECT_TEAM"]) });
                        parameters.Add(new SqlParameter("@PROJECT_TEAM_TEXT", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["PROJECT_TEAM_TEXT"]) });
                        parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(row["CREATE_BY"]) });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(row["UPDATE_BY"]) });
                        #endregion insert
                    }
                    else if (action_type == "update")
                    {
                        #region update
                        sqlstr = "UPDATE EPHA_T_LIST_WORKSHEET SET " +
                                 "NO = @NO, " +
                                 "SEQ_LIST_SYSTEM = @SEQ_LIST_SYSTEM, SEQ_LIST_SUB_SYSTEM = @SEQ_LIST_SUB_SYSTEM, SEQ_CAUSES = @SEQ_CAUSES, SEQ_CONSEQUENCES = @SEQ_CONSEQUENCES, SEQ_CATEGORY = @SEQ_CATEGORY, " +
                                 "LIST_SYSTEM_NO = @LIST_SYSTEM_NO, LIST_SYSTEM = @LIST_SYSTEM, LIST_SUB_SYSTEM_NO = @LIST_SUB_SYSTEM_NO, LIST_SUB_SYSTEM = @LIST_SUB_SYSTEM, " +
                                 "CAUSES_NO = @CAUSES_NO, CAUSES = @CAUSES, CONSEQUENCES_NO = @CONSEQUENCES_NO, CONSEQUENCES = @CONSEQUENCES, " +
                                 "CATEGORY_NO = @CATEGORY_NO, CATEGORY_TYPE = @CATEGORY_TYPE, RAM_BEFOR_SECURITY = @RAM_BEFOR_SECURITY, RAM_BEFOR_LIKELIHOOD = @RAM_BEFOR_LIKELIHOOD, RAM_BEFOR_RISK = @RAM_BEFOR_RISK, " +
                                 "MAJOR_ACCIDENT_EVENT = @MAJOR_ACCIDENT_EVENT, SAFETY_CRITICAL_EQUIPMENT = @SAFETY_CRITICAL_EQUIPMENT, SAFETY_CRITICAL_EQUIPMENT_TAG = @SAFETY_CRITICAL_EQUIPMENT_TAG, EXISTING_SAFEGUARDS = @EXISTING_SAFEGUARDS, " +
                                 "RAM_AFTER_SECURITY = @RAM_AFTER_SECURITY, RAM_AFTER_LIKELIHOOD = @RAM_AFTER_LIKELIHOOD, RAM_AFTER_RISK = @RAM_AFTER_RISK, " +
                                 "FK_RECOMMENDATIONS = @FK_RECOMMENDATIONS, SEQ_RECOMMENDATIONS = @SEQ_RECOMMENDATIONS, RECOMMENDATIONS_NO = @RECOMMENDATIONS_NO, RECOMMENDATIONS = @RECOMMENDATIONS, " +
                                 "RESPONDER_USER_NAME = @RESPONDER_USER_NAME, RESPONDER_USER_DISPLAYNAME = @RESPONDER_USER_DISPLAYNAME, " +
                                 "RAM_ACTION_SECURITY = @RAM_ACTION_SECURITY, RAM_ACTION_LIKELIHOOD = @RAM_ACTION_LIKELIHOOD, RAM_ACTION_RISK = @RAM_ACTION_RISK, " +
                                 "ESTIMATED_START_DATE = @ESTIMATED_START_DATE, ESTIMATED_END_DATE = @ESTIMATED_END_DATE, ACTION_STATUS = @ACTION_STATUS, " +
                                 "IMPLEMENT = @IMPLEMENT, RESPONDER_ACTION_TYPE = @RESPONDER_ACTION_TYPE, RESPONDER_ACTION_DATE = @RESPONDER_ACTION_DATE, REVIEWER_ACTION_TYPE = @REVIEWER_ACTION_TYPE, REVIEWER_ACTION_DATE = @REVIEWER_ACTION_DATE, " +
                                 "ACTION_PROJECT_TEAM = @ACTION_PROJECT_TEAM, PROJECT_TEAM_TEXT = @PROJECT_TEAM_TEXT, " +
                                 "UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                 "WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA AND ID_LIST = @ID_LIST";

                        parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["NO"]) });
                        parameters.Add(new SqlParameter("@SEQ_LIST_SYSTEM", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ_LIST_SYSTEM"]) });
                        parameters.Add(new SqlParameter("@SEQ_LIST_SUB_SYSTEM", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ_LIST_SUB_SYSTEM"]) });
                        parameters.Add(new SqlParameter("@SEQ_CAUSES", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ_CAUSES"]) });
                        parameters.Add(new SqlParameter("@SEQ_CONSEQUENCES", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ_CONSEQUENCES"]) });
                        parameters.Add(new SqlParameter("@SEQ_CATEGORY", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ_CATEGORY"]) });
                        parameters.Add(new SqlParameter("@LIST_SYSTEM_NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["LIST_SYSTEM_NO"]) });
                        parameters.Add(new SqlParameter("@LIST_SYSTEM", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["LIST_SYSTEM"]) });
                        parameters.Add(new SqlParameter("@LIST_SUB_SYSTEM_NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["LIST_SUB_SYSTEM_NO"]) });
                        parameters.Add(new SqlParameter("@LIST_SUB_SYSTEM", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["LIST_SUB_SYSTEM"]) });
                        parameters.Add(new SqlParameter("@CAUSES_NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["CAUSES_NO"]) });
                        parameters.Add(new SqlParameter("@CAUSES", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["CAUSES"]) });
                        parameters.Add(new SqlParameter("@CONSEQUENCES_NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["CONSEQUENCES_NO"]) });
                        parameters.Add(new SqlParameter("@CONSEQUENCES", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["CONSEQUENCES"]) });
                        parameters.Add(new SqlParameter("@CATEGORY_NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["CATEGORY_NO"]) });
                        parameters.Add(new SqlParameter("@CATEGORY_TYPE", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["CATEGORY_TYPE"]) });
                        parameters.Add(new SqlParameter("@RAM_BEFOR_SECURITY", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(row["RAM_BEFOR_SECURITY"]) });
                        parameters.Add(new SqlParameter("@RAM_BEFOR_LIKELIHOOD", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(row["RAM_BEFOR_LIKELIHOOD"]) });
                        parameters.Add(new SqlParameter("@RAM_BEFOR_RISK", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(row["RAM_BEFOR_RISK"]) });
                        parameters.Add(new SqlParameter("@MAJOR_ACCIDENT_EVENT", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(row["MAJOR_ACCIDENT_EVENT"]) });
                        parameters.Add(new SqlParameter("@SAFETY_CRITICAL_EQUIPMENT", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(row["SAFETY_CRITICAL_EQUIPMENT"]) });
                        parameters.Add(new SqlParameter("@SAFETY_CRITICAL_EQUIPMENT_TAG", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["SAFETY_CRITICAL_EQUIPMENT_TAG"]) });
                        parameters.Add(new SqlParameter("@EXISTING_SAFEGUARDS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["EXISTING_SAFEGUARDS"]) });
                        parameters.Add(new SqlParameter("@RAM_AFTER_SECURITY", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(row["RAM_AFTER_SECURITY"]) });
                        parameters.Add(new SqlParameter("@RAM_AFTER_LIKELIHOOD", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(row["RAM_AFTER_LIKELIHOOD"]) });
                        parameters.Add(new SqlParameter("@RAM_AFTER_RISK", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(row["RAM_AFTER_RISK"]) });
                        parameters.Add(new SqlParameter("@FK_RECOMMENDATIONS", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["FK_RECOMMENDATIONS"]) });
                        parameters.Add(new SqlParameter("@SEQ_RECOMMENDATIONS", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ_RECOMMENDATIONS"]) });
                        parameters.Add(new SqlParameter("@RECOMMENDATIONS_NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["RECOMMENDATIONS_NO"]) });
                        parameters.Add(new SqlParameter("@RECOMMENDATIONS", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["RECOMMENDATIONS"]) });
                        parameters.Add(new SqlParameter("@RESPONDER_USER_NAME", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(row["RESPONDER_USER_NAME"]) });
                        parameters.Add(new SqlParameter("@RESPONDER_USER_DISPLAYNAME", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["RESPONDER_USER_DISPLAYNAME"]) });
                        parameters.Add(new SqlParameter("@RAM_ACTION_SECURITY", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(row["RAM_ACTION_SECURITY"]) });
                        parameters.Add(new SqlParameter("@RAM_ACTION_LIKELIHOOD", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(row["RAM_ACTION_LIKELIHOOD"]) });
                        parameters.Add(new SqlParameter("@RAM_ACTION_RISK", SqlDbType.VarChar, 10) { Value = ConvertToDBNull(row["RAM_ACTION_RISK"]) });
                        parameters.Add(new SqlParameter("@ESTIMATED_START_DATE", SqlDbType.Date) { Value = ConvertToDBNull(row["ESTIMATED_START_DATE"]) });
                        parameters.Add(new SqlParameter("@ESTIMATED_END_DATE", SqlDbType.Date) { Value = ConvertToDBNull(row["ESTIMATED_END_DATE"]) });
                        parameters.Add(new SqlParameter("@ACTION_STATUS", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(row["ACTION_STATUS"]) });
                        parameters.Add(new SqlParameter("@IMPLEMENT", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["IMPLEMENT"]) });
                        parameters.Add(new SqlParameter("@RESPONDER_ACTION_TYPE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["RESPONDER_ACTION_TYPE"]) });
                        parameters.Add(new SqlParameter("@RESPONDER_ACTION_DATE", SqlDbType.Date) { Value = ConvertToDBNull(row["RESPONDER_ACTION_DATE"]) });
                        parameters.Add(new SqlParameter("@REVIEWER_ACTION_TYPE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["REVIEWER_ACTION_TYPE"]) });
                        parameters.Add(new SqlParameter("@REVIEWER_ACTION_DATE", SqlDbType.Date) { Value = ConvertToDBNull(row["REVIEWER_ACTION_DATE"]) });
                        parameters.Add(new SqlParameter("@ACTION_PROJECT_TEAM", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ACTION_PROJECT_TEAM"]) });
                        parameters.Add(new SqlParameter("@PROJECT_TEAM_TEXT", SqlDbType.VarChar, 4000) { Value = ConvertToDBNull(row["PROJECT_TEAM_TEXT"]) });
                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.VarChar, 50) { Value = ConvertToDBNull(row["UPDATE_BY"]) });

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = seq_header_now });
                        parameters.Add(new SqlParameter("@ID_LIST", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_LIST"]) });
                        #endregion update
                    }
                    else if (action_type == "delete")
                    {
                        #region delete
                        sqlstr = "DELETE FROM EPHA_T_LIST_WORKSHEET " +
                                 "WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA AND ID_LIST = @ID_LIST";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID"]) });
                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = seq_header_now });
                        parameters.Add(new SqlParameter("@ID_LIST", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_LIST"]) });
                        #endregion delete
                    }

                    if (!string.IsNullOrEmpty(sqlstr))
                    {
                        ClassConnectionDb cls_conn = new ClassConnectionDb();
                        ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                        if (ret != "true") throw new Exception(ret);
                    }
                }

                if (ret == "true")
                {
                    // Delete data not in task list
                    sqlstr = "DELETE FROM EPHA_T_LIST_WORKSHEET WHERE id_pha = @ID_PHA AND id_list NOT IN (SELECT id FROM EPHA_T_LIST WHERE id_pha = @ID_PHA)";

                    List<SqlParameter> parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_header_now) });
                    if (!string.IsNullOrEmpty(sqlstr))
                    {
                        ClassConnectionDb cls_conn = new ClassConnectionDb();
                        ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                        if (ret != "true") throw new Exception(ret);
                    }
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }


            return ret;
        }

        public string set_hra_partii(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet? dsData, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";
            List<SqlParameter> parameters = new List<SqlParameter>();
            if (dsData?.Tables.Count > 0)
            {
                ret = UpdateSubareasData(user_name, conn, transaction, dsData, seq_header_now);
                if (ret != "true") { return ret; }
            }
            if (dsData?.Tables.Count > 0)
            {
                ret = UpdateHazardData(user_name, conn, transaction, dsData, seq_header_now);
                if (ret != "true") { return ret; }
            }

            return ret;

        }
        public string UpdateSubareasData(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet dsData, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";

            try
            {
                // Copy the DataTable from the DataSet
                DataTable dt = dsData?.Tables["subareas"]?.Copy() ?? new DataTable();
                if (dt.Rows.Count == 0) return "true";

                #region Update Data Subareas
                foreach (DataRow row in dt.Rows)
                {
                    string action_type = row["action_type"]?.ToString() ?? "";
                    string sqlstr = "";
                    List<SqlParameter> parameters = new List<SqlParameter>();

                    if (!string.IsNullOrEmpty(action_type))
                    {
                        switch (action_type)
                        {
                            case "insert":
                                #region Insert
                                sqlstr = "INSERT INTO EPHA_T_TABLE1_SUBAREAS (" +
                                         "SEQ, ID, ID_PHA, NO, ID_BUSINESS_UNIT, ID_SUB_AREA, SUB_AREA, WORK_OF_TASK, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY" +
                                         ") VALUES (" +
                                         "@SEQ, @ID, @ID_PHA, @NO, @ID_BUSINESS_UNIT, @ID_SUB_AREA, @SUB_AREA, @WORK_OF_TASK, GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ"]) });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID"]) });
                                parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_header_now) });
                                parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["NO"]) });
                                parameters.Add(new SqlParameter("@ID_BUSINESS_UNIT", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_BUSINESS_UNIT"]) });
                                parameters.Add(new SqlParameter("@ID_SUB_AREA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_SUB_AREA"]) });
                                parameters.Add(new SqlParameter("@SUB_AREA", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["SUB_AREA"]) });
                                parameters.Add(new SqlParameter("@WORK_OF_TASK", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["WORK_OF_TASK"]) });
                                parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = ConvertToDBNull(row["CREATE_BY"]) });
                                parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = ConvertToDBNull(row["UPDATE_BY"]) });
                                #endregion
                                break;

                            case "update":
                                #region Update
                                sqlstr = "UPDATE EPHA_T_TABLE1_SUBAREAS SET " +
                                         "NO = @NO, ID_BUSINESS_UNIT = @ID_BUSINESS_UNIT, ID_SUB_AREA = @ID_SUB_AREA, SUB_AREA = @SUB_AREA, WORK_OF_TASK = @WORK_OF_TASK, " +
                                         "UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                         "WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA";

                                parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["NO"]) });
                                parameters.Add(new SqlParameter("@ID_BUSINESS_UNIT", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_BUSINESS_UNIT"]) });
                                parameters.Add(new SqlParameter("@ID_SUB_AREA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_SUB_AREA"]) });
                                parameters.Add(new SqlParameter("@SUB_AREA", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["SUB_AREA"]) });
                                parameters.Add(new SqlParameter("@WORK_OF_TASK", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["WORK_OF_TASK"]) });
                                parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = ConvertToDBNull(row["UPDATE_BY"]) });
                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ"]) });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID"]) });
                                parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_PHA"]) });
                                #endregion
                                break;

                            case "delete":
                                #region Delete
                                sqlstr = "DELETE FROM EPHA_T_TABLE1_SUBAREAS " +
                                         "WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA";

                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ"]) });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID"]) });
                                parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_PHA"]) });
                                #endregion
                                break;
                        }

                        if (!string.IsNullOrEmpty(sqlstr))
                        {
                            ClassConnectionDb cls_conn = new ClassConnectionDb();
                            ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                            if (ret != "true") throw new Exception(ret);
                        }
                    }
                }
                #endregion Update Data Subareas

            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return ret;
        }
        public string UpdateHazardData(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet dsData, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";

            try
            {
                DataTable dt = dsData?.Tables["hazard"]?.Copy() ?? new DataTable();
                if (dt.Rows.Count == 0) return "true";

                #region Update Data Hazard
                foreach (DataRow row in dt.Rows)
                {
                    string action_type = row["action_type"]?.ToString() ?? "";
                    string sqlstr = "";
                    List<SqlParameter> parameters = new List<SqlParameter>();

                    if (!string.IsNullOrEmpty(action_type))
                    {
                        switch (action_type)
                        {
                            case "insert":
                                #region Insert
                                sqlstr = "INSERT INTO EPHA_T_TABLE1_HAZARD (" +
                                         "SEQ, ID, ID_PHA, NO, ID_BUSINESS_UNIT, ID_TYPE_HAZARD, ID_HEALTH_HAZARD, ID_HEALTH_EFFECT, TYPE_HAZARD, HEALTH_HAZARD, HEALTH_EFFECT_RATING, " +
                                         "NO_SUBAREAS, ID_SUBAREAS, SUB_AREA, WORK_OF_TASK, NO_TYPE_HAZARD, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY" +
                                         ") VALUES (" +
                                         "@SEQ, @ID, @ID_PHA, @NO, @ID_BUSINESS_UNIT, @ID_TYPE_HAZARD, @ID_HEALTH_HAZARD, @ID_HEALTH_EFFECT, @TYPE_HAZARD, @HEALTH_HAZARD, @HEALTH_EFFECT_RATING, " +
                                         "@NO_SUBAREAS, @ID_SUBAREAS, @SUB_AREA, @WORK_OF_TASK, @NO_TYPE_HAZARD, GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ"]) });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID"]) });
                                parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_header_now) });
                                parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["NO"]) });
                                parameters.Add(new SqlParameter("@ID_BUSINESS_UNIT", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_BUSINESS_UNIT"]) });
                                parameters.Add(new SqlParameter("@ID_TYPE_HAZARD", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_TYPE_HAZARD"]) });
                                parameters.Add(new SqlParameter("@ID_HEALTH_HAZARD", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_HEALTH_HAZARD"]) });
                                parameters.Add(new SqlParameter("@ID_HEALTH_EFFECT", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_HEALTH_EFFECT"]) });
                                parameters.Add(new SqlParameter("@TYPE_HAZARD", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["TYPE_HAZARD"]) });
                                parameters.Add(new SqlParameter("@HEALTH_HAZARD", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["HEALTH_HAZARD"]) });
                                parameters.Add(new SqlParameter("@HEALTH_EFFECT_RATING", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["HEALTH_EFFECT_RATING"]) });
                                parameters.Add(new SqlParameter("@NO_SUBAREAS", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["NO_SUBAREAS"]) });
                                parameters.Add(new SqlParameter("@ID_SUBAREAS", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_SUBAREAS"]) });
                                parameters.Add(new SqlParameter("@SUB_AREA", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["SUB_AREA"]) });
                                parameters.Add(new SqlParameter("@WORK_OF_TASK", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["WORK_OF_TASK"]) });
                                parameters.Add(new SqlParameter("@NO_TYPE_HAZARD", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["NO_TYPE_HAZARD"]) });
                                parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = ConvertToDBNull(row["CREATE_BY"]) });
                                parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = ConvertToDBNull(row["UPDATE_BY"]) });
                                #endregion
                                break;

                            case "update":
                                #region Update
                                sqlstr = "UPDATE EPHA_T_TABLE1_HAZARD SET " +
                                         "NO = @NO, ID_BUSINESS_UNIT = @ID_BUSINESS_UNIT, ID_TYPE_HAZARD = @ID_TYPE_HAZARD, ID_HEALTH_HAZARD = @ID_HEALTH_HAZARD, ID_HEALTH_EFFECT = @ID_HEALTH_EFFECT, " +
                                         "TYPE_HAZARD = @TYPE_HAZARD, HEALTH_HAZARD = @HEALTH_HAZARD, HEALTH_EFFECT_RATING = @HEALTH_EFFECT_RATING, " +
                                         "NO_SUBAREAS = @NO_SUBAREAS, ID_SUBAREAS = @ID_SUBAREAS, SUB_AREA = @SUB_AREA, WORK_OF_TASK = @WORK_OF_TASK, " +
                                         "NO_TYPE_HAZARD = @NO_TYPE_HAZARD, UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                         "WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA";

                                parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["NO"]) });
                                parameters.Add(new SqlParameter("@ID_BUSINESS_UNIT", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_BUSINESS_UNIT"]) });
                                parameters.Add(new SqlParameter("@ID_TYPE_HAZARD", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_TYPE_HAZARD"]) });
                                parameters.Add(new SqlParameter("@ID_HEALTH_HAZARD", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_HEALTH_HAZARD"]) });
                                parameters.Add(new SqlParameter("@ID_HEALTH_EFFECT", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_HEALTH_EFFECT"]) });
                                parameters.Add(new SqlParameter("@TYPE_HAZARD", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["TYPE_HAZARD"]) });
                                parameters.Add(new SqlParameter("@HEALTH_HAZARD", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["HEALTH_HAZARD"]) });
                                parameters.Add(new SqlParameter("@HEALTH_EFFECT_RATING", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["HEALTH_EFFECT_RATING"]) });
                                parameters.Add(new SqlParameter("@NO_SUBAREAS", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["NO_SUBAREAS"]) });
                                parameters.Add(new SqlParameter("@ID_SUBAREAS", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_SUBAREAS"]) });
                                parameters.Add(new SqlParameter("@SUB_AREA", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["SUB_AREA"]) });
                                parameters.Add(new SqlParameter("@WORK_OF_TASK", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["WORK_OF_TASK"]) });
                                parameters.Add(new SqlParameter("@NO_TYPE_HAZARD", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["NO_TYPE_HAZARD"]) });
                                parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = ConvertToDBNull(row["UPDATE_BY"]) });
                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ"]) });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID"]) });
                                parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_PHA"]) });
                                #endregion
                                break;

                            case "delete":
                                #region Delete
                                sqlstr = "DELETE FROM EPHA_T_TABLE1_HAZARD " +
                                         "WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA";

                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ"]) });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID"]) });
                                parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_PHA"]) });
                                #endregion
                                break;
                        }

                        if (!string.IsNullOrEmpty(sqlstr))
                        {
                            ClassConnectionDb cls_conn = new ClassConnectionDb();
                            ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                            if (ret != "true") throw new Exception(ret);
                        }
                    }
                }
                #endregion Update Data Hazard

            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return ret;
        }
        public string set_hra_partiii(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet? dsData, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";
            List<SqlParameter> parameters = new List<SqlParameter>();
            if (dsData?.Tables.Count > 0)
            {
                ret = UpdateTasksData(user_name, conn, transaction, dsData, seq_header_now);
                if (ret != "true") { return ret; }
            }
            if (dsData?.Tables.Count > 0)
            {
                ret = UpdateWorkersData(user_name, conn, transaction, dsData, seq_header_now);
                if (ret != "true") { return ret; }
            }
            if (dsData?.Tables.Count > 0)
            {
                ret = UpdateDescriptionsData(user_name, conn, transaction, dsData, seq_header_now);
                if (ret != "true") { return ret; }
            }
            return ret;

        }
        public string UpdateTasksData(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet? dsData, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";

            try
            {
                #region Update Data Tasks
                if (dsData?.Tables["tasks"] != null)
                {
                    DataTable dt = dsData?.Tables["tasks"]?.Copy() ?? new DataTable();

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                        string sqlstr = "";
                        List<SqlParameter> parameters = new List<SqlParameter>();

                        if (!string.IsNullOrEmpty(action_type))
                        {
                            if (action_type == "insert")
                            {
                                #region Insert
                                sqlstr = "INSERT INTO EPHA_T_TABLE2_TASKS (" +
                                         "SEQ, ID, ID_PHA, NO, ID_BUSINESS_UNIT, ID_WORKER_GROUP, WORKER_GROUP, WORK_OR_TASK, NUMBERS_OF_WORKERS, TASKS_TYPE_OTHER, " +
                                         "CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY" +
                                         ") VALUES (" +
                                         "@SEQ, @ID, @ID_PHA, @NO, @ID_BUSINESS_UNIT, @ID_WORKER_GROUP, @WORKER_GROUP, @WORK_OR_TASK, @NUMBERS_OF_WORKERS, @TASKS_TYPE_OTHER, " +
                                         "GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                                parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_header_now) });
                                parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["NO"]) });
                                parameters.Add(new SqlParameter("@ID_BUSINESS_UNIT", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_BUSINESS_UNIT"]) });
                                parameters.Add(new SqlParameter("@ID_WORKER_GROUP", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_WORKER_GROUP"]) });
                                parameters.Add(new SqlParameter("@WORKER_GROUP", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["WORKER_GROUP"]) });
                                parameters.Add(new SqlParameter("@WORK_OR_TASK", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["WORK_OR_TASK"]) });
                                parameters.Add(new SqlParameter("@NUMBERS_OF_WORKERS", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["NUMBERS_OF_WORKERS"]) });
                                parameters.Add(new SqlParameter("@TASKS_TYPE_OTHER", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["TASKS_TYPE_OTHER"]) });
                                parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["CREATE_BY"]) });
                                parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });
                                #endregion
                            }
                            else if (action_type == "update")
                            {
                                #region Update
                                sqlstr = "UPDATE EPHA_T_TABLE2_TASKS SET " +
                                         "NO = @NO, ID_BUSINESS_UNIT = @ID_BUSINESS_UNIT, ID_WORKER_GROUP = @ID_WORKER_GROUP, WORKER_GROUP = @WORKER_GROUP, " +
                                         "WORK_OR_TASK = @WORK_OR_TASK, NUMBERS_OF_WORKERS = @NUMBERS_OF_WORKERS, TASKS_TYPE_OTHER = @TASKS_TYPE_OTHER, " +
                                         "UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                         "WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA";

                                parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["NO"]) });
                                parameters.Add(new SqlParameter("@ID_BUSINESS_UNIT", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_BUSINESS_UNIT"]) });
                                parameters.Add(new SqlParameter("@ID_WORKER_GROUP", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_WORKER_GROUP"]) });
                                parameters.Add(new SqlParameter("@WORKER_GROUP", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["WORKER_GROUP"]) });
                                parameters.Add(new SqlParameter("@WORK_OR_TASK", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["WORK_OR_TASK"]) });
                                parameters.Add(new SqlParameter("@NUMBERS_OF_WORKERS", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["NUMBERS_OF_WORKERS"]) });
                                parameters.Add(new SqlParameter("@TASKS_TYPE_OTHER", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["TASKS_TYPE_OTHER"]) });
                                parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });
                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                                parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_PHA"]) });
                                #endregion
                            }
                            else if (action_type == "delete")
                            {
                                #region Delete
                                sqlstr = "DELETE FROM EPHA_T_TABLE2_TASKS WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA";

                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                                parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_PHA"]) });
                                #endregion
                            }

                            if (!string.IsNullOrEmpty(sqlstr))
                            {
                                ClassConnectionDb cls_conn = new ClassConnectionDb();
                                ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                                if (ret != "true") throw new Exception(ret);
                            }
                        }
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return ret;
        }
        public string UpdateWorkersData(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet dsData, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";

            try
            {
                #region Update Data Workers
                if (dsData?.Tables["workers"] != null)
                {
                    DataTable dt = dsData.Tables["workers"].Copy() ?? new DataTable();

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                        string sqlstr = "";
                        List<SqlParameter> parameters = new List<SqlParameter>();

                        if (!string.IsNullOrEmpty(action_type))
                        {
                            if (action_type == "insert")
                            {
                                #region Insert
                                sqlstr = "INSERT INTO EPHA_T_TABLE2_WORKERS (" +
                                         "SEQ, ID, ID_PHA, ID_TASKS, NO, USER_NAME, USER_DISPLAYNAME, USER_TITLE, USER_TYPE, " +
                                         "CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY" +
                                         ") VALUES (" +
                                         "@SEQ, @ID, @ID_PHA, @ID_TASKS, @NO, @USER_NAME, @USER_DISPLAYNAME, @USER_TITLE, @USER_TYPE, " +
                                         "GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                                parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_header_now) });
                                parameters.Add(new SqlParameter("@ID_TASKS", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_TASKS"]) });
                                parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["NO"]) });
                                parameters.Add(new SqlParameter("@USER_NAME", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["USER_NAME"]) });
                                parameters.Add(new SqlParameter("@USER_DISPLAYNAME", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["USER_DISPLAYNAME"]) });
                                parameters.Add(new SqlParameter("@USER_TITLE", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["USER_TITLE"]) });
                                parameters.Add(new SqlParameter("@USER_TYPE", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["USER_TYPE"]) });
                                parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["CREATE_BY"]) });
                                parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });
                                #endregion
                            }
                            else if (action_type == "update")
                            {
                                #region Update
                                // No update case as per original comment
                                sqlstr = "UPDATE EPHA_T_TABLE2_WORKERS SET " +
                                         "NO = @NO, USER_NAME = @USER_NAME, USER_DISPLAYNAME = @USER_DISPLAYNAME, USER_TITLE = @USER_TITLE, USER_TYPE = @USER_TYPE, " +
                                         "UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                         "WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA";

                                parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["NO"]) });
                                parameters.Add(new SqlParameter("@USER_NAME", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["USER_NAME"]) });
                                parameters.Add(new SqlParameter("@USER_DISPLAYNAME", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["USER_DISPLAYNAME"]) });
                                parameters.Add(new SqlParameter("@USER_TITLE", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["USER_TITLE"]) });
                                parameters.Add(new SqlParameter("@USER_TYPE", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(dt.Rows[i]["USER_TYPE"]) });
                                parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = ConvertToDBNull(dt.Rows[i]["UPDATE_BY"]) });
                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                                parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_PHA"]) });
                                #endregion
                            }
                            else if (action_type == "delete")
                            {
                                #region Delete
                                sqlstr = "DELETE FROM EPHA_T_TABLE2_WORKERS WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA";

                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["SEQ"]) });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID"]) });
                                parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(dt.Rows[i]["ID_PHA"]) });
                                #endregion
                            }

                            if (!string.IsNullOrEmpty(sqlstr))
                            {
                                ClassConnectionDb cls_conn = new ClassConnectionDb();
                                ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                                if (ret != "true") throw new Exception(ret);
                            }
                        }
                    }
                }
                #endregion Update Data Workers
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return ret;
        }
        public string UpdateDescriptionsData(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet dsData, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";

            try
            {
                #region Update Data Descriptions
                if (dsData?.Tables["descriptions"] != null)
                {
                    DataTable dt = dsData.Tables["descriptions"].Copy() ?? new DataTable();

                    foreach (DataRow row in dt.Rows)
                    {
                        string action_type = (row["action_type"] + "").ToString();
                        string sqlstr = "";
                        List<SqlParameter> parameters = new List<SqlParameter>();

                        if (!string.IsNullOrEmpty(action_type))
                        {
                            if (action_type == "insert")
                            {
                                #region Insert
                                sqlstr = "INSERT INTO EPHA_T_TABLE2_DESCRIPTIONS (" +
                                         "SEQ, ID, ID_PHA, ID_TASKS, NO, DESCRIPTIONS, " +
                                         "CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY" +
                                         ") VALUES (" +
                                         "@SEQ, @ID, @ID_PHA, @ID_TASKS, @NO, @DESCRIPTIONS, " +
                                         "GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ"]) });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID"]) });
                                parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_header_now) });
                                parameters.Add(new SqlParameter("@ID_TASKS", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_TASKS"]) });
                                parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["NO"]) });
                                parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["DESCRIPTIONS"]) });
                                parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = ConvertToDBNull(row["CREATE_BY"]) });
                                parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = ConvertToDBNull(row["UPDATE_BY"]) });
                                #endregion
                            }
                            else if (action_type == "update")
                            {
                                #region Update
                                // No update case as per original comment
                                sqlstr = "UPDATE EPHA_T_TABLE2_DESCRIPTIONS SET " +
                                         "NO = @NO, DESCRIPTIONS = @DESCRIPTIONS, " +
                                         "UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                         "WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA";

                                parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["NO"]) });
                                parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["DESCRIPTIONS"]) });
                                parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = ConvertToDBNull(row["UPDATE_BY"]) });
                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ"]) });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID"]) });
                                parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_PHA"]) });
                                #endregion
                            }
                            else if (action_type == "delete")
                            {
                                #region Delete
                                sqlstr = "DELETE FROM EPHA_T_TABLE2_DESCRIPTIONS WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA";

                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ"]) });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID"]) });
                                parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_PHA"]) });
                                #endregion
                            }

                            if (!string.IsNullOrEmpty(sqlstr))
                            {
                                ClassConnectionDb cls_conn = new ClassConnectionDb();
                                ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                                if (ret != "true") throw new Exception(ret);
                            }
                        }
                    }

                }
                #endregion Update Data Descriptions
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return ret;
        }
        public string set_hra_partiv(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet? dsData, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";
            if (dsData?.Tables.Count > 0)
            {
                ret = UpdateWorksheetData(user_name, conn, transaction, dsData, seq_header_now);
                if (ret != "true") { return ret; }
            }

            return ret;
        }
        public string UpdateWorksheetData(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet dsData, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";

            try
            {
                #region Update Data Worksheet
                if (dsData?.Tables["worksheet"] != null)
                {
                    DataTable dt = dsData.Tables["worksheet"].Copy() ?? new DataTable();

                    foreach (DataRow row in dt.Rows)
                    {
                        string action_type = (row["action_type"] + "").ToString();
                        string sqlstr = "";
                        List<SqlParameter> parameters = new List<SqlParameter>();

                        if (!string.IsNullOrEmpty(action_type))
                        {
                            if (action_type == "insert")
                            {
                                #region Insert
                                sqlstr = "INSERT INTO EPHA_T_TABLE3_WORKSHEET (" +
                                         "SEQ, ID, ID_PHA, ROW_TYPE, NO, " +
                                         "ID_BUSINESS_UNIT, ID_ACTIVITY, ID_TASKS, ID_HAZARD, ID_FREQUENCY_LEVEL, ID_EXPOSURE_LEVEL, ID_EXPOSURE_RATING, ID_INITIAL_RISK_RATING, " +
                                         "UNIT_VALUE, FREQUENCY_LEVEL, EXPOSURE_BAND, " +
                                         "ID_STANDARD_TYPE, STANDARD_VALUE, STANDARD_UNIT, STANDARD_DESC, " +
                                         "EXPOSURE_LEVEL, EXPOSURE_RATING, INITIAL_RISK_RATING, " +
                                         "HIERARCHY_OF_CONTROL, EFFECTIVE, RESIDUAL_RISK_RATING, FK_RECOMMENDATIONS, SEQ_RECOMMENDATIONS, RECOMMENDATIONS_NO, RECOMMENDATIONS, " +
                                         "ESTIMATED_START_DATE, ESTIMATED_END_DATE, " +
                                         "RESPONDER_USER_NAME, RESPONDER_USER_DISPLAYNAME, ACTION_STATUS, " +
                                         "IMPLEMENT, RESPONDER_ACTION_TYPE, RESPONDER_ACTION_DATE, REVIEWER_ACTION_TYPE, REVIEWER_ACTION_DATE, " +
                                         "ACTION_PROJECT_TEAM, PROJECT_TEAM_TEXT, " +
                                         "CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY" +
                                         ") VALUES (" +
                                         "@SEQ, @ID, @ID_PHA, @ROW_TYPE, @NO, " +
                                         "@ID_BUSINESS_UNIT, @ID_ACTIVITY, @ID_TASKS, @ID_HAZARD, @ID_FREQUENCY_LEVEL, @ID_EXPOSURE_LEVEL, @ID_EXPOSURE_RATING, @ID_INITIAL_RISK_RATING, " +
                                         "@UNIT_VALUE, @FREQUENCY_LEVEL, @EXPOSURE_BAND, " +
                                         "@ID_STANDARD_TYPE, @STANDARD_VALUE, @STANDARD_UNIT, @STANDARD_DESC, " +
                                         "@EXPOSURE_LEVEL, @EXPOSURE_RATING, @INITIAL_RISK_RATING, " +
                                         "@HIERARCHY_OF_CONTROL, @EFFECTIVE, @RESIDUAL_RISK_RATING, @FK_RECOMMENDATIONS, @SEQ_RECOMMENDATIONS, @RECOMMENDATIONS_NO, @RECOMMENDATIONS, " +
                                         "@ESTIMATED_START_DATE, @ESTIMATED_END_DATE, " +
                                         "@RESPONDER_USER_NAME, @RESPONDER_USER_DISPLAYNAME, @ACTION_STATUS, " +
                                         "@IMPLEMENT, @RESPONDER_ACTION_TYPE, @RESPONDER_ACTION_DATE, @REVIEWER_ACTION_TYPE, @REVIEWER_ACTION_DATE, " +
                                         "@ACTION_PROJECT_TEAM, @PROJECT_TEAM_TEXT, " +
                                         "GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ"]) });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID"]) });
                                parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_header_now) });
                                parameters.Add(new SqlParameter("@ROW_TYPE", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["ROW_TYPE"]) });
                                parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["NO"]) });
                                parameters.Add(new SqlParameter("@ID_BUSINESS_UNIT", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_BUSINESS_UNIT"]) });
                                parameters.Add(new SqlParameter("@ID_ACTIVITY", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_ACTIVITY"]) });
                                parameters.Add(new SqlParameter("@ID_TASKS", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_TASKS"]) });
                                parameters.Add(new SqlParameter("@ID_HAZARD", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_HAZARD"]) });
                                parameters.Add(new SqlParameter("@ID_FREQUENCY_LEVEL", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_FREQUENCY_LEVEL"]) });
                                parameters.Add(new SqlParameter("@ID_EXPOSURE_LEVEL", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_EXPOSURE_LEVEL"]) });
                                parameters.Add(new SqlParameter("@ID_EXPOSURE_RATING", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_EXPOSURE_RATING"]) });
                                parameters.Add(new SqlParameter("@ID_INITIAL_RISK_RATING", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_INITIAL_RISK_RATING"]) });
                                parameters.Add(new SqlParameter("@UNIT_VALUE", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["UNIT_VALUE"]) });
                                parameters.Add(new SqlParameter("@FREQUENCY_LEVEL", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["FREQUENCY_LEVEL"]) });
                                parameters.Add(new SqlParameter("@EXPOSURE_BAND", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["EXPOSURE_BAND"]) });
                                parameters.Add(new SqlParameter("@ID_STANDARD_TYPE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_STANDARD_TYPE"]) });
                                parameters.Add(new SqlParameter("@STANDARD_VALUE", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["STANDARD_VALUE"]) });
                                parameters.Add(new SqlParameter("@STANDARD_UNIT", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["STANDARD_UNIT"]) });
                                parameters.Add(new SqlParameter("@STANDARD_DESC", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["STANDARD_DESC"]) });
                                parameters.Add(new SqlParameter("@EXPOSURE_LEVEL", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["EXPOSURE_LEVEL"]) });
                                parameters.Add(new SqlParameter("@EXPOSURE_RATING", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["EXPOSURE_RATING"]) });
                                parameters.Add(new SqlParameter("@INITIAL_RISK_RATING", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["INITIAL_RISK_RATING"]) });
                                parameters.Add(new SqlParameter("@HIERARCHY_OF_CONTROL", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["HIERARCHY_OF_CONTROL"]) });
                                parameters.Add(new SqlParameter("@EFFECTIVE", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["EFFECTIVE"]) });
                                parameters.Add(new SqlParameter("@RESIDUAL_RISK_RATING", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["RESIDUAL_RISK_RATING"]) });
                                parameters.Add(new SqlParameter("@FK_RECOMMENDATIONS", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["FK_RECOMMENDATIONS"]) });
                                parameters.Add(new SqlParameter("@SEQ_RECOMMENDATIONS", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ_RECOMMENDATIONS"]) });
                                parameters.Add(new SqlParameter("@RECOMMENDATIONS_NO", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["RECOMMENDATIONS_NO"]) });
                                parameters.Add(new SqlParameter("@RECOMMENDATIONS", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["RECOMMENDATIONS"]) });
                                parameters.Add(new SqlParameter("@ESTIMATED_START_DATE", SqlDbType.Date) { Value = ConvertToDBNull(row["ESTIMATED_START_DATE"]) });
                                parameters.Add(new SqlParameter("@ESTIMATED_END_DATE", SqlDbType.Date) { Value = ConvertToDBNull(row["ESTIMATED_END_DATE"]) });
                                parameters.Add(new SqlParameter("@RESPONDER_USER_NAME", SqlDbType.NVarChar, 50) { Value = ConvertToDBNull(row["RESPONDER_USER_NAME"]) });
                                parameters.Add(new SqlParameter("@RESPONDER_USER_DISPLAYNAME", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["RESPONDER_USER_DISPLAYNAME"]) });
                                parameters.Add(new SqlParameter("@ACTION_STATUS", SqlDbType.NVarChar, 50) { Value = ConvertToDBNull(row["ACTION_STATUS"]) });
                                parameters.Add(new SqlParameter("@IMPLEMENT", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["IMPLEMENT"]) });
                                parameters.Add(new SqlParameter("@RESPONDER_ACTION_TYPE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["RESPONDER_ACTION_TYPE"]) });
                                parameters.Add(new SqlParameter("@RESPONDER_ACTION_DATE", SqlDbType.Date) { Value = ConvertToDBNull(row["RESPONDER_ACTION_DATE"]) });
                                parameters.Add(new SqlParameter("@REVIEWER_ACTION_TYPE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["REVIEWER_ACTION_TYPE"]) });
                                parameters.Add(new SqlParameter("@REVIEWER_ACTION_DATE", SqlDbType.Date) { Value = ConvertToDBNull(row["REVIEWER_ACTION_DATE"]) });
                                parameters.Add(new SqlParameter("@ACTION_PROJECT_TEAM", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ACTION_PROJECT_TEAM"]) });
                                parameters.Add(new SqlParameter("@PROJECT_TEAM_TEXT", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["PROJECT_TEAM_TEXT"]) });
                                parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = ConvertToDBNull(row["CREATE_BY"]) });
                                parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = ConvertToDBNull(row["UPDATE_BY"]) });
                                #endregion
                            }
                            else if (action_type == "update")
                            {
                                #region Update
                                sqlstr = "UPDATE EPHA_T_TABLE3_WORKSHEET SET " +
                                         "NO = @NO, ID_ACTIVITY = @ID_ACTIVITY, ID_FREQUENCY_LEVEL = @ID_FREQUENCY_LEVEL, ID_EXPOSURE_LEVEL = @ID_EXPOSURE_LEVEL, " +
                                         "ID_EXPOSURE_RATING = @ID_EXPOSURE_RATING, ID_INITIAL_RISK_RATING = @ID_INITIAL_RISK_RATING, " +
                                         "UNIT_VALUE = @UNIT_VALUE, FREQUENCY_LEVEL = @FREQUENCY_LEVEL, EXPOSURE_BAND = @EXPOSURE_BAND, " +
                                         "ID_STANDARD_TYPE = @ID_STANDARD_TYPE, STANDARD_VALUE = @STANDARD_VALUE, STANDARD_UNIT = @STANDARD_UNIT, STANDARD_DESC = @STANDARD_DESC, " +
                                         "EXPOSURE_LEVEL = @EXPOSURE_LEVEL, EXPOSURE_RATING = @EXPOSURE_RATING, INITIAL_RISK_RATING = @INITIAL_RISK_RATING, " +
                                         "HIERARCHY_OF_CONTROL = @HIERARCHY_OF_CONTROL, EFFECTIVE = @EFFECTIVE, RESIDUAL_RISK_RATING = @RESIDUAL_RISK_RATING, " +
                                         "FK_RECOMMENDATIONS = @FK_RECOMMENDATIONS, SEQ_RECOMMENDATIONS = @SEQ_RECOMMENDATIONS, RECOMMENDATIONS_NO = @RECOMMENDATIONS_NO, RECOMMENDATIONS = @RECOMMENDATIONS, " +
                                         "ESTIMATED_START_DATE = @ESTIMATED_START_DATE, ESTIMATED_END_DATE = @ESTIMATED_END_DATE, " +
                                         "RESPONDER_USER_NAME = @RESPONDER_USER_NAME, RESPONDER_USER_DISPLAYNAME = @RESPONDER_USER_DISPLAYNAME, ACTION_STATUS = @ACTION_STATUS, " +
                                         "IMPLEMENT = @IMPLEMENT, RESPONDER_ACTION_TYPE = @RESPONDER_ACTION_TYPE, RESPONDER_ACTION_DATE = @RESPONDER_ACTION_DATE, " +
                                         "REVIEWER_ACTION_TYPE = @REVIEWER_ACTION_TYPE, REVIEWER_ACTION_DATE = @REVIEWER_ACTION_DATE, " +
                                         "ACTION_PROJECT_TEAM = @ACTION_PROJECT_TEAM, PROJECT_TEAM_TEXT = @PROJECT_TEAM_TEXT, " +
                                         "UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                         "WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA";

                                parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["NO"]) });
                                parameters.Add(new SqlParameter("@ID_ACTIVITY", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_ACTIVITY"]) });
                                parameters.Add(new SqlParameter("@ID_FREQUENCY_LEVEL", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_FREQUENCY_LEVEL"]) });
                                parameters.Add(new SqlParameter("@ID_EXPOSURE_LEVEL", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_EXPOSURE_LEVEL"]) });
                                parameters.Add(new SqlParameter("@ID_EXPOSURE_RATING", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_EXPOSURE_RATING"]) });
                                parameters.Add(new SqlParameter("@ID_INITIAL_RISK_RATING", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_INITIAL_RISK_RATING"]) });
                                parameters.Add(new SqlParameter("@UNIT_VALUE", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["UNIT_VALUE"]) });
                                parameters.Add(new SqlParameter("@FREQUENCY_LEVEL", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["FREQUENCY_LEVEL"]) });
                                parameters.Add(new SqlParameter("@EXPOSURE_BAND", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["EXPOSURE_BAND"]) });
                                parameters.Add(new SqlParameter("@ID_STANDARD_TYPE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_STANDARD_TYPE"]) });
                                parameters.Add(new SqlParameter("@STANDARD_VALUE", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["STANDARD_VALUE"]) });
                                parameters.Add(new SqlParameter("@STANDARD_UNIT", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["STANDARD_UNIT"]) });
                                parameters.Add(new SqlParameter("@STANDARD_DESC", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["STANDARD_DESC"]) });
                                parameters.Add(new SqlParameter("@EXPOSURE_LEVEL", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["EXPOSURE_LEVEL"]) });
                                parameters.Add(new SqlParameter("@EXPOSURE_RATING", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["EXPOSURE_RATING"]) });
                                parameters.Add(new SqlParameter("@INITIAL_RISK_RATING", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["INITIAL_RISK_RATING"]) });
                                parameters.Add(new SqlParameter("@HIERARCHY_OF_CONTROL", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["HIERARCHY_OF_CONTROL"]) });
                                parameters.Add(new SqlParameter("@EFFECTIVE", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["EFFECTIVE"]) });
                                parameters.Add(new SqlParameter("@RESIDUAL_RISK_RATING", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["RESIDUAL_RISK_RATING"]) });
                                parameters.Add(new SqlParameter("@FK_RECOMMENDATIONS", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["FK_RECOMMENDATIONS"]) });
                                parameters.Add(new SqlParameter("@SEQ_RECOMMENDATIONS", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ_RECOMMENDATIONS"]) });
                                parameters.Add(new SqlParameter("@RECOMMENDATIONS_NO", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["RECOMMENDATIONS_NO"]) });
                                parameters.Add(new SqlParameter("@RECOMMENDATIONS", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["RECOMMENDATIONS"]) });
                                parameters.Add(new SqlParameter("@ESTIMATED_START_DATE", SqlDbType.Date) { Value = ConvertToDBNull(row["ESTIMATED_START_DATE"]) });
                                parameters.Add(new SqlParameter("@ESTIMATED_END_DATE", SqlDbType.Date) { Value = ConvertToDBNull(row["ESTIMATED_END_DATE"]) });
                                parameters.Add(new SqlParameter("@RESPONDER_USER_NAME", SqlDbType.NVarChar, 50) { Value = ConvertToDBNull(row["RESPONDER_USER_NAME"]) });
                                parameters.Add(new SqlParameter("@RESPONDER_USER_DISPLAYNAME", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["RESPONDER_USER_DISPLAYNAME"]) });
                                parameters.Add(new SqlParameter("@ACTION_STATUS", SqlDbType.NVarChar, 50) { Value = ConvertToDBNull(row["ACTION_STATUS"]) });
                                parameters.Add(new SqlParameter("@IMPLEMENT", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["IMPLEMENT"]) });
                                parameters.Add(new SqlParameter("@RESPONDER_ACTION_TYPE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["RESPONDER_ACTION_TYPE"]) });
                                parameters.Add(new SqlParameter("@RESPONDER_ACTION_DATE", SqlDbType.Date) { Value = ConvertToDBNull(row["RESPONDER_ACTION_DATE"]) });
                                parameters.Add(new SqlParameter("@REVIEWER_ACTION_TYPE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["REVIEWER_ACTION_TYPE"]) });
                                parameters.Add(new SqlParameter("@REVIEWER_ACTION_DATE", SqlDbType.Date) { Value = ConvertToDBNull(row["REVIEWER_ACTION_DATE"]) });
                                parameters.Add(new SqlParameter("@ACTION_PROJECT_TEAM", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ACTION_PROJECT_TEAM"]) });
                                parameters.Add(new SqlParameter("@PROJECT_TEAM_TEXT", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["PROJECT_TEAM_TEXT"]) });
                                parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = ConvertToDBNull(row["UPDATE_BY"]) });
                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ"]) });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID"]) });
                                parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_PHA"]) });
                                #endregion
                            }
                            else if (action_type == "delete")
                            {
                                #region Delete
                                sqlstr = "DELETE FROM EPHA_T_TABLE3_WORKSHEET WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA";

                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ"]) });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID"]) });
                                parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_PHA"]) });
                                #endregion
                            }

                            if (!string.IsNullOrEmpty(sqlstr))
                            {
                                ClassConnectionDb cls_conn = new ClassConnectionDb();
                                ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                                if (ret != "true") throw new Exception(ret);
                            }
                        }
                    }
                }
                #endregion Update Data Worksheet
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return ret;
        }
        public string set_approver(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet? dsData)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";
            if (dsData?.Tables.Count > 0)
            {
                ret = UpdateApproverData(user_name, conn, transaction, dsData);
                if (ret != "true") { return ret; }
            }
            return ret;

        }
        public string UpdateApproverData(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet dsData)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";

            try
            {
                #region Update Data Approver
                if (dsData.Tables["approver"] != null)
                {
                    DataTable dt = dsData?.Tables["approver"]?.Copy() ?? new DataTable();
                    dt.AcceptChanges();

                    foreach (DataRow row in dt.Rows)
                    {
                        string id_approver = row["ID"]?.ToString() ?? string.Empty;
                        string action_type = row["action_type"]?.ToString() ?? string.Empty;
                        string sqlstr = "";
                        List<SqlParameter> parameters = new List<SqlParameter>();

                        if (action_type == "update" || action_type == "delete")
                        {
                            sqlstr = "DELETE FROM EPHA_T_APPROVER_TA3 WHERE ID_PHA = @ID_PHA AND ID_SESSION = @ID_SESSION AND ID_APPROVER = @ID_APPROVER";
                            parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = row["ID_PHA"] });
                            parameters.Add(new SqlParameter("@ID_SESSION", SqlDbType.Int) { Value = row["ID_SESSION"] });
                            parameters.Add(new SqlParameter("@ID_APPROVER", SqlDbType.Int) { Value = id_approver });

                            if (!string.IsNullOrEmpty(sqlstr))
                            {
                                ClassConnectionDb cls_conn = new ClassConnectionDb();
                                ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                                if (ret != "true") throw new Exception(ret);
                            }
                        }

                        if (action_type == "insert" || action_type == "update")
                        {
                            sqlstr = "INSERT INTO EPHA_T_APPROVER_TA3 (ID_PHA, ID_SESSION, ID_APPROVER) VALUES (@ID_PHA, @ID_SESSION, @ID_APPROVER)";
                            parameters = new List<SqlParameter>();
                            parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = row["ID_PHA"] });
                            parameters.Add(new SqlParameter("@ID_SESSION", SqlDbType.Int) { Value = row["ID_SESSION"] });
                            parameters.Add(new SqlParameter("@ID_APPROVER", SqlDbType.Int) { Value = id_approver });

                            if (!string.IsNullOrEmpty(sqlstr))
                            {
                                ClassConnectionDb cls_conn = new ClassConnectionDb();
                                ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                                if (ret != "true") throw new Exception(ret);
                            }
                        }
                    }
                }
                #endregion Update Data Approver
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }
            return ret;
        }

        public string set_recommendations_part(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet? dsData, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";
            if (dsData?.Tables.Count > 0)
            {
                ret = UpdateRecommendationsData(user_name, conn, transaction, dsData, seq_header_now);
                if (ret != "true") { return ret; }
            }
            return ret;

        }
        public string UpdateRecommendationsData(string user_name, SqlConnection conn, SqlTransaction transaction, DataSet dsData, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";

            try
            {
                #region Update Data Recommendations
                if (dsData?.Tables["recommendations"] != null)
                {
                    DataTable dt = dsData?.Tables["recommendations"]?.Copy() ?? new DataTable();

                    foreach (DataRow row in dt.Rows)
                    {
                        string action_type = (row["action_type"] + "").ToString();
                        string sqlstr = "";
                        List<SqlParameter> parameters = new List<SqlParameter>();

                        if (!string.IsNullOrEmpty(action_type))
                        {
                            if (action_type == "insert")
                            {
                                #region Insert
                                sqlstr = "INSERT INTO EPHA_T_RECOMMENDATIONS (" +
                                         "SEQ, ID, ID_PHA, ID_WORKSHEET, NO, RECOMMENDATIONS, " +
                                         "CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY" +
                                         ") VALUES (" +
                                         "@SEQ, @ID, @ID_PHA, @ID_WORKSHEET, @NO, @RECOMMENDATIONS, " +
                                         "GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ"]) });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID"]) });
                                parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_header_now) });
                                parameters.Add(new SqlParameter("@ID_WORKSHEET", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_WORKSHEET"]) });
                                parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["NO"]) });
                                parameters.Add(new SqlParameter("@RECOMMENDATIONS", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["RECOMMENDATIONS"]) });
                                parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = ConvertToDBNull(row["CREATE_BY"]) });
                                parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = ConvertToDBNull(row["UPDATE_BY"]) });
                                #endregion
                            }
                            else if (action_type == "update")
                            {
                                #region Update
                                sqlstr = "UPDATE EPHA_T_RECOMMENDATIONS SET " +
                                         "NO = @NO, RECOMMENDATIONS = @RECOMMENDATIONS, " +
                                         "UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                         "WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA AND ID_WORKSHEET = @ID_WORKSHEET";

                                parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["NO"]) });
                                parameters.Add(new SqlParameter("@RECOMMENDATIONS", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["RECOMMENDATIONS"]) });
                                parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = ConvertToDBNull(row["UPDATE_BY"]) });
                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ"]) });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID"]) });
                                parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_PHA"]) });
                                parameters.Add(new SqlParameter("@ID_WORKSHEET", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_WORKSHEET"]) });
                                #endregion
                            }
                            else if (action_type == "delete")
                            {
                                #region Delete
                                sqlstr = "DELETE FROM EPHA_T_RECOMMENDATIONS WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA AND ID_WORKSHEET = @ID_WORKSHEET";

                                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ"]) });
                                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID"]) });
                                parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_PHA"]) });
                                parameters.Add(new SqlParameter("@ID_WORKSHEET", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_WORKSHEET"]) });
                                #endregion
                            }

                            if (!string.IsNullOrEmpty(sqlstr))
                            {
                                ClassConnectionDb cls_conn = new ClassConnectionDb();
                                ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                                if (ret != "true") throw new Exception(ret);
                            }
                        }
                    }
                }
                #endregion Update Data Recommendations
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return ret;
        }


        #endregion set page worksheet details

        #region set page master ram
        public string set_ram_level(string user_name, SqlConnection conn, SqlTransaction transaction, DataTable _dtDef, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";
            if (_dtDef?.Rows.Count > 0)
            {
                ret = UpdateRamLevelData(user_name, conn, transaction, _dtDef);
                if (ret != "true") { return ret; }
            }
            return ret;
        }
        public string UpdateRamLevelData(string user_name, SqlConnection conn, SqlTransaction transaction, DataTable _dtDef)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";

            try
            {
                #region Update Data Ram Level
                DataTable dt = _dtDef?.Copy() ?? new DataTable();

                foreach (DataRow row in dt.Rows)
                {
                    string action_type = (row["action_type"] + "").ToString();
                    string sqlstr = "";
                    List<SqlParameter> parameters = new List<SqlParameter>();

                    if (action_type == "update")
                    {
                        #region Update
                        sqlstr = "UPDATE EPHA_M_RAM_LEVEL SET " +
                                 "people = @people, assets = @assets, enhancement = @enhancement, reputation = @reputation, " +
                                 "product_quality = @product_quality, security_level = @security_level, " +
                                 "opportunity_level = @opportunity_level, opportunity_desc = @opportunity_desc, security_text = @security_text, " +
                                 "UPDATE_DATE = GETDATE()";

                        parameters.Add(new SqlParameter("@people", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["people"]) });
                        parameters.Add(new SqlParameter("@assets", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["assets"]) });
                        parameters.Add(new SqlParameter("@enhancement", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["enhancement"]) });
                        parameters.Add(new SqlParameter("@reputation", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["reputation"]) });
                        parameters.Add(new SqlParameter("@product_quality", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["product_quality"]) });
                        parameters.Add(new SqlParameter("@security_level", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["security_level"]) });

                        for (int c = 1; c < 11; c++)
                        {
                            sqlstr += $", likelihood{c}_level = @likelihood{c}_level, likelihood{c}_text = @likelihood{c}_text, " +
                                      $"likelihood{c}_desc = @likelihood{c}_desc, likelihood{c}_criterion = @likelihood{c}_criterion, " +
                                      $"ram{c}_text = @ram{c}_text, ram{c}_priority = @ram{c}_priority, ram{c}_desc = @ram{c}_desc, ram{c}_color = @ram{c}_color";

                            parameters.Add(new SqlParameter($"@likelihood{c}_level", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row[$"likelihood{c}_level"]) });
                            parameters.Add(new SqlParameter($"@likelihood{c}_text", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row[$"likelihood{c}_text"]) });
                            parameters.Add(new SqlParameter($"@likelihood{c}_desc", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row[$"likelihood{c}_desc"]) });
                            parameters.Add(new SqlParameter($"@likelihood{c}_criterion", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row[$"likelihood{c}_criterion"]) });
                            parameters.Add(new SqlParameter($"@ram{c}_text", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row[$"ram{c}_text"]) });
                            parameters.Add(new SqlParameter($"@ram{c}_priority", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row[$"ram{c}_priority"]) });
                            parameters.Add(new SqlParameter($"@ram{c}_desc", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row[$"ram{c}_desc"]) });
                            parameters.Add(new SqlParameter($"@ram{c}_color", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row[$"ram{c}_color"]) });
                        }

                        parameters.Add(new SqlParameter("@opportunity_level", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["opportunity_level"]) });
                        parameters.Add(new SqlParameter("@opportunity_desc", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["opportunity_desc"]) });
                        parameters.Add(new SqlParameter("@security_text", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["security_text"]) });

                        sqlstr += " WHERE SEQ = @SEQ AND ID_RAM = @ID_RAM";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID_RAM", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_RAM"]) });
                        #endregion
                    }

                    if (!string.IsNullOrEmpty(sqlstr))
                    {
                        ClassConnectionDb cls_conn = new ClassConnectionDb();
                        ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                        if (ret != "true") throw new Exception(ret);
                    }
                }
                #endregion Update Data Ram Level
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return ret;
        }

        public string set_ram_master(string user_name, SqlConnection conn, SqlTransaction transaction, DataTable _dtDef, string seq_header_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ  
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";
            if (_dtDef?.Rows.Count > 0)
            {
                ret = UpdateRamData(user_name, conn, transaction, _dtDef);
                if (ret != "true") { return ret; }
            }
            return ret;
        }
        public string UpdateRamData(string user_name, SqlConnection conn, SqlTransaction transaction, DataTable _dtDef)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";

            try
            {
                #region Update Data Ram
                DataTable dt = _dtDef?.Copy() ?? new DataTable();

                foreach (DataRow row in dt.Rows)
                {
                    string action_type = (row["action_type"] + "").ToString();
                    string sqlstr = "";
                    List<SqlParameter> parameters = new List<SqlParameter>();

                    if (action_type == "update")
                    {
                        #region Update
                        sqlstr = "UPDATE EPHA_M_RAM SET " +
                                 "DOCUMENT_FILE_NAME = @DOCUMENT_FILE_NAME, " +
                                 "DOCUMENT_FILE_PATH = @DOCUMENT_FILE_PATH, " +
                                 "ROWS_LEVEL = @ROWS_LEVEL, " +
                                 "COLUMNS_LEVEL = @COLUMNS_LEVEL, " +
                                 "UPDATE_DATE = GETDATE()";

                        parameters.Add(new SqlParameter("@DOCUMENT_FILE_NAME", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["DOCUMENT_FILE_NAME"]) });
                        parameters.Add(new SqlParameter("@DOCUMENT_FILE_PATH", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["DOCUMENT_FILE_PATH"]) });
                        parameters.Add(new SqlParameter("@ROWS_LEVEL", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ROWS_LEVEL"]) });
                        parameters.Add(new SqlParameter("@COLUMNS_LEVEL", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["COLUMNS_LEVEL"]) });

                        sqlstr += " WHERE SEQ = @SEQ AND ID_RAM = @ID_RAM";

                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ"]) });
                        parameters.Add(new SqlParameter("@ID_RAM", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ID_RAM"]) });
                        #endregion
                    }

                    if (!string.IsNullOrEmpty(sqlstr))
                    {
                        ClassConnectionDb cls_conn = new ClassConnectionDb();
                        ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                        if (ret != "true") throw new Exception(ret);
                    }
                }
                #endregion Update Data Ram
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return ret;
        }

        public string set_master_ram(SetDataWorkflowModel param)
        {
            string msg = string.Empty;
            string ret = "true";
            ClassJSON cls_json = new ClassJSON();
            DataTable dt = new DataTable();
            DataSet dsData = new DataSet();
            string user_name = param.user_name ?? string.Empty;
            string jsper = param.json_ram_master ?? string.Empty;

            // ตรวจสอบสิทธิ์ก่อนดำเนินการ string user_name, 
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            try
            {
                if (!string.IsNullOrWhiteSpace(jsper))
                {
                    dt = cls_json.ConvertJSONresult(jsper);
                    if (dt != null)
                    {
                        dt.TableName = "ram_master";
                        dsData.Tables.Add(dt.Copy());
                        dsData.AcceptChanges();
                        ret = string.Empty;
                    }
                }
            }
            catch (Exception ex)
            {
                msg = ex.Message;
                ret = "Error";
            }
            if (ret.Equals("Error", StringComparison.OrdinalIgnoreCase))
            {
                goto Next_Line_Convert;
            }

            #region connection transaction
            int seq_now = 0;
            int seq_level_now = 0;

            string sqlstr = "SELECT MAX(a.seq) + 1 AS max_seq FROM EPHA_M_RAM a";
            List<SqlParameter> parameters = new List<SqlParameter>();

            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
            if (dt?.Rows.Count > 0)
            {
                seq_now = Convert.ToInt32(dt.Rows[0]["max_seq"]);
            }

            sqlstr = "SELECT MAX(a.seq) + 1 AS max_seq FROM EPHA_M_RAM_LEVEL a";
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
            if (dt?.Rows.Count > 0)
            {
                seq_level_now = Convert.ToInt32(dt.Rows[0]["max_seq"]);
            }

            #endregion connection transaction

            ret = UpdateRamDataAndGenerateRamLevel(user_name, dsData, seq_now, seq_level_now);

        Next_Line_Convert:


            dt = new DataTable();
            dt = ClassFile.refMsg(ret, msg);
            if (dt != null) { dsData = new DataSet(); dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges(); }


            if (ret.Equals("true", StringComparison.OrdinalIgnoreCase))
            {
                ClassHazop cls = new ClassHazop();
                cls.get_master_ram(ref dsData);
            }

            string json = Newtonsoft.Json.JsonConvert.SerializeObject(dsData, Newtonsoft.Json.Formatting.Indented);
            return json;
        }
        public string UpdateRamDataAndGenerateRamLevel(string user_name, DataSet dsData, int seq_now, int seq_level_now)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";
            using (var conn = new SqlConnection(ClassConnectionDb.ConnectionString()))
            {
                conn.Open();
                using (var transaction = conn.BeginTransaction())
                {
                    try
                    {
                        #region Update Data Ram
                        DataTable dtRamMaster = dsData?.Tables["ram_master"]?.Copy() ?? new DataTable();

                        foreach (DataRow row in dtRamMaster.Rows)
                        {
                            string action_type = (row["action_type"] + "").ToString();
                            if (!string.IsNullOrEmpty(action_type))
                            {
                                string sqlstr = "";
                                List<SqlParameter> parameters = new List<SqlParameter>();

                                if (action_type == "insert")
                                {
                                    #region Insert
                                    sqlstr = "INSERT INTO EPHA_M_RAM (" +
                                             "SEQ, ID, NAME, DESCRIPTIONS, ACTIVE_TYPE, CATEGORY_TYPE, DOCUMENT_FILE_NAME, DOCUMENT_FILE_PATH, DOCUMENT_FILE_SIZE, ROWS_LEVEL, COLUMNS_LEVEL, " +
                                             "CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY" +
                                             ") VALUES (" +
                                             "@SEQ, @ID, @NAME, @DESCRIPTIONS, @ACTIVE_TYPE, @CATEGORY_TYPE, @DOCUMENT_FILE_NAME, @DOCUMENT_FILE_PATH, @DOCUMENT_FILE_SIZE, @ROWS_LEVEL, @COLUMNS_LEVEL, " +
                                             "GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_now) });
                                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_now) });
                                    parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["NAME"]) });
                                    parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["DESCRIPTIONS"]) });
                                    parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ACTIVE_TYPE"]) });
                                    parameters.Add(new SqlParameter("@CATEGORY_TYPE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["CATEGORY_TYPE"]) });
                                    parameters.Add(new SqlParameter("@DOCUMENT_FILE_NAME", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["DOCUMENT_FILE_NAME"]) });
                                    parameters.Add(new SqlParameter("@DOCUMENT_FILE_PATH", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["DOCUMENT_FILE_PATH"]) });
                                    //parameters.Add(new SqlParameter("@DOCUMENT_FILE_SIZE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["DOCUMENT_FILE_SIZE"]) });
                                    parameters.Add(new SqlParameter("@DOCUMENT_FILE_SIZE", SqlDbType.Int)
                                    {
                                        Value = row.Table.Columns.Contains("DOCUMENT_FILE_SIZE") ? ConvertToDBNull(row["DOCUMENT_FILE_SIZE"]) : DBNull.Value
                                    });
                                    parameters.Add(new SqlParameter("@ROWS_LEVEL", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ROWS_LEVEL"]) });
                                    parameters.Add(new SqlParameter("@COLUMNS_LEVEL", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["COLUMNS_LEVEL"]) });
                                    parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = ConvertToDBNull(row["CREATE_BY"]) });
                                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = ConvertToDBNull(row["UPDATE_BY"]) });
                                    #endregion
                                }
                                else if (action_type == "update")
                                {
                                    #region Update
                                    sqlstr = "UPDATE EPHA_M_RAM SET " +
                                             "NAME = @NAME, DESCRIPTIONS = @DESCRIPTIONS, ACTIVE_TYPE = @ACTIVE_TYPE, CATEGORY_TYPE = @CATEGORY_TYPE, " +
                                             "DOCUMENT_FILE_NAME = @DOCUMENT_FILE_NAME, DOCUMENT_FILE_PATH = @DOCUMENT_FILE_PATH, DOCUMENT_FILE_SIZE = @DOCUMENT_FILE_SIZE, " +
                                             "ROWS_LEVEL = @ROWS_LEVEL, COLUMNS_LEVEL = @COLUMNS_LEVEL, " +
                                             "UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY " +
                                             "WHERE SEQ = @SEQ";

                                    parameters.Add(new SqlParameter("@NAME", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["NAME"]) });
                                    parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["DESCRIPTIONS"]) });
                                    parameters.Add(new SqlParameter("@ACTIVE_TYPE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ACTIVE_TYPE"]) });
                                    parameters.Add(new SqlParameter("@CATEGORY_TYPE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["CATEGORY_TYPE"]) });
                                    parameters.Add(new SqlParameter("@DOCUMENT_FILE_NAME", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["DOCUMENT_FILE_NAME"]) });
                                    parameters.Add(new SqlParameter("@DOCUMENT_FILE_PATH", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(row["DOCUMENT_FILE_PATH"]) });
                                    //parameters.Add(new SqlParameter("@DOCUMENT_FILE_SIZE", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["DOCUMENT_FILE_SIZE"]) });
                                    parameters.Add(new SqlParameter("@DOCUMENT_FILE_SIZE", SqlDbType.Int)
                                    {
                                        Value = row.Table.Columns.Contains("DOCUMENT_FILE_SIZE") ? ConvertToDBNull(row["DOCUMENT_FILE_SIZE"]) : DBNull.Value
                                    });
                                    parameters.Add(new SqlParameter("@ROWS_LEVEL", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["ROWS_LEVEL"]) });
                                    parameters.Add(new SqlParameter("@COLUMNS_LEVEL", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["COLUMNS_LEVEL"]) });
                                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = ConvertToDBNull(row["UPDATE_BY"]) });
                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ"]) });
                                    #endregion
                                }
                                else if (action_type == "delete")
                                {
                                    #region Delete
                                    sqlstr = "DELETE FROM EPHA_M_RAM WHERE SEQ = @SEQ";

                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(row["SEQ"]) });
                                    #endregion
                                }

                                if (!string.IsNullOrEmpty(sqlstr))
                                {
                                    ClassConnectionDb cls_conn = new ClassConnectionDb();
                                    ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                                    if (ret != "true") throw new Exception(ret);
                                }
                            }
                        }
                        #endregion Update Data Ram

                        #region Generate Ram Level
                        DataTable dtRamLevel = dsData?.Tables["ram_level"]?.Copy() ?? new DataTable();

                        foreach (DataRow row in dtRamLevel.Rows)
                        {
                            string action_type = (row["action_type"] + "").ToString();
                            if (!string.IsNullOrEmpty(action_type) && action_type == "insert")
                            {
                                int rows_level = Convert.ToInt32((row["rows_level"] + "").ToString());
                                int columns_level = Convert.ToInt32((row["columns_level"] + "").ToString());
                                for (int ir = 0; ir < rows_level; ir++)
                                {
                                    string sqlstr = "INSERT INTO EPHA_M_RAM_LEVEL (SEQ, ID, ID_RAM, SORT_BY, SECURITY_LEVEL";
                                    for (int ic = 1; ic <= columns_level; ic++)
                                    {
                                        sqlstr += $", LIKELIHOOD{ic}_TEXT";
                                    }
                                    sqlstr += ", CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) VALUES (" +
                                              "@SEQ, @ID, @ID_RAM, @SORT_BY, @SECURITY_LEVEL";

                                    List<SqlParameter> parameters = new List<SqlParameter>
                            {
                                new SqlParameter("@SEQ", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_level_now) },
                                new SqlParameter("@ID", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_level_now) },
                                new SqlParameter("@ID_RAM", SqlDbType.Int) { Value = ConvertToIntOrDBNull(seq_now) },
                                new SqlParameter("@SORT_BY", SqlDbType.Int) { Value = ConvertToIntOrDBNull(ir + 1) },
                                new SqlParameter("@SECURITY_LEVEL", SqlDbType.Int) { Value = ConvertToIntOrDBNull(rows_level - ir) }
                            };

                                    for (int ic = 1; ic <= columns_level; ic++)
                                    {
                                        parameters.Add(new SqlParameter($"@LIKELIHOOD{ic}_TEXT", SqlDbType.NVarChar, 4000) { Value = ConvertToDBNull(ic) });
                                    }

                                    sqlstr += ", GETDATE(), NULL, @CREATE_BY, @UPDATE_BY)";

                                    parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = ConvertToDBNull(row["CREATE_BY"]) });
                                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = ConvertToDBNull(row["UPDATE_BY"]) });

                                    ClassConnectionDb cls_conn = new ClassConnectionDb();
                                    ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                                    if (ret != "true") throw new Exception(ret);
                                }
                            }
                        }
                        #endregion Generate Ram Level

                        if (ret == "true")
                        {
                            transaction.Commit();
                        }
                        else
                        {
                            transaction.Rollback();
                            ret = "error";
                        }
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
        #endregion set page master ram



        #region set page worksheet
        public string FlowActionSubmit(string user_name,
         string phaSubSoftware, string flowAction, string expenseType, string subExpenseType,
         string seqHeaderNow, string phaNoNow, string versionNow, string phaStatusNow,
         string phaStatus, string requestApprover)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ string user_name, 
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string phaStatusNew = phaStatus ?? "";
            if (string.IsNullOrEmpty(phaStatusNew))
            {
                return "No Data PHA Status.";
            }

            string ret = "true";

            if (phaSubSoftware == "hra")
            {
                expenseType = expenseType.ToLower() == "moc" ? "opex" : "capex";
            }

            #region flow action submit

            ClassEmail clsmail = new ClassEmail();

            if ((flowAction == "submit" || flowAction == "submit_without") && subExpenseType == "Normal")
            {
                if (phaStatus == "11")
                {
                    //12 WP PHA Conduct
                    phaStatusNew = "12";

                    // Task Register Revision += 1
                    // Save and keep current revision, then copy to new revision
                    ret = update_status_table_now(user_name, seqHeaderNow, phaNoNow, phaStatusNew);
                    ret = update_revision_table_now(user_name, seqHeaderNow, phaNoNow, versionNow, phaStatusNew, string.Empty, phaSubSoftware);

                    keep_version(user_name, ref seqHeaderNow, ref versionNow, phaStatusNew, phaSubSoftware, false, false, false, false);

                    if (flowAction != "submit_without")
                    {
                        clsmail = new ClassEmail();
                        clsmail.MailNotificationWorkshopInvitation(seqHeaderNow, phaSubSoftware);
                    }
                }
                else if (phaStatus == "12")
                {
                    //13 WF Waiting Follow Up
                    //21 WA Waiting Approve Review

                    phaStatusNew = (requestApprover == "1" || expenseType.ToLower() == "capex") ? "21" : "13";
                    if (phaSubSoftware == "jsea") phaStatusNew = "21";

                    // Submit Revision += 1
                    // Update and keep current revision, then copy to new revision
                    ret = update_status_table_now(user_name, seqHeaderNow, phaNoNow, phaStatusNew);
                    ret = update_revision_table_now(user_name, seqHeaderNow, phaNoNow, versionNow, phaStatusNew, expenseType, phaSubSoftware);

                    keep_version(user_name, ref seqHeaderNow, ref versionNow, phaStatusNew, phaSubSoftware,
                        expenseType.ToLower() == "opex",
                        expenseType.ToLower() == "capex",
                        false, false);

                    if (phaStatusNew == "13")
                    {
                        clsmail = new ClassEmail();
                        clsmail.MailNotificationOutstandingAction(string.Empty, seqHeaderNow, phaSubSoftware);
                    }
                    else if (phaStatusNew == "21")
                    {
                        clsmail = new ClassEmail();
                        if (phaSubSoftware == "hazop" || phaSubSoftware == "whatif")
                        {
                            clsmail.MailNotificationApproverTA2(seqHeaderNow, phaSubSoftware, phaStatus, string.Empty);
                        }
                        else if (phaSubSoftware == "jsea")
                        {
                            clsmail.MailNotificationApproverSafetyReviewer(seqHeaderNow, phaSubSoftware, phaStatus);
                        }
                        else if (phaSubSoftware == "hra")
                        {
                            clsmail.MailNotificationApproverQMTSReviewer(seqHeaderNow, phaSubSoftware, phaStatus);
                        }
                    }
                }
                else if (phaStatus == "22")
                {
                    //21 WA Waiting Approve Review
                    phaStatusNew = "21";

                    // Submit by Originator edit and Submit after TA2 Reject, Revision += 1
                    // Update and keep current revision, then copy to new revision
                    ret = update_status_table_now(user_name, seqHeaderNow, phaNoNow, phaStatusNew);
                    ret = update_revision_table_now(user_name, seqHeaderNow, phaNoNow, versionNow, phaStatusNew, string.Empty, phaSubSoftware);

                    keep_version(user_name, ref seqHeaderNow, ref versionNow, phaStatusNew, phaSubSoftware, false, true, false, false);

                    // Clear data for specific send back
                    update_status_table_approver_sendback(user_name, seqHeaderNow);

                    clsmail = new ClassEmail();
                    if (phaSubSoftware == "hazop" || phaSubSoftware == "whatif")
                    {
                        clsmail.MailNotificationApproverTA2(seqHeaderNow, phaSubSoftware, phaStatus, string.Empty);
                    }
                    else if (phaSubSoftware == "hra")
                    {
                        clsmail.MailNotificationApproverQMTSReviewer(seqHeaderNow, phaSubSoftware, phaStatus);
                    }
                }
            }
            else if (flowAction == "confirm_submit_generate" || flowAction == "confirm_submit_generate_without")
            {
                phaStatusNew = phaStatusNow ?? "";
                if (phaStatusNew != "")
                {
                    // Generate Full Report Revision += 1
                    // Save and keep current revision, then copy to new revision
                    ret = update_status_table_now(user_name, seqHeaderNow, phaNoNow, phaStatusNew);
                    ret = update_revision_table_now(user_name, seqHeaderNow, phaNoNow, versionNow, phaStatusNew, string.Empty, phaSubSoftware);

                    keep_version(user_name, ref seqHeaderNow, ref versionNow, phaStatusNew, phaSubSoftware, false, false, false, false);

                    if (flowAction == "confirm_submit_generate")
                    {
                        clsmail = new ClassEmail();
                        clsmail.MailNotificationMemberReview(seqHeaderNow, phaSubSoftware);
                    }
                }
            }
            else if (flowAction == "submit" && subExpenseType == "Study")
            {
                if (phaStatus == "11")
                {
                    // Only one set of data, revision = 0
                    phaStatusNew = "91";
                    versionNow = "1";
                    ret = update_revision_table_now(user_name, seqHeaderNow, phaNoNow, versionNow, phaStatusNew, string.Empty, phaSubSoftware);

                    keep_version(user_name, ref seqHeaderNow, ref versionNow, phaStatusNew, phaSubSoftware, false, false, false, false);

                    clsmail = new ClassEmail();
                    clsmail.MailToAdminCaseStudy(seqHeaderNow, phaSubSoftware);
                }
            }
            else if (flowAction == "submit_moc")
            {
                if (phaStatusNew != "")
                {
                    // Submit to e-MOC Revision += 1
                    // Update and keep current revision, then copy to new revision
                    ret = update_revision_table_now(user_name, seqHeaderNow, phaNoNow, versionNow, phaStatusNew, expenseType, phaSubSoftware);

                    keep_version(user_name, ref seqHeaderNow, ref versionNow, phaStatusNew, phaSubSoftware, false, false, false, false, true);

                    clsmail = new ClassEmail();
                    clsmail.MailNotificationApproverTA2eMOC(seqHeaderNow, phaSubSoftware);
                }
            }

            #endregion flow action submit

            return ret;
        }
        public string set_workflow(SetDataWorkflowModel param)
        {
            string msg = string.Empty;
            string ret = "true";

            ClassJSON cls_json = new ClassJSON();

            DataSet dsDataOld = new DataSet();
            DataTable dt = new DataTable();
            DataSet dsData = new DataSet();
            string seq_header = param.token_doc ?? string.Empty;
            string pha_status = param.pha_status ?? string.Empty;
            string pha_version = param.pha_version ?? string.Empty;
            string user_name = param.user_name ?? string.Empty;
            string role_type = param.role_type ?? string.Empty;
            string flow_action = param.flow_action ?? string.Empty; // submit_register, submit_without, submit_generate_full_report
            string seq = seq_header;
            string pha_sub_software = param.sub_software ?? string.Empty;

            // action type = insert, update, delete, old_data 
            string year_now = DateTime.Now.ToString("yyyy");
            string seq_header_max = get_max("epha_t_header").ToString();
            string pha_no_max = get_pha_no(pha_sub_software, year_now).ToString();
            string seq_header_now = seq;
            string pha_no_now = string.Empty;
            string version_now = string.Empty;
            string pha_status_now = string.Empty;

            bool submit_generate = (flow_action == "confirm_submit_generate" || flow_action == "confirm_submit_generate_without" || flow_action == "submit_without");

            string expense_type = string.Empty;
            string sub_expense_type = string.Empty;
            string request_approver = string.Empty;

            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }

            ConvertJSONresultToDataSet(user_name, ref msg, ref ret, ref dsData, param, pha_status, pha_sub_software);
            if (ret.ToLower() == "error") { goto Next_Line_Convert; }


            DataTable dtHeader = dsData?.Tables["header"]?.Copy() ?? new DataTable();
            DataTable dtGeneral = dsData?.Tables["general"]?.Copy() ?? new DataTable();

            flow_action = flow_action == "submit_complete" ? "submit" : flow_action;

            List<SqlParameter> parameters = new List<SqlParameter>();

            #region check pha_new_version
            if (true)
            {
                bool update_seq = false;
                if ((flow_action == "save" || flow_action == "submit" || flow_action == "submit_without") && pha_status == "11")
                {
                    #region check / update seq กรณีที่ Create ใบงานและ Save ครั้งแรก --> update เฉพาะข้อมูลที่มีในหน้าจอ
                    if (dtHeader?.Rows.Count > 0)
                    {
                        pha_no_now = dtHeader.Rows[0]["pha_no"]?.ToString() ?? string.Empty;
                        seq_header_now = dtHeader.Rows[0]["seq"]?.ToString() ?? string.Empty;

                        if (dtHeader.Rows[0]["action_type"]?.ToString() == "insert")
                        {
                            // กรณีที่มี seq ซ้ำต้อง generate ใหม่
                            string sqlstr = "SELECT seq FROM epha_t_header WHERE seq = @seq ";

                            parameters = new List<SqlParameter>();
                            parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });
                            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
                            if (dt?.Rows.Count > 0)
                            {
                                seq_header_now = seq_header_max;
                                pha_no_now = pha_no_max;

                                update_seq = true;
                            }
                        }
                    }
                    #endregion check / update seq กรณีที่ Create ใบงานและ Save ครั้งแรก --> update เฉพาะข้อมูลที่มีในหน้าจอ
                }
                #region update seq
                if (update_seq)
                {
                    if (dsData?.Tables.Count > 0)
                    {
                        foreach (DataTable table in dsData.Tables)
                        {
                            foreach (DataRow row in table.Rows)
                            {
                                try
                                {
                                    if (table.TableName == "header")
                                    {
                                        row["pha_no"] = pha_no_now;
                                        row["seq"] = seq_header_now;
                                        row["id"] = seq_header_now;
                                    }
                                    else
                                    {
                                        row["id_pha"] = seq_header_now;
                                    }
                                }
                                catch { }
                            }
                        }
                        dsData.AcceptChanges();
                    }
                }
                #endregion update seq
            }
            #endregion check pha_new_version

            if (dsData?.Tables["header"]?.Rows.Count > 0)
            {
                dtHeader = dsData?.Tables["header"]?.Copy() ?? new DataTable();
                dtGeneral = dsData?.Tables["general"]?.Copy() ?? new DataTable();

                if (dtHeader?.Rows.Count > 0)
                {
                    pha_no_now = dtHeader.Rows[0]["pha_no"]?.ToString() ?? string.Empty;
                    seq_header_now = dtHeader.Rows[0]["seq"]?.ToString() ?? string.Empty;
                    version_now = Convert.ToInt32(dtHeader.Rows[0]["PHA_VERSION"]?.ToString() ?? "0").ToString();
                    pha_status_now = dtHeader.Rows[0]["PHA_STATUS"]?.ToString() ?? string.Empty;
                    expense_type = dtGeneral.Rows[0]["expense_type"]?.ToString() ?? string.Empty;
                    sub_expense_type = dtGeneral.Rows[0]["sub_expense_type"]?.ToString() ?? string.Empty;
                    if (sub_expense_type == "Internal Study") { sub_expense_type = "Study"; }

                    request_approver = dtHeader.Rows[0]["request_approver"]?.ToString() ?? string.Empty;
                }
            }

            if (flow_action == "cancel")
            {
                if (pha_status == "11")
                {
                    string pha_status_new = "81";

                    dt = dtHeader?.Copy() ?? new DataTable();
                    dt.AcceptChanges();

                    using (var conn = new SqlConnection(ClassConnectionDb.ConnectionString()))
                    {
                        conn.Open();
                        using (var transaction = conn.BeginTransaction())
                        {
                            try
                            {
                                if (dt.Rows.Count > 0)
                                {
                                    #region update
                                    string sqlstr = "UPDATE epha_t_header SET ";
                                    sqlstr += " PHA_STATUS = @PHA_STATUS";
                                    sqlstr += " WHERE SEQ = @SEQ";
                                    sqlstr += " AND ID = @ID";
                                    sqlstr += " AND YEAR = @YEAR";
                                    sqlstr += " AND PHA_NO = @PHA_NO";

                                    parameters = new List<SqlParameter>
                            {
                                new SqlParameter("@PHA_STATUS", SqlDbType.Int) { Value = pha_status_new },
                                new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[0]["SEQ"] },
                                new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[0]["ID"] },
                                new SqlParameter("@YEAR", SqlDbType.Int) { Value = dt.Rows[0]["YEAR"] },
                                new SqlParameter("@PHA_NO", SqlDbType.VarChar, 200) { Value = dt.Rows[0]["PHA_NO"]?.ToString() ?? string.Empty }
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
                                    }
                                    #endregion update
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
                return cls_json.SetJSONresult(ClassFile.refMsg(ret, msg, seq_header_now));
            }


            if (dsData != null)
            {
                if (dsData?.Tables.Count > 0)
                {
                    if (dtHeader?.Rows.Count > 0)
                    {
                        try
                        {
                            using (TransactionScope scope = new TransactionScope())
                            using (var conn = new SqlConnection(ClassConnectionDb.ConnectionString()))
                            {
                                conn.Open();
                                using (var transaction = conn.BeginTransaction())
                                {
                                    try
                                    {
                                        if (pha_status == "11")
                                        {
                                            ret = set_header(user_name, conn, transaction, dsData, ref seq_header_now, ref version_now, submit_generate);
                                            if (ret != "true") goto Next_Line;
                                        }
                                        #region update details sub software 
                                        if (!(pha_status == "81"))
                                        {
                                            #region update case SAFETY_CRITICAL_EQUIPMENT_SHOW
                                            if (pha_sub_software == "hazop" || pha_sub_software == "jsea" || pha_sub_software == "whatif")
                                            {
                                                try
                                                {
                                                    dt = dtHeader.Copy();
                                                    dt.AcceptChanges();
                                                    if (dt.Rows[0]["action_type"]?.ToString() == "update")
                                                    {
                                                        string sqlstr = "UPDATE epha_t_header SET ";
                                                        sqlstr += " SAFETY_CRITICAL_EQUIPMENT_SHOW = @SAFETY_CRITICAL_EQUIPMENT_SHOW";

                                                        if (pha_sub_software == "jsea")
                                                        {
                                                            sqlstr += " ,request_approver = 1 ";
                                                        }
                                                        sqlstr += " WHERE SEQ = @SEQ";
                                                        sqlstr += " AND ID = @ID";
                                                        sqlstr += " AND YEAR = @YEAR";
                                                        sqlstr += " AND PHA_NO = @PHA_NO";

                                                        parameters = new List<SqlParameter>();
                                                        parameters.Add(new SqlParameter("@SAFETY_CRITICAL_EQUIPMENT_SHOW", SqlDbType.Int) { Value = dt.Rows[0]["SAFETY_CRITICAL_EQUIPMENT_SHOW"] });
                                                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = dt.Rows[0]["SEQ"] });
                                                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = dt.Rows[0]["ID"] });
                                                        parameters.Add(new SqlParameter("@YEAR", SqlDbType.Int) { Value = dt.Rows[0]["YEAR"] });
                                                        parameters.Add(new SqlParameter("@PHA_NO", SqlDbType.VarChar, 200) { Value = dt.Rows[0]["PHA_NO"]?.ToString() ?? string.Empty });


                                                        ClassConnectionDb cls_conn = new ClassConnectionDb();
                                                        ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn);
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    ret = "error: " + ex.Message;
                                                }
                                                if (ret != "true") goto Next_Line;
                                            }
                                            #endregion update case SAFETY_CRITICAL_EQUIPMENT_SHOW

                                            if (pha_sub_software != "")
                                            {
                                                ret = set_parti(user_name, conn, transaction, dsData, seq_header_now, dsDataOld);
                                                if (ret != "true") goto Next_Line;
                                            }
                                            if (pha_sub_software == "hazop")
                                            {
                                                if (dsData != null)
                                                {
                                                    ret = set_hazop_partii(user_name, conn, transaction, dsData, seq_header_now);
                                                    if (ret != "true") goto Next_Line;
                                                }
                                                if (dsData != null)
                                                {
                                                    ret = set_hazop_partiii(user_name, conn, transaction, dsData, seq_header_now);
                                                    if (ret != "true") goto Next_Line;
                                                }
                                            }
                                            else if (pha_sub_software == "jsea")
                                            {
                                                if (dsData != null)
                                                {
                                                    ret = set_jsea_partii(user_name, conn, transaction, dsData, seq_header_now);
                                                    if (ret != "true") goto Next_Line;
                                                }
                                                if (dsData != null)
                                                {
                                                    ret = set_jsea_partiii(user_name, conn, transaction, dsData, seq_header_now);
                                                    if (ret != "true") goto Next_Line;
                                                }
                                            }
                                            else if (pha_sub_software == "whatif")
                                            {
                                                if (dsData != null)
                                                {
                                                    ret = set_whatif_partii(user_name, conn, transaction, dsData, seq_header_now);
                                                    if (ret != "true") goto Next_Line;
                                                }
                                                if (dsData != null)
                                                {
                                                    ret = set_whatif_partiii(user_name, conn, transaction, dsData, seq_header_now);
                                                    if (ret != "true") goto Next_Line;
                                                }
                                            }
                                            else if (pha_sub_software == "hra")
                                            {
                                                if (dsData != null)
                                                {
                                                    ret = set_hra_partii(user_name, conn, transaction, dsData, seq_header_now);
                                                    if (ret != "true") goto Next_Line;
                                                }
                                                if (dsData != null)
                                                {
                                                    ret = set_hra_partiii(user_name, conn, transaction, dsData, seq_header_now);
                                                    if (ret != "true") goto Next_Line;
                                                }
                                                if (dsData != null)
                                                {
                                                    ret = set_hra_partiv(user_name, conn, transaction, dsData, seq_header_now);
                                                    if (ret != "true") goto Next_Line;
                                                }
                                            }
                                            if (pha_sub_software == "hazop" || pha_sub_software == "whatif" || pha_sub_software == "hra")
                                            {
                                                if (dsData != null)
                                                {
                                                    ret = set_recommendations_part(user_name, conn, transaction, dsData, seq_header_now);
                                                    if (ret != "true") goto Next_Line;
                                                }
                                            }

                                            if (pha_sub_software == "hazop" || pha_sub_software == "jsea" || pha_sub_software == "whatif")
                                            {
                                                if (dsData != null)
                                                {
                                                    if (dsData.Tables["ram_level"] != null)
                                                    {
                                                        DataTable dtDef = dsData?.Tables["ram_level"]?.Copy() ?? new DataTable();
                                                        dtDef.AcceptChanges();
                                                        ret = set_ram_level(user_name, conn, transaction, dtDef, seq_header_now);
                                                        if (ret != "true") goto Next_Line;
                                                    }
                                                }
                                                if (dsData != null)
                                                {
                                                    if (dsData.Tables["ram_master"] != null)
                                                    {
                                                        DataTable dtDef = dsData?.Tables["ram_master"]?.Copy() ?? new DataTable();
                                                        dtDef.AcceptChanges();
                                                        ret = set_ram_master(user_name, conn, transaction, dtDef, seq_header_now);
                                                        if (ret != "true") goto Next_Line;
                                                    }
                                                }
                                            }
                                        }
                                        #endregion update details sub software


                                        if (ret == "true")
                                        {
                                            // Commit the SQL transaction if everything is successful
                                            transaction.Commit();

                                            // Mark the transaction scope as complete
                                            scope.Complete();
                                        }
                                        else
                                        {
                                            ret = "error";
                                        }

                                    }
                                    catch (Exception exTransaction)
                                    {
                                        transaction.Rollback();
                                        ret = "error: " + exTransaction.Message;
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            ret = "error: " + ex.Message;
                        }

                    Next_Line:
                        try
                        {
                            if (ret == "true")
                            {
                                FlowActionSubmit(user_name, pha_sub_software, flow_action, expense_type, sub_expense_type, seq_header_now, pha_no_now, version_now, pha_status_now, pha_status, request_approver);
                            }
                        }
                        catch (Exception ex)
                        {
                            ret = ex.Message;
                        }
                    }
                }
            }

        Next_Line_Convert:
            return cls_json.SetJSONresult(ClassFile.refMsgSave(ret, msg, seq_header_now == seq ? string.Empty : seq_header_now, seq_header_now, pha_no_now, string.Empty));
        }
        public string set_workflow_change_employee(SetDataWorkflowModel param)
        {
            string msg = string.Empty;
            string ret = "true";
            ClassJSON cls_json = new ClassJSON();

            DataSet dsData = new DataSet();
            DataTable dt = new DataTable();
            string seq_header = param.token_doc ?? string.Empty;
            string pha_status = param.pha_status ?? string.Empty;
            string pha_version = param.pha_version ?? string.Empty;
            string user_name = param.user_name ?? string.Empty;
            string role_type = param.role_type ?? string.Empty;
            string flow_action = param.flow_action ?? string.Empty;  //change_action_owner, change_approver

            string seq = seq_header;
            string pha_sub_software = param.sub_software ?? string.Empty;

            // action type = insert, update, delete, old_data 
            string year_now = DateTime.Now.ToString("yyyy");
            string seq_header_max = get_max("epha_t_header").ToString();
            string pha_no_max = get_pha_no(pha_sub_software, year_now).ToString();
            string seq_header_now = seq;
            string pha_no_now = string.Empty;
            string version_now = string.Empty;
            string pha_status_now = string.Empty;

            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }

            DataSet dsDataOld = new DataSet();

            ConvertJSONresultToDataSet(user_name, ref msg, ref ret, ref dsData, param, pha_status, pha_sub_software);
            if (ret.ToLower() == "error") { goto Next_Line_Convert; }

            if (dsData.Tables.Count == 0 || dsData.Tables["header"]?.Rows.Count == 0)
            {
                ret = "error: No data found in header table";
                goto Next_Line_Convert;
            }

            DataTable dtHeader = dsData.Tables["header"]?.Copy() ?? new DataTable();
            DataTable dtGeneral = dsData.Tables["general"]?.Copy() ?? new DataTable();

            #region check pha_new_version 
            string expense_type = string.Empty;
            string sub_expense_type = string.Empty;
            string request_approver = string.Empty;

            if (dtHeader?.Rows.Count > 0)
            {
                pha_no_now = dtHeader.Rows[0]["pha_no"]?.ToString() ?? string.Empty;
                seq_header_now = dtHeader.Rows[0]["seq"]?.ToString() ?? string.Empty;

                string sqlstr = "SELECT seq FROM epha_t_header WHERE seq = @seq ";

                List<SqlParameter> parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq ?? "-1" });

                dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
                if (dt?.Rows.Count == 0)
                {
                    seq_header_now = seq_header_max;
                    pha_no_now = pha_no_max;
                }

                if (true)
                {
                    for (int t = 0; t < dsData.Tables.Count; t++)
                    {
                        for (int i = 0; i < dsData.Tables[t]?.Rows.Count; i++)
                        {
                            try
                            {
                                if (dsData.Tables[t].TableName == "header")
                                {
                                    dsData.Tables[t].Rows[i]["pha_no"] = pha_no_now;
                                    dsData.Tables[t].Rows[i]["seq"] = seq_header_now;
                                    dsData.Tables[t].Rows[i]["id"] = seq_header_now;
                                }
                                else
                                {
                                    dsData.Tables[t].Rows[i]["id_pha"] = seq_header_now;
                                }
                            }
                            catch { }
                        }
                    }
                    dsData.AcceptChanges();
                }
            }
            #endregion check pha_new_version

            if (dtHeader?.Rows.Count > 0)
            {
                dtHeader = dsData.Tables["header"].Copy() ?? new DataTable();
                dtGeneral = dsData.Tables["general"]?.Copy() ?? new DataTable();
                pha_no_now = dtHeader.Rows[0]["pha_no"]?.ToString() ?? string.Empty;
                seq_header_now = dtHeader.Rows[0]["seq"]?.ToString() ?? string.Empty;
                version_now = Convert.ToInt32(dtHeader.Rows[0]["PHA_VERSION"]?.ToString() ?? "0").ToString();
                pha_status_now = dtHeader.Rows[0]["PHA_STATUS"]?.ToString() ?? string.Empty;
                expense_type = dtGeneral.Rows[0]["expense_type"]?.ToString() ?? string.Empty;
                sub_expense_type = dtGeneral.Rows[0]["sub_expense_type"]?.ToString() ?? string.Empty;
                if (sub_expense_type == "Internal Study") { sub_expense_type = "Study"; }

                request_approver = dtHeader.Rows[0]["request_approver"]?.ToString() ?? string.Empty;
            }

            if (dtHeader?.Rows.Count > 0)
            {
                try
                {
                    using (TransactionScope scope = new TransactionScope())
                    using (var conn = new SqlConnection(ClassConnectionDb.ConnectionString()))
                    {
                        conn.Open();
                        using (var transaction = conn.BeginTransaction())
                        {
                            try
                            {
                                #region update details sub software
                                ret = set_parti(user_name, conn, transaction, dsData, seq_header_now, dsDataOld, flow_action);
                                if (ret != "true") { goto Next_Line; }

                                if (flow_action == "change_approver")
                                {
                                    ret = set_approver(user_name, conn, transaction, dsData);
                                    if (ret != "true") { goto Next_Line; }
                                }

                                if (flow_action == "change_action_owner")
                                {
                                    if (pha_sub_software == "hazop")
                                    {
                                        ret = set_hazop_partiii(user_name, conn, transaction, dsData, seq_header_now);
                                        if (ret != "true") { goto Next_Line; }
                                    }
                                    else if (pha_sub_software == "whatif")
                                    {
                                        ret = set_whatif_partiii(user_name, conn, transaction, dsData, seq_header_now);
                                        if (ret != "true") { goto Next_Line; }
                                    }
                                }

                            Next_Line:

                                if (ret == "true")
                                {
                                    // Commit the SQL transaction if everything is successful
                                    transaction.Commit();

                                    // Mark the transaction scope as complete
                                    scope.Complete();
                                }
                                else
                                {
                                    ret = "error";
                                }
                                #endregion update details sub software

                            }
                            catch (Exception exTransaction)
                            {
                                transaction.Rollback();
                                ret = "error: " + exTransaction.Message;
                            }

                        }
                    }
                }
                catch (Exception ex)
                {
                    ret = "error: " + ex.Message;
                }


                if (ret == "true")
                {
                    string pha_status_new = pha_status;
                    ret = keep_version(user_name, ref seq_header_now, ref version_now, pha_status_new, pha_sub_software, false, false, false, false);
                }

                if (flow_action == "change_action_owner")
                {
                    string seq_active_list = string.Empty;
                    if (dsData != null)
                    {
                        if (dsData.Tables.Count > 0)
                        {
                            if (pha_sub_software == "whatif" && dsData.Tables.Contains("listworksheet"))
                            {
                                dt = dsData.Tables["listworksheet"]?.Copy() ?? new DataTable();
                            }
                            else if (dsData.Tables.Contains("nodeworksheet"))
                            {
                                dt = dsData.Tables["nodeworksheet"]?.Copy() ?? new DataTable();
                            }
                            else
                            {
                                dt = new DataTable(); // or handle the case when neither table exists
                            }
                            dt.AcceptChanges();
                            if (dt != null)
                            {
                                for (int i = 0; i < dt?.Rows.Count; i++)
                                {
                                    string action_type = dt.Rows[i]["action_type"]?.ToString() ?? string.Empty;
                                    string action_change = dt.Rows[i]["action_change"]?.ToString() ?? string.Empty;
                                    string seq_active = dt.Rows[i]["seq"]?.ToString() ?? string.Empty;
                                    if (action_type == "update" && action_change == "1")
                                    {
                                        seq_active_list += (string.IsNullOrEmpty(seq_active_list) ? string.Empty : ",") + seq_active;
                                    }
                                }
                                if (!string.IsNullOrEmpty(seq_active_list))
                                {
                                    ClassEmail clsmail = new ClassEmail();
                                    clsmail.MailNotificationChangeActionOwner(seq_header_now, pha_sub_software, seq_active_list);
                                }
                            }
                        }
                    }
                    pha_status_now = seq_active_list;
                }
                else if (flow_action == "change_approver")
                {
                    if (dsData != null)
                    {
                        dt = dsData.Tables["approver"]?.Copy() ?? new DataTable();
                        dt.AcceptChanges();
                        if (dt != null)
                        {
                            for (int i = 0; i < dt?.Rows.Count; i++)
                            {
                                string action_type = dt.Rows[i]["action_type"]?.ToString() ?? string.Empty;
                                string action_change = dt.Rows[i]["action_change"]?.ToString() ?? string.Empty;
                                string approver_user_name = dt.Rows[i]["user_name"]?.ToString() ?? string.Empty;
                                if ((action_type == "insert" && action_change == "1") || (action_type == "update" && action_change == "1"))
                                {
                                    ClassEmail clsmail = new ClassEmail();
                                    clsmail.MailNotificationApproverTA2(seq_header_now, pha_sub_software, pha_status, approver_user_name);
                                }
                            }
                        }
                    }
                }
            }

        Next_Line_Convert:
            return cls_json.SetJSONresult(ClassFile.refMsgSave(ret, msg, seq_header_now == seq ? string.Empty : seq_header_now, seq_header_now, pha_no_now, pha_status_now));
        }
        public string edit_worksheet(SetDocWorksheetModel param)
        {
            string msg = string.Empty;
            string ret = "true";
            ClassJSON cls_json = new ClassJSON();

            string role_type = param.role_type ?? string.Empty;
            string user_name = param.user_name ?? string.Empty;
            string pha_seq = param.token_doc ?? string.Empty;
            string pha_status = param.pha_status ?? string.Empty;
            string action = param.action ?? string.Empty;

            string pha_sub_software = (param.sub_software ?? string.Empty).Trim();
            string id_pha = pha_seq;

            string pha_no_now = string.Empty;
            string version_now = string.Empty;
            string seq_header_now = pha_seq;
            string pha_status_new = pha_status;

            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }

            DataSet dsData = new DataSet();
            string jsper = param.json_worksheet ?? string.Empty;

            #region get worksheet   
            try
            {
                if (!string.IsNullOrWhiteSpace(jsper))
                {
                    DataTable dt = cls_json.ConvertJSONresult(jsper);
                    if (dt != null)
                    {
                        dt.TableName = pha_sub_software switch
                        {
                            "hazop" => "nodeworksheet",
                            "jsea" => "tasks_worksheet",
                            "whatif" => "listworksheet",
                            _ => "worksheet"
                        };
                        dsData.Tables.Add(dt.Copy());
                        dsData.AcceptChanges();
                        ret = "true";
                    }
                }
            }
            catch (Exception ex)
            {
                msg = ex.Message;
                ret = "Error";
            }
            #endregion get worksheet 

            if (ret.ToLower() == "error")
            {
                return cls_json.SetJSONresult(ClassFile.refMsgSave(ret, msg, string.Empty, pha_seq, string.Empty, pha_status_new));
            }

            #region update data 
            try
            {
                using (TransactionScope scope = new TransactionScope())
                using (var conn = new SqlConnection(ClassConnectionDb.ConnectionString()))
                {
                    conn.Open();
                    using (var transaction = conn.BeginTransaction())
                    {
                        try
                        {

                            if (pha_sub_software == "hazop")
                            {
                                ret = set_hazop_partiii(user_name, conn, transaction, dsData, seq_header_now);
                            }
                            else if (pha_sub_software == "jsea")
                            {
                                ret = set_jsea_partiii(user_name, conn, transaction, dsData, seq_header_now);
                            }
                            else if (pha_sub_software == "whatif")
                            {
                                ret = set_whatif_partiii(user_name, conn, transaction, dsData, seq_header_now);
                            }
                            else if (pha_sub_software == "hra")
                            {
                                ret = set_hra_partiii(user_name, conn, transaction, dsData, seq_header_now);
                            }

                            if (ret == "true")
                            {
                                // Commit the SQL transaction if everything is successful
                                transaction.Commit();

                                // Mark the transaction scope as complete
                                scope.Complete();
                            }
                            else
                            {
                                ret = "error";
                            }
                        }
                        catch (Exception exTransaction)
                        {
                            transaction.Rollback();
                            ret = "error: " + exTransaction.Message;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }
            #endregion update data 

            return cls_json.SetJSONresult(ClassFile.refMsgSave(ret, msg, string.Empty, pha_seq, string.Empty, pha_status_new));
        }

        public string set_approve(SetDocApproveModel param)
        {
            string msg = "";
            string ret = "true";
            ClassJSON cls_json = new ClassJSON();
            ClassConnectionDb cls_conn = new ClassConnectionDb();

            string role_type = param.role_type ?? "";
            string user_name = param.user_name ?? "";
            string pha_seq = param.token_doc ?? "";
            string pha_status = param.pha_status ?? "";
            string action = param.action ?? "";

            string pha_sub_software = (param.sub_software ?? "").Trim();
            string id_pha = pha_seq;
            string id_session = param.id_session?.ToString() ?? "";
            string seq_approve = param.seq?.ToString() ?? "";
            string action_review = param.action_review ?? "";
            string action_status = param.action_status ?? "";
            string comment = param.comment ?? "";
            string approver_action_type = (action == "save" ? 1 : 2).ToString();
            string user_approver = param.user_approver ?? "";

            string pha_no_now = "";
            string version_now = "";
            string pha_version_text = "";
            string pha_version_desc = "";
            string seq_header_now = pha_seq;

            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }

            #region get drawing  
            DataSet dsData = new DataSet();
            string jsper = param.json_drawing_approver ?? "";
            try
            {
                if (!string.IsNullOrWhiteSpace(jsper))
                {
                    DataTable dt = cls_json.ConvertJSONresult(jsper);
                    if (dt != null)
                    {
                        dt.TableName = "drawing_approver";
                        dsData.Tables.Add(dt.Copy());
                        dsData.AcceptChanges();
                        ret = "";
                    }
                }
            }
            catch (Exception ex) { msg = ex.Message; ret = "Error"; }

            jsper = param.json_approver ?? "";
            try
            {
                if (!string.IsNullOrWhiteSpace(jsper))
                {
                    DataTable dt = cls_json.ConvertJSONresult(jsper);
                    if (dt != null)
                    {
                        dt.TableName = "approver";
                        dsData.Tables.Add(dt.Copy());
                        dsData.AcceptChanges();
                        ret = "";
                    }
                }
            }
            catch (Exception ex) { msg = ex.Message; ret = "Error"; }

            #endregion get drawing 

            string pha_status_new = pha_status;

            #region ตรวจสอบว่า Approver All หรือมีการ Reject หรือไม่
            Boolean bApproveAll = false;
            if (action == "submit")
            {
                if (action_status == "approve")
                {
                    pha_status_new = (pha_sub_software == "jsea" ? "91" : "13");
                }
                else if (action_status == "reject" || action_status == "reject_no_comment")
                {
                    pha_status_new = "22";
                }

                if (pha_status_new == "22")
                {
                    bApproveAll = true;
                }
                else
                {
                    int icount_check = 0;
                    string sqlstr = @"select ta2.* from epha_t_approver ta2 where ta2.id_pha = @id_pha and ta2.id_session = @id_session";

                    List<SqlParameter> parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("@id_pha", SqlDbType.Int) { Value = id_pha });
                    parameters.Add(new SqlParameter("@id_session", SqlDbType.Int) { Value = id_session });
                    DataTable dtcheck = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
                    icount_check = dtcheck?.Rows.Count ?? 0;

                    if (icount_check == 1)
                    {
                        bApproveAll = true;
                    }
                    else
                    {
                        if (pha_sub_software == "hazop" || pha_sub_software == "whatif" || pha_sub_software == "jsea" || pha_sub_software == "hra")
                        {
                            sqlstr = @"select ta2.* from epha_t_approver ta2 where ta2.approver_action_type = 2 and ta2.id_pha = @id_pha and ta2.id_session = @id_session and ta2.seq <> @seq";

                            parameters = new List<SqlParameter>();
                            parameters.Add(new SqlParameter("@id_pha", SqlDbType.Int) { Value = id_pha });
                            parameters.Add(new SqlParameter("@id_session", SqlDbType.Int) { Value = id_session });
                            parameters.Add(new SqlParameter("@seq", SqlDbType.Int) { Value = seq_approve ?? "-1" });

                            dtcheck = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);

                            if (dtcheck?.Rows.Count == (icount_check - 1))
                            {
                                bApproveAll = true;
                            }
                        }
                    }
                }
            }
            #endregion ตรวจสอบว่า Approver All หรือมีการ Reject หรือไม่

            #region ตรวจสอบ version now
            if (true)
            {
                string sqlstr = @"select distinct h.pha_status, h.pha_no, h.pha_version, h.pha_version_text, h.pha_version_desc 
                          from epha_t_header h
                          inner join EPHA_T_GENERAL g on lower(h.id) = lower(g.id_pha) 
                          inner join EPHA_T_SESSION s on lower(h.id) = lower(s.id_pha)  
                          inner join EPHA_T_APPROVER ta2 on lower(h.id) = lower(ta2.id_pha) and s.seq = ta2.id_session  
                          inner join VW_EPHA_PERSON_DETAILS emp on lower(ta2.user_name) = lower(emp.user_name) 
                          inner join VW_EPHA_PERSON_DETAILS empre on lower(h.pha_request_by) = lower(empre.user_name) 
                          inner join (select max(id) as id_session, id_pha from EPHA_T_SESSION group by id_pha ) s1 on h.id = s1.id_pha and s.id = s1.id_session and s.id_pha = s1.id_pha  
                          where h.seq = @seq";

                List<SqlParameter> parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@seq", SqlDbType.Int) { Value = seq_header_now ?? "-1" });
                DataTable dtHeader = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);

                if (dtHeader?.Rows.Count > 0)
                {
                    pha_no_now = dtHeader.Rows[0]["pha_no"]?.ToString() ?? "";
                    version_now = dtHeader.Rows[0]["pha_version"]?.ToString() ?? "";
                    pha_version_text = dtHeader.Rows[0]["pha_version_text"]?.ToString() ?? "";
                    pha_version_desc = dtHeader.Rows[0]["pha_version_desc"]?.ToString() ?? "";
                }
            }
            #endregion ตรวจสอบ version now

            #region update data 
            try
            {
                using (TransactionScope scope = new TransactionScope())
                {
                    string chk_date_review = "getdate()";
                    if (dsData != null)
                    {
                        if (dsData?.Tables["approver"]?.Rows.Count > 0)
                        {
                            DataRow[] dr = dsData.Tables["approver"].Select($"seq={seq_approve}");
                            if (dr?.Length > 0 && !string.IsNullOrEmpty(dr[0]["date_review"].ToString()))
                            {
                                chk_date_review = dr[0]["date_review"]?.ToString() ?? "";
                            }
                        }

                        using (var conn = new SqlConnection(ClassConnectionDb.ConnectionString()))
                        {
                            conn.Open();
                            using (var transaction = conn.BeginTransaction())
                            {
                                try
                                {
                                    ret = StepApprover_UpdateApprover(action_status, comment, chk_date_review, action, user_name, seq_approve, id_pha, id_session, conn, transaction);
                                    if (ret != "true") throw new Exception(ret);

                                    if (dsData?.Tables["drawing_approver"] != null)
                                    {
                                        if (dsData?.Tables["drawing_approver"]?.Rows.Count > 0)
                                        {
                                            ret = StepApprover_UpdateDrawingApprover(user_name, dsData.Tables["drawing_approver"], pha_sub_software, conn, transaction);
                                            if (ret != "true") throw new Exception(ret);
                                        }
                                    }

                                    if (action == "submit" && bApproveAll)
                                    {
                                        ret = StepApprover_UpdateHeader(pha_status_new, pha_version_text, pha_version_desc, user_name, seq_header_now, conn, transaction);
                                        if (ret != "true") throw new Exception(ret);
                                    }

                                    if (ret == "true")
                                    {
                                        transaction.Commit();
                                        scope.Complete();
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
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }
            #endregion update data

            if (action == "submit" && ret == "true")
            {
                try
                {
                    ClassEmail clsmail = new ClassEmail();
                    if (bApproveAll)
                    {
                        if (pha_status_new == "13" || pha_status_new == "91")
                        {
                            ret = update_status_table_now(user_name, seq_header_now, pha_no_now, pha_status_new);
                            ret = update_revision_table_now(user_name, seq_header_now, pha_no_now, version_now, pha_status_new, "", pha_sub_software);
                            keep_version(user_name, ref seq_header_now, ref version_now, pha_status_new, pha_sub_software, false, false, true, false);

                            if (pha_status_new == "13")
                            {
                                clsmail.MailApprovByApprover(pha_seq, pha_sub_software.ToLower(), user_approver);
                                clsmail.MailNotificationOutstandingAction("", pha_seq, pha_sub_software.ToLower());
                            }
                            else if (pha_status_new == "91")
                            {
                                clsmail.MailNotificationReviewerClosedAll(pha_seq, pha_sub_software);
                            }
                        }
                        else if (pha_status_new == "22")
                        {
                            clsmail.MailRejectByApprover(pha_seq, pha_sub_software, user_approver);
                        }
                    }
                    else
                    {
                        clsmail.MailApprovByApprover(pha_seq, pha_sub_software.ToLower(), user_approver);
                    }
                }
                catch (Exception ex_mail) { msg = ex_mail.Message.ToString(); }
            }

            return cls_json.SetJSONresult(ClassFile.refMsgSave(ret, msg, "", pha_seq, "", pha_status_new));
        }

        private string StepApprover_UpdateApprover(string action_status, string comment, string chk_date_review, string action, string user_name, string seq, string id_pha, string id_session, SqlConnection conn, SqlTransaction transaction)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string sqlstr = "update EPHA_T_APPROVER set ";
            sqlstr += " ACTION_STATUS = @action_status";
            sqlstr += " ,COMMENT = @comment";
            sqlstr += " ,DATE_REVIEW = @chk_date_review";
            sqlstr += " ,ACTION_REVIEW = @action_review";
            sqlstr += " ,APPROVER_ACTION_TYPE = @approver_action_type";
            sqlstr += " ,UPDATE_DATE = getdate()";
            sqlstr += " ,UPDATE_BY = @update_by";
            sqlstr += " where SEQ = @seq";
            sqlstr += " and ID = @id";
            sqlstr += " and ID_PHA = @id_pha";
            sqlstr += " and ID_SESSION = @id_session";

            //test function ChkSqlDateYYYYMMDD to string in sql
            List<SqlParameter> parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("@action_status", SqlDbType.NVarChar, 4000) { Value = action_status });
            parameters.Add(new SqlParameter("@comment", SqlDbType.NVarChar, 4000) { Value = comment });
            parameters.Add(new SqlParameter("@chk_date_review", SqlDbType.DateTime) { Value = ConvertToDateTimeOrDBNull(chk_date_review) });
            parameters.Add(new SqlParameter("@action_review", SqlDbType.Int) { Value = (action == "submit") ? 2 : 1 });
            parameters.Add(new SqlParameter("@approver_action_type", SqlDbType.Int) { Value = (action == "submit") ? 2 : 1 });
            parameters.Add(new SqlParameter("@update_by", SqlDbType.NVarChar, 50) { Value = user_name });
            parameters.Add(new SqlParameter("@seq", SqlDbType.Int) { Value = seq });
            parameters.Add(new SqlParameter("@id", SqlDbType.Int) { Value = seq });
            parameters.Add(new SqlParameter("@id_pha", SqlDbType.Int) { Value = id_pha });
            parameters.Add(new SqlParameter("@id_session", SqlDbType.Int) { Value = id_session });

            ClassConnectionDb cls_conn = new ClassConnectionDb();
            return cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
        }

        private string StepApprover_UpdateDrawingApprover(string user_name, DataTable? drawingApproverTable, string pha_sub_software, SqlConnection conn, SqlTransaction transaction)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ 
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";
            try
            {
                string sqlstr = "";
                string _module_name = pha_sub_software;

                DataTable dt = drawingApproverTable?.Copy() ?? new DataTable();
                dt.AcceptChanges();

                foreach (DataRow row in dt.Rows)
                {
                    row["DOCUMENT_MODULE"] = _module_name;
                    string action_type = row["action_type"]?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(action_type))
                    {
                        sqlstr = "";
                        List<SqlParameter> parameters = new List<SqlParameter>();

                        if (action_type == "insert")
                        {
                            sqlstr = "insert into EPHA_T_DRAWING_APPROVER (" +
                                     "SEQ, ID, ID_PHA, ID_SESSION, ID_APPROVER, NO, DOCUMENT_NAME, DOCUMENT_NO, DOCUMENT_FILE_NAME, DOCUMENT_FILE_PATH, DOCUMENT_FILE_SIZE, DESCRIPTIONS, DOCUMENT_MODULE, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY) " +
                                     "values (@SEQ, @ID, @ID_PHA, @ID_SESSION, @ID_APPROVER, @NO, @DOCUMENT_NAME, @DOCUMENT_NO, @DOCUMENT_FILE_NAME, @DOCUMENT_FILE_PATH, @DOCUMENT_FILE_SIZE, @DESCRIPTIONS, @DOCUMENT_MODULE, getdate(), null, @CREATE_BY, @UPDATE_BY)";
                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = row["ID_PHA"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@ID_SESSION", SqlDbType.Int) { Value = row["ID_SESSION"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@ID_APPROVER", SqlDbType.Int) { Value = row["ID_APPROVER"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = row["NO"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@DOCUMENT_NAME", SqlDbType.NVarChar, 4000) { Value = row["DOCUMENT_NAME"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@DOCUMENT_NO", SqlDbType.NVarChar, 4000) { Value = row["DOCUMENT_NO"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@DOCUMENT_FILE_NAME", SqlDbType.NVarChar, 4000) { Value = row["DOCUMENT_FILE_NAME"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@DOCUMENT_FILE_PATH", SqlDbType.NVarChar, 4000) { Value = row["DOCUMENT_FILE_PATH"] ?? DBNull.Value });
                            //parameters.Add(new SqlParameter("@DOCUMENT_FILE_SIZE", SqlDbType.Int) { Value = row["DOCUMENT_FILE_SIZE"] ?? DBNull.Value }); 
                            parameters.Add(new SqlParameter("@DOCUMENT_FILE_SIZE", SqlDbType.Int)
                            {
                                Value = row.Table.Columns.Contains("DOCUMENT_FILE_SIZE") ? ConvertToDBNull(row["DOCUMENT_FILE_SIZE"]) : DBNull.Value
                            });
                            parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = row["DESCRIPTIONS"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@DOCUMENT_MODULE", SqlDbType.NVarChar, 4000) { Value = row["DOCUMENT_MODULE"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = row["CREATE_BY"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] ?? DBNull.Value });
                        }
                        else if (action_type == "update")
                        {
                            sqlstr = "update EPHA_T_DRAWING_APPROVER set ";
                            sqlstr += "NO = @NO, DOCUMENT_NAME = @DOCUMENT_NAME, DOCUMENT_NO = @DOCUMENT_NO, DOCUMENT_FILE_NAME = @DOCUMENT_FILE_NAME, ";
                            sqlstr += "DOCUMENT_FILE_PATH = @DOCUMENT_FILE_PATH, DOCUMENT_FILE_SIZE = @DOCUMENT_FILE_SIZE, DESCRIPTIONS = @DESCRIPTIONS, DOCUMENT_MODULE = @DOCUMENT_MODULE, ";
                            sqlstr += "UPDATE_DATE = getdate(), UPDATE_BY = @UPDATE_BY ";
                            sqlstr += "where SEQ = @SEQ and ID = @ID and ID_PHA = @ID_PHA and ID_SESSION = @ID_SESSION and ID_APPROVER = @ID_APPROVER";

                            parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = row["NO"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@DOCUMENT_NAME", SqlDbType.NVarChar, 4000) { Value = row["DOCUMENT_NAME"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@DOCUMENT_NO", SqlDbType.NVarChar, 4000) { Value = row["DOCUMENT_NO"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@DOCUMENT_FILE_NAME", SqlDbType.NVarChar, 4000) { Value = row["DOCUMENT_FILE_NAME"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@DOCUMENT_FILE_PATH", SqlDbType.NVarChar, 4000) { Value = row["DOCUMENT_FILE_PATH"] ?? DBNull.Value });
                            //parameters.Add(new SqlParameter("@DOCUMENT_FILE_SIZE", SqlDbType.Int) { Value = row["DOCUMENT_FILE_SIZE"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@DOCUMENT_FILE_SIZE", SqlDbType.Int)
                            {
                                Value = row.Table.Columns.Contains("DOCUMENT_FILE_SIZE") ? ConvertToDBNull(row["DOCUMENT_FILE_SIZE"]) : DBNull.Value
                            });
                            parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = row["DESCRIPTIONS"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@DOCUMENT_MODULE", SqlDbType.NVarChar, 4000) { Value = row["DOCUMENT_MODULE"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = row["ID_PHA"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@ID_SESSION", SqlDbType.Int) { Value = row["ID_SESSION"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@ID_APPROVER", SqlDbType.Int) { Value = row["ID_APPROVER"] ?? DBNull.Value });
                        }
                        else if (action_type == "delete")
                        {
                            sqlstr = "delete from EPHA_T_DRAWING_APPROVER where SEQ = @SEQ and ID = @ID and ID_PHA = @ID_PHA and ID_SESSION = @ID_SESSION and ID_APPROVER = @ID_APPROVER";
                            parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = row["ID_PHA"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@ID_SESSION", SqlDbType.Int) { Value = row["ID_SESSION"] ?? DBNull.Value });
                            parameters.Add(new SqlParameter("@ID_APPROVER", SqlDbType.Int) { Value = row["ID_APPROVER"] ?? DBNull.Value });
                        }

                        if (!string.IsNullOrEmpty(sqlstr))
                        {
                            ClassConnectionDb cls_conn = new ClassConnectionDb();
                            ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                            if (ret != "true") throw new Exception(ret);
                        }
                    }
                }
            }
            catch (Exception ex) { ret = ex.Message.ToString(); }

            return ret;
        }

        private string StepApprover_UpdateHeader(string pha_status_new, string version_text, string version_desc, string user_name, string seq_header_now, SqlConnection conn, SqlTransaction transaction)
        {
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ  string user_name ,
            if (!IsAuthorized(user_name)) { return "User is not authorized to perform this action."; }

            string ret = "true";
            try
            {
                string sqlstr = "update epha_t_header set ";
                sqlstr += "PHA_STATUS = @PHA_STATUS, PHA_VERSION_TEXT = @PHA_VERSION_TEXT, PHA_VERSION_DESC = @PHA_VERSION_DESC, ";
                sqlstr += "APPROVE_ACTION_TYPE = 2, APPROVE_STATUS = @APPROVE_STATUS, UPDATE_BY = @UPDATE_BY, UPDATE_DATE = getdate() ";
                sqlstr += "where SEQ = @SEQ and ID = @ID";

                List<SqlParameter> parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@PHA_STATUS", SqlDbType.Int) { Value = pha_status_new });
                parameters.Add(new SqlParameter("@PHA_VERSION_TEXT", SqlDbType.NVarChar, 200) { Value = ConvertToDBNull(version_text) });
                parameters.Add(new SqlParameter("@PHA_VERSION_DESC", SqlDbType.NVarChar, 200) { Value = ConvertToDBNull(version_desc) });
                parameters.Add(new SqlParameter("@APPROVE_STATUS", SqlDbType.NVarChar, 200) { Value = (pha_status_new == "13") ? "approver" : "reject" });
                parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 200) { Value = user_name.ToLower() });
                parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = seq_header_now });
                parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = seq_header_now });

                ClassConnectionDb cls_conn = new ClassConnectionDb();
                ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
            }
            catch (Exception ex) { ret = ex.Message.ToString(); }

            return ret;
        }

        public string set_approve_ta3(SetDocApproveTa3Model param)
        {
            string msg = "";
            string ret = "true";
            ClassJSON cls_json = new ClassJSON();
            ClassConnectionDb cls_conn = new ClassConnectionDb();

            string role_type = param.role_type ?? "";
            string user_name = param.user_name ?? "";
            string pha_seq = param.token_doc ?? "";
            string action = param.action ?? "";
            string pha_sub_software = (param.sub_software ?? "").Trim();
            string pha_status_new = "";

            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }

            DataTable dtDef = new DataTable();
            DataSet dsData = new DataSet();

            #region get json
            try
            {
                if (!string.IsNullOrWhiteSpace(param.json_header))
                {
                    DataTable dt = cls_json.ConvertJSONresult(param.json_header);
                    if (dt != null)
                    {
                        dt.TableName = "header";
                        dsData.Tables.Add(dt.Copy());
                        dsData.AcceptChanges();
                    }
                }

                if (!string.IsNullOrWhiteSpace(param.json_approver))
                {
                    DataTable dt = cls_json.ConvertJSONresult(param.json_approver);
                    if (dt != null)
                    {
                        dt.TableName = "approver";
                        dsData.Tables.Add(dt.Copy());
                        dsData.AcceptChanges();
                    }
                }

                if (!string.IsNullOrWhiteSpace(param.json_approver_ta3))
                {
                    DataTable dt = cls_json.ConvertJSONresult(param.json_approver_ta3);
                    if (dt != null)
                    {
                        dt.TableName = "approver_ta3";
                        dsData.Tables.Add(dt.Copy());
                        dsData.AcceptChanges();
                    }
                }
            }
            catch (Exception ex)
            {
                msg = ex.Message;
                return cls_json.SetJSONresult(ClassFile.refMsgSave("Error", msg));
            }
            #endregion get json

            #region update data
            string connectionString = ClassConnectionDb.ConnectionString();
            using (var conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (var transaction = conn.BeginTransaction())
                {
                    try
                    {
                        string seq_header_now = dsData.Tables["header"].Rows[0]["seq"]?.ToString() ?? "";

                        if (dsData.Tables["approver_ta3"] != null)
                        {
                            DataTable dt = dsData.Tables["approver_ta3"].Copy();
                            dt.AcceptChanges();

                            dtDef = dt.AsEnumerable()
                                .Where(row => row.Field<string>("action_type") == "insert" || row.Field<string>("action_type") == "update")
                                .CopyToDataTable();

                            foreach (DataRow row in dt.Rows)
                            {
                                string action_type = row["action_type"]?.ToString() ?? "";
                                string sqlstr = "";
                                List<SqlParameter> parameters = new List<SqlParameter>();

                                if (action_type == "insert")
                                {
                                    sqlstr = @"
                                INSERT INTO EPHA_T_APPROVER_TA3 (
                                    SEQ, ID, ID_APPROVER, ID_SESSION, ID_PHA, NO, USER_NAME, USER_DISPLAYNAME, APPROVER_ACTION_TYPE, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY
                                ) VALUES (
                                    @SEQ, @ID, @ID_APPROVER, @ID_SESSION, @ID_PHA, @NO, @USER_NAME, @USER_DISPLAYNAME, 0, GETDATE(), NULL, @CREATE_BY, @UPDATE_BY
                                )";

                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ID_APPROVER", SqlDbType.Int) { Value = row["ID_APPROVER"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ID_SESSION", SqlDbType.Int) { Value = row["ID_SESSION"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = int.Parse(seq_header_now) });
                                    parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = row["NO"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@USER_NAME", SqlDbType.NVarChar, 50) { Value = row["USER_NAME"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@USER_DISPLAYNAME", SqlDbType.NVarChar, 4000) { Value = row["USER_DISPLAYNAME"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = row["CREATE_BY"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] ?? DBNull.Value });
                                }
                                else if (action_type == "update")
                                {
                                    sqlstr = @"
                                UPDATE EPHA_T_APPROVER_TA3 SET
                                    NO = @NO, USER_NAME = @USER_NAME, USER_DISPLAYNAME = @USER_DISPLAYNAME, APPROVER_ACTION_TYPE = 0, UPDATE_DATE = GETDATE(), UPDATE_BY = @UPDATE_BY
                                WHERE
                                    SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA AND ID_SESSION = @ID_SESSION AND ID_APPROVER = @ID_APPROVER";

                                    parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = row["NO"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@USER_NAME", SqlDbType.NVarChar, 50) { Value = row["USER_NAME"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@USER_DISPLAYNAME", SqlDbType.NVarChar, 4000) { Value = row["USER_DISPLAYNAME"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = row["ID_PHA"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ID_SESSION", SqlDbType.Int) { Value = row["ID_SESSION"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ID_APPROVER", SqlDbType.Int) { Value = row["ID_APPROVER"] ?? DBNull.Value });
                                }
                                else if (action_type == "delete")
                                {
                                    sqlstr = @"
                                DELETE FROM EPHA_T_APPROVER_TA3
                                WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA AND ID_SESSION = @ID_SESSION AND ID_APPROVER = @ID_APPROVER";

                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = row["ID_PHA"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ID_SESSION", SqlDbType.Int) { Value = row["ID_SESSION"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ID_APPROVER", SqlDbType.Int) { Value = row["ID_APPROVER"] ?? DBNull.Value });
                                }

                                if (!string.IsNullOrEmpty(sqlstr))
                                {
                                    ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                                    if (ret != "true") throw new Exception(ret);
                                }
                            }

                            if (ret == "") ret = "true";
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
                        msg = ex.Message;
                        return cls_json.SetJSONresult(ClassFile.refMsgSave("Error", msg));
                    }
                }
            }
            #endregion update data

            if (action == "submit" && ret == "true")
            {
                if (dtDef != null && dtDef.Rows.Count > 0)
                {
                    DataTable dtApprover = dsData.Tables["approver"].Copy();
                    dtApprover.AcceptChanges();

                    ClassEmail clsmail = new ClassEmail();

                    foreach (DataRow row in dtDef.Rows)
                    {
                        DataRow[] dr = dtApprover.Select("id = " + row["id_approver"]);
                        if (dr.Length == 0) continue;

                        string seq_approver = dr[0]["id"]?.ToString() ?? "";
                        string user_approver = dr[0]["user_name"]?.ToString() ?? "";

                        clsmail.MailNotificationApproverTA3(pha_seq, pha_sub_software.ToLower(), seq_approver, user_approver);
                    }
                }
            }

            return cls_json.SetJSONresult(ClassFile.refMsgSave(ret, msg, "", pha_seq, "", pha_status_new));
        }

        public string set_transfer_monitoring(SetDocTransferMonitoringModel param)
        {
            string msg = "";
            string ret = "true";
            ClassJSON cls_json = new ClassJSON();

            string role_type = param.role_type ?? "";
            string user_name = param.user_name ?? "";
            string pha_seq = param.token_doc ?? "";
            string action = param.action ?? "";
            string pha_sub_software = (param.sub_software ?? "").Trim();
            string pha_status_new = "";
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }
            DataSet dsData = new DataSet();
            DataTable dtDef = new DataTable();

            #region get json  
            string jsper = param.json_header ?? "";
            try
            {
                if (!string.IsNullOrWhiteSpace(jsper))
                {
                    DataTable dt = cls_json.ConvertJSONresult(jsper);
                    if (dt != null)
                    {
                        dt.TableName = "header";
                        dsData.Tables.Add(dt.Copy());
                        dsData.AcceptChanges();
                        ret = "";
                    }
                }
            }
            catch (Exception ex) { msg = ex.Message; ret = "Error"; }

            jsper = param.json_recom_setting ?? "";
            try
            {
                if (!string.IsNullOrWhiteSpace(jsper))
                {
                    DataTable dt = cls_json.ConvertJSONresult(jsper);
                    if (dt != null)
                    {
                        dt.TableName = "recom_setting";
                        dsData.Tables.Add(dt.Copy());
                        dsData.AcceptChanges();
                        ret = "";
                    }
                }
            }
            catch (Exception ex) { msg = ex.Message; ret = "Error"; }

            jsper = param.json_recom_follow ?? "";
            try
            {
                if (!string.IsNullOrWhiteSpace(jsper))
                {
                    DataTable dt = cls_json.ConvertJSONresult(jsper);
                    if (dt != null)
                    {
                        dt.TableName = "recom_follow";
                        dsData.Tables.Add(dt.Copy());
                        dsData.AcceptChanges();
                        ret = "";
                    }
                }
            }
            catch (Exception ex) { msg = ex.Message; ret = "Error"; }
            #endregion get json 

            #region update data
            ClassFunctions cls = new ClassFunctions();

            try
            {
                using (var conn = new SqlConnection(ClassConnectionDb.ConnectionString()))
                {
                    conn.Open();
                    using (var transaction = conn.BeginTransaction())
                    {
                        string seq_header_now = dsData.Tables["header"]?.Rows[0]["seq"]?.ToString() ?? "";

                        if (dsData.Tables["recom_setting"] != null)
                        {
                            DataTable dt = dsData.Tables["recom_setting"].Copy();
                            dt.AcceptChanges();

                            foreach (DataRow row in dt.Rows)
                            {
                                string action_type = row["action_type"]?.ToString() ?? "";
                                string sqlstr = "";
                                List<SqlParameter> parameters = new List<SqlParameter>();

                                if (action_type == "insert")
                                {
                                    sqlstr = "insert into EPHA_T_RECOM_SETTING (" +
                                             "ID_PHA,SEQ,ID,RECOMMENDATIONS,ID_RANGTYPE,RANGTYPE_VALUES,TARGET_START_DATE,TARGET_END_DATE" +
                                             ",CREATE_DATE,UPDATE_DATE,CREATE_BY,UPDATE_BY) " +
                                             "values (@ID_PHA,@SEQ,@ID,@RECOMMENDATIONS,@ID_RANGTYPE,@RANGTYPE_VALUES,@TARGET_START_DATE,@TARGET_END_DATE,getdate(),null,@CREATE_BY,@UPDATE_BY)";

                                    parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = seq_header_now });
                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@RECOMMENDATIONS", SqlDbType.NVarChar, 4000) { Value = row["RECOMMENDATIONS"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ID_RANGTYPE", SqlDbType.Int) { Value = row["ID_RANGTYPE"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@RANGTYPE_VALUES", SqlDbType.Int) { Value = row["RANGTYPE_VALUES"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@TARGET_START_DATE", SqlDbType.Date) { Value = row["TARGET_START_DATE"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@TARGET_END_DATE", SqlDbType.Date) { Value = row["TARGET_END_DATE"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = row["CREATE_BY"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] ?? DBNull.Value });
                                }
                                else if (action_type == "update")
                                {
                                    sqlstr = "update EPHA_T_RECOM_SETTING set " +
                                             "RECOMMENDATIONS = @RECOMMENDATIONS, ID_RANGTYPE = @ID_RANGTYPE, RANGTYPE_VALUES = @RANGTYPE_VALUES, " +
                                             "TARGET_START_DATE = @TARGET_START_DATE, TARGET_END_DATE = @TARGET_END_DATE, " +
                                             "UPDATE_DATE = getdate(), UPDATE_BY = @UPDATE_BY " +
                                             "where SEQ = @SEQ and ID = @ID and ID_PHA = @ID_PHA";

                                    parameters.Add(new SqlParameter("@RECOMMENDATIONS", SqlDbType.NVarChar, 4000) { Value = row["RECOMMENDATIONS"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ID_RANGTYPE", SqlDbType.Int) { Value = row["ID_RANGTYPE"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@RANGTYPE_VALUES", SqlDbType.Int) { Value = row["RANGTYPE_VALUES"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@TARGET_START_DATE", SqlDbType.Date) { Value = row["TARGET_START_DATE"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@TARGET_END_DATE", SqlDbType.Date) { Value = row["TARGET_END_DATE"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = row["ID_PHA"] ?? DBNull.Value });
                                }
                                else if (action_type == "delete")
                                {
                                    sqlstr = "delete from EPHA_T_RECOM_SETTING where SEQ = @SEQ and ID = @ID and ID_PHA = @ID_PHA";

                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = row["ID_PHA"] ?? DBNull.Value });
                                }

                                if (!string.IsNullOrEmpty(sqlstr))
                                {
                                    ClassConnectionDb cls_conn = new ClassConnectionDb();
                                    ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                                    if (ret != "true") throw new Exception(ret);
                                }
                            }

                            if (ret != "true") return ret;
                        }

                        if (dsData.Tables["recom_follow"] != null)
                        {
                            DataTable dt = dsData.Tables["recom_follow"].Copy();
                            dt.AcceptChanges();

                            foreach (DataRow row in dt.Rows)
                            {
                                string action_type = row["action_type"]?.ToString() ?? "";
                                string sqlstr = "";
                                List<SqlParameter> parameters = new List<SqlParameter>();

                                if (action_type == "insert")
                                {
                                    sqlstr = "insert into EPHA_T_RECOM_FOLLOW (" +
                                             "ID_PHA,SEQ,ID,NO,ID_RECOM,CHECK_TYPE,CHECK_DATE,CREATE_DATE,UPDATE_DATE,CREATE_BY,UPDATE_BY) " +
                                             "values (@ID_PHA,@SEQ,@ID,@NO,@ID_RECOM,@CHECK_TYPE,@CHECK_DATE,getdate(),null,@CREATE_BY,@UPDATE_BY)";

                                    parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = seq_header_now });
                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = row["NO"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ID_RECOM", SqlDbType.Int) { Value = row["ID_RECOM"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@CHECK_TYPE", SqlDbType.Int) { Value = row["CHECK_TYPE"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@CHECK_DATE", SqlDbType.Date) { Value = row["CHECK_DATE"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = row["CREATE_BY"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] ?? DBNull.Value });
                                }
                                else if (action_type == "update")
                                {
                                    sqlstr = "update EPHA_T_RECOM_FOLLOW set " +
                                             "NO = @NO, CHECK_TYPE = @CHECK_TYPE, CHECK_DATE = @CHECK_DATE, " +
                                             "UPDATE_DATE = getdate(), UPDATE_BY = @UPDATE_BY " +
                                             "where SEQ = @SEQ and ID = @ID and ID_PHA = @ID_PHA and ID_RECOM = @ID_RECOM";

                                    parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = row["NO"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@CHECK_TYPE", SqlDbType.Int) { Value = row["CHECK_TYPE"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@CHECK_DATE", SqlDbType.Date) { Value = row["CHECK_DATE"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = row["ID_PHA"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ID_RECOM", SqlDbType.Int) { Value = row["ID_RECOM"] ?? DBNull.Value });
                                }
                                else if (action_type == "delete")
                                {
                                    sqlstr = "delete from EPHA_T_RECOM_FOLLOW where SEQ = @SEQ and ID = @ID and ID_PHA = @ID_PHA and ID_RECOM = @ID_RECOM";

                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = row["ID_PHA"] ?? DBNull.Value });
                                    parameters.Add(new SqlParameter("@ID_RECOM", SqlDbType.Int) { Value = row["ID_RECOM"] ?? DBNull.Value });
                                }

                                if (!string.IsNullOrEmpty(sqlstr))
                                {
                                    ClassConnectionDb cls_conn = new ClassConnectionDb();
                                    ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                                    if (ret != "true") throw new Exception(ret);
                                }
                            }

                            if (ret != "true") return ret;
                        }

                        if (ret == "true")
                        {
                            transaction.Commit();
                        }
                        else
                        {
                            transaction.Rollback();
                            ret = "error";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }
            finally
            {
            }
            #endregion update data

            return cls_json.SetJSONresult(ClassFile.refMsgSave(ret, msg, "", pha_seq, "", pha_status_new));
        }

        #endregion set page worksheet


        //******************************************************************

        #region followup & review followup
        public string set_follow_up(SetDataWorkflowModel param)
        {
            string msg = "";
            string ret = "true";
            ClassJSON cls_json = new ClassJSON();
            ClassConnectionDb cls_conn = new ClassConnectionDb();

            string user_name = param.user_name ?? "";
            string flow_action = param.flow_action ?? "";
            string sub_software = param.sub_software ?? "";
            string role_type = param.role_type ?? "";

            string table_name = sub_software switch
            {
                "hazop" => "EPHA_T_NODE_WORKSHEET",
                "jsea" => "EPHA_T_TASKS_WORKSHEET",
                "whatif" => "EPHA_T_LIST_WORKSHEET",
                "hra" => "EPHA_T_TABLE3_WORKSHEET",
                _ => throw new ArgumentException("Invalid sub_software value")
            };
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }

            DataSet dsData = new DataSet();
            try
            {
                getJsontoDataSet(param.json_managerecom ?? "", ref dsData, "managerecom", user_name);
                getJsontoDataSet(param.json_drawingworksheet ?? "", ref dsData, "drawingworksheet", user_name);
            }
            catch (Exception ex)
            {
                msg = ex.Message;
                ret = "Error";
                return cls_json.SetJSONresult(ClassFile.refMsg(ret, msg));
            }

            if (dsData != null)
            {
                if (dsData.Tables["managerecom"] != null)
                {
                    DataTable dt = dsData.Tables["managerecom"]?.Copy() ?? new DataTable();
                    dt.AcceptChanges();

                    try
                    {
                        using (var conn = new SqlConnection(ClassConnectionDb.ConnectionString()))
                        {
                            conn.Open();
                            using (var transaction = conn.BeginTransaction())
                            {
                                foreach (DataRow row in dt.Rows)
                                {
                                    List<SqlParameter> parameters = new List<SqlParameter>();

                                    parameters.Add(new SqlParameter("@DOCUMENT_FILE_NAME", SqlDbType.NVarChar, 4000) { Value = row["DOCUMENT_FILE_NAME"] });
                                    parameters.Add(new SqlParameter("@DOCUMENT_FILE_PATH", SqlDbType.NVarChar, 4000) { Value = row["DOCUMENT_FILE_PATH"] });
                                    //parameters.Add(new SqlParameter("@DOCUMENT_FILE_SIZE", SqlDbType.Int) { Value = row["DOCUMENT_FILE_SIZE"] });
                                    parameters.Add(new SqlParameter("@DOCUMENT_FILE_SIZE", SqlDbType.Int)
                                    {
                                        Value = row.Table.Columns.Contains("DOCUMENT_FILE_SIZE") ? ConvertToDBNull(row["DOCUMENT_FILE_SIZE"]) : DBNull.Value
                                    });
                                    parameters.Add(new SqlParameter("@RESPONDER_COMMENT", SqlDbType.NVarChar, 4000) { Value = row["RESPONDER_COMMENT"] });
                                    parameters.Add(new SqlParameter("@IMPLEMENT", SqlDbType.Int) { Value = row["IMPLEMENT"] });


                                    if (sub_software == "hazop" || sub_software == "whatif")
                                    {
                                        parameters.Add(new SqlParameter("@RAM_ACTION_SECURITY", SqlDbType.NVarChar, 50) { Value = row["RAM_ACTION_SECURITY"] });
                                        parameters.Add(new SqlParameter("@RAM_ACTION_LIKELIHOOD", SqlDbType.NVarChar, 50) { Value = row["RAM_ACTION_LIKELIHOOD"] });
                                        parameters.Add(new SqlParameter("@RAM_ACTION_RISK", SqlDbType.NVarChar, 50) { Value = row["RAM_ACTION_RISK"] });
                                    }
                                    else if (sub_software == "hra")
                                    {
                                        parameters.Add(new SqlParameter("@RESIDUAL_RISK_RATING", SqlDbType.NVarChar, 4000) { Value = row["RESIDUAL_RISK_RATING"] });
                                    }

                                    if (flow_action == "save")
                                    {
                                        parameters.Add(new SqlParameter("@RESPONDER_ACTION_TYPE", SqlDbType.Int) { Value = 1 });
                                        parameters.Add(new SqlParameter("@ACTION_STATUS", SqlDbType.NVarChar, 50) { Value = "Open" });
                                    }
                                    else
                                    {
                                        parameters.Add(new SqlParameter("@RESPONDER_ACTION_TYPE", SqlDbType.Int) { Value = 2 });
                                        parameters.Add(new SqlParameter("@RESPONDER_ACTION_DATE", SqlDbType.DateTime) { Value = DateTime.Now });
                                        parameters.Add(new SqlParameter("@ACTION_STATUS", SqlDbType.NVarChar, 50) { Value = "Responed" });
                                    }

                                    parameters.Add(new SqlParameter("@UPDATE_DATE", SqlDbType.DateTime) { Value = DateTime.Now });
                                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = user_name });
                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] });
                                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] });
                                    parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = row["ID_PHA"] });
                                    parameters.Add(new SqlParameter("@RESPONDER_USER_NAME", SqlDbType.NVarChar, 50) { Value = row["RESPONDER_USER_NAME"] });

                                    string sqlstr = $@"
                            UPDATE {table_name} SET 
                                DOCUMENT_FILE_NAME = @DOCUMENT_FILE_NAME,
                                DOCUMENT_FILE_PATH = @DOCUMENT_FILE_PATH,
                                DOCUMENT_FILE_SIZE = @DOCUMENT_FILE_SIZE,
                                RESPONDER_COMMENT = @RESPONDER_COMMENT,
                                IMPLEMENT = @IMPLEMENT,
                                {(sub_software == "hazop" || sub_software == "whatif" ? "RAM_ACTION_SECURITY = @RAM_ACTION_SECURITY, RAM_ACTION_LIKELIHOOD = @RAM_ACTION_LIKELIHOOD, RAM_ACTION_RISK = @RAM_ACTION_RISK," : "")}
                                {(sub_software == "hra" ? "RESIDUAL_RISK_RATING = @RESIDUAL_RISK_RATING," : "")}
                                RESPONDER_ACTION_TYPE = @RESPONDER_ACTION_TYPE,
                                {(flow_action != "save" ? "RESPONDER_ACTION_DATE = @RESPONDER_ACTION_DATE," : "")}
                                ACTION_STATUS = @ACTION_STATUS,
                                UPDATE_DATE = @UPDATE_DATE,
                                UPDATE_BY = @UPDATE_BY
                            WHERE 
                                SEQ = @SEQ AND 
                                ID = @ID AND 
                                ID_PHA = @ID_PHA AND 
                                RESPONDER_USER_NAME = @RESPONDER_USER_NAME";

                                    ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                                    if (ret != "true") throw new Exception(ret);
                                }

                                if (dsData.Tables["drawingworksheet"] != null)
                                {
                                    DataTable dtDrawing = dsData.Tables["drawingworksheet"].Copy();
                                    dtDrawing.AcceptChanges();

                                    foreach (DataRow row in dtDrawing.Rows)
                                    {
                                        List<SqlParameter> parameters = new List<SqlParameter>();

                                        parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] });
                                        parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] });
                                        parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = row["ID_PHA"] });
                                        parameters.Add(new SqlParameter("@ID_WORKSHEET", SqlDbType.Int) { Value = row["ID_WORKSHEET"] });
                                        parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = row["NO"] });
                                        parameters.Add(new SqlParameter("@DOCUMENT_NAME", SqlDbType.NVarChar, 4000) { Value = row["DOCUMENT_NAME"] });
                                        parameters.Add(new SqlParameter("@DOCUMENT_NO", SqlDbType.NVarChar, 4000) { Value = row["DOCUMENT_NO"] });
                                        parameters.Add(new SqlParameter("@DOCUMENT_FILE_NAME", SqlDbType.NVarChar, 4000) { Value = row["DOCUMENT_FILE_NAME"] });
                                        parameters.Add(new SqlParameter("@DOCUMENT_FILE_PATH", SqlDbType.NVarChar, 4000) { Value = row["DOCUMENT_FILE_PATH"] });
                                        parameters.Add(new SqlParameter("@DOCUMENT_FILE_SIZE", SqlDbType.Int) { Value = row["DOCUMENT_FILE_SIZE"] });
                                        parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = row["DESCRIPTIONS"] });
                                        parameters.Add(new SqlParameter("@DOCUMENT_MODULE", SqlDbType.NVarChar, 4000) { Value = sub_software });
                                        parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = row["CREATE_BY"] });
                                        parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] });

                                        string action_type = row["action_type"]?.ToString() ?? "";
                                        string sqlstr = "";

                                        if (action_type == "insert")
                                        {
                                            sqlstr = @"
                                    INSERT INTO EPHA_T_DRAWING_WORKSHEET (
                                        SEQ, ID, ID_PHA, ID_WORKSHEET, NO, DOCUMENT_NAME, DOCUMENT_NO, DOCUMENT_FILE_NAME, DOCUMENT_FILE_PATH, DOCUMENT_FILE_SIZE, DESCRIPTIONS, DOCUMENT_MODULE, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY
                                    ) VALUES (
                                        @SEQ, @ID, @ID_PHA, @ID_WORKSHEET, @NO, @DOCUMENT_NAME, @DOCUMENT_NO, @DOCUMENT_FILE_NAME, @DOCUMENT_FILE_PATH, @DOCUMENT_FILE_SIZE, @DESCRIPTIONS, @DOCUMENT_MODULE, getdate(), null, @CREATE_BY, @UPDATE_BY
                                    )";
                                        }
                                        else if (action_type == "update")
                                        {
                                            sqlstr = @"
                                    UPDATE EPHA_T_DRAWING_WORKSHEET SET 
                                        NO = @NO, DOCUMENT_NAME = @DOCUMENT_NAME, DOCUMENT_NO = @DOCUMENT_NO, DOCUMENT_FILE_NAME = @DOCUMENT_FILE_NAME, 
                                        DOCUMENT_FILE_PATH = @DOCUMENT_FILE_PATH, DOCUMENT_FILE_SIZE = @DOCUMENT_FILE_SIZE, DESCRIPTIONS = @DESCRIPTIONS, DOCUMENT_MODULE = @DOCUMENT_MODULE, 
                                        UPDATE_DATE = getdate(), UPDATE_BY = @UPDATE_BY 
                                    WHERE 
                                        SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA AND ID_WORKSHEET = @ID_WORKSHEET";
                                        }
                                        else if (action_type == "delete")
                                        {
                                            sqlstr = @"
                                    DELETE FROM EPHA_T_DRAWING_WORKSHEET 
                                    WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA AND ID_WORKSHEET = @ID_WORKSHEET";
                                        }

                                        if (!string.IsNullOrEmpty(sqlstr))
                                        {
                                            ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                                            if (ret != "true") throw new Exception(ret);
                                        }
                                    }
                                }

                                transaction.Commit();
                                ret = "true";
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        ret = "error: " + ex.Message;
                    }
                }

                if (ret == "true" && flow_action == "submit")
                {
                    DataTable dt = dsData.Tables["managerecom"]?.Copy() ?? new DataTable();
                    dt.AcceptChanges();

                    if (dt?.Rows.Count > 0)
                    {
                        bool bResponderCloseAll = false;
                        bool bCloseAll = false;

                        string id_pha = dt.Rows[0]["ID_PHA"].ToString() ?? "";
                        string responder_user_name = dt.Rows[0]["RESPONDER_USER_NAME"].ToString() ?? "";

                        #region Check if all items are updated by the responder
                        string sqlstr = @"
                SELECT SUM(a) - SUM(b) as xcount 
                FROM (
                    SELECT COUNT(1) as a, SUM(CASE WHEN ISNULL(RESPONDER_ACTION_TYPE, 0) = 2 THEN 1 ELSE 0 END) as b 
                    FROM {0} 
                    WHERE ID_PHA = @ID_PHA AND RESPONDER_USER_NAME = @RESPONDER_USER_NAME
                ) z";

                        sqlstr = string.Format(sqlstr, table_name);

                        List<SqlParameter> parameters = new List<SqlParameter>
            {
                new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = id_pha },
                new SqlParameter("@RESPONDER_USER_NAME", SqlDbType.NVarChar, 50) { Value = responder_user_name }
            };

                        DataTable dtcheck = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);

                        if (dtcheck?.Rows.Count > 0 && Convert.ToInt32(dtcheck.Rows[0]["xcount"]) == 0)
                        {
                            bResponderCloseAll = true;
                        }

                        sqlstr = @"
                SELECT SUM(a) - SUM(b) as xcount 
                FROM (
                    SELECT COUNT(1) as a, SUM(CASE WHEN ISNULL(RESPONDER_ACTION_TYPE, 0) = 2 THEN 1 ELSE 0 END) as b 
                    FROM {0} 
                    WHERE RESPONDER_USER_NAME IS NOT NULL AND ID_PHA = @ID_PHA
                ) z";

                        sqlstr = string.Format(sqlstr, table_name);

                        dtcheck = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);

                        if (dtcheck?.Rows.Count > 0 && Convert.ToInt32(dtcheck.Rows[0]["xcount"]) == 0)
                        {
                            bCloseAll = true;
                        }
                        #endregion

                        #region Update PHA status and send notifications
                        if (!string.IsNullOrEmpty(id_pha))
                        {
                            if (bResponderCloseAll)
                            {
                                if (bCloseAll)
                                {
                                    string pha_status_new = "14";
                                    ret = update_status_table_now(user_name, id_pha, "", pha_status_new);

                                    ret = copy_document_file_responder_to_reviewer(user_name, id_pha, sub_software);

                                    ClassEmail clsmail = new ClassEmail();
                                    clsmail.MailNotificationReviewerReviewFollowup(id_pha, "", sub_software, bResponderCloseAll);
                                }
                                else
                                {
                                    ClassEmail clsmail = new ClassEmail();
                                    clsmail.MailNotificationReviewerReviewFollowup(id_pha, responder_user_name, sub_software, bResponderCloseAll);
                                }
                            }
                            else
                            {
                                ClassEmail clsmail = new ClassEmail();
                                clsmail.MailNotificationReviewerReviewFollowup(id_pha, responder_user_name, sub_software, bResponderCloseAll);
                            }
                        }
                        #endregion
                    }
                }
            }
            else { ret = "false"; msg = "No Data."; }

            return cls_json.SetJSONresult(ClassFile.refMsg(ret, msg));
        }

        public string set_follow_up_review(SetDataWorkflowModel param)
        {
            string msg = "";
            string ret = "true";
            ClassJSON cls_json = new ClassJSON();

            DataSet dsData = new DataSet();
            string user_name = param.user_name ?? "";
            string flow_action = param.flow_action ?? "";
            string sub_software = param.sub_software ?? "";
            string pha_seq = param.token_doc ?? "";

            string pha_no_now = "";
            string version_now = "";

            string role_type = param.role_type ?? "";
            string table_name = sub_software switch
            {
                "hazop" => "EPHA_T_NODE_WORKSHEET",
                "jsea" => "EPHA_T_TASKS_WORKSHEET",
                "whatif" => "EPHA_T_LIST_WORKSHEET",
                "hra" => "EPHA_T_TABLE3_WORKSHEET",
                _ => throw new ArgumentException("Invalid sub_software value")
            };
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }

            try
            {
                getJsontoDataSet(param.json_managerecom ?? "", ref dsData, "managerecom", user_name);
                getJsontoDataSet(param.json_general ?? "", ref dsData, "general", user_name);
                getJsontoDataSet(param.json_drawingworksheet ?? "", ref dsData, "drawingworksheet", user_name);
            }
            catch (Exception ex)
            {
                msg = ex.Message;
                ret = "Error";
                return cls_json.SetJSONresult(ClassFile.refMsg(ret, msg));
            }
            if (dsData != null)
            {

                #region ตรวจสอบ version now
                string sqlstr = @"
        SELECT DISTINCT h.pha_status, h.pha_no, h.pha_version, h.pha_version_text, h.pha_version_desc 
        FROM epha_t_header h
        INNER JOIN EPHA_T_GENERAL g ON LOWER(h.id) = LOWER(g.id_pha)
        INNER JOIN VW_EPHA_MAX_SEQ_BY_PHA_NO sm ON LOWER(h.id) = LOWER(sm.id_pha)
        WHERE h.seq = @pha_seq";

                List<SqlParameter> parameters = new List<SqlParameter>
    {
        new SqlParameter("@pha_seq", SqlDbType.Int) { Value = pha_seq }
    };

                DataTable dtHeader = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);
                if (dtHeader?.Rows.Count > 0)
                {
                    pha_no_now = dtHeader.Rows[0]["pha_no"]?.ToString() ?? "";
                    version_now = dtHeader.Rows[0]["pha_version"]?.ToString() ?? "";
                }
                #endregion

                try
                {
                    using (TransactionScope scope = new TransactionScope())
                    using (var conn = new SqlConnection(ClassConnectionDb.ConnectionString()))
                    {
                        conn.Open();
                        using (var transaction = conn.BeginTransaction())
                        {
                            if (dsData.Tables["managerecom"] != null)
                            {
                                DataTable dt = dsData.Tables["managerecom"].Copy();
                                dt.AcceptChanges();

                                foreach (DataRow row in dt.Rows)
                                {
                                    parameters = new List<SqlParameter>();

                                    parameters.Add(new SqlParameter("@DOCUMENT_FILE_ADMIN_NAME", SqlDbType.NVarChar, 4000) { Value = row["DOCUMENT_FILE_NAME"] });
                                    parameters.Add(new SqlParameter("@DOCUMENT_FILE_ADMIN_PATH", SqlDbType.NVarChar, 4000) { Value = row["DOCUMENT_FILE_PATH"] });
                                    //parameters.Add(new SqlParameter("@DOCUMENT_FILE_ADMIN_SIZE", SqlDbType.Int) { Value = row["DOCUMENT_FILE_SIZE"] });
                                    parameters.Add(new SqlParameter("@DOCUMENT_FILE_ADMIN_SIZE", SqlDbType.Int)
                                    {
                                        Value = row.Table.Columns.Contains("DOCUMENT_FILE_ADMIN_SIZE") ? ConvertToDBNull(row["DOCUMENT_FILE_SIZE"]) : DBNull.Value
                                    });
                                    parameters.Add(new SqlParameter("@ACTION_STATUS", SqlDbType.NVarChar, 50) { Value = row["ACTION_STATUS"] });
                                    parameters.Add(new SqlParameter("@REVIEWER_COMMENT", SqlDbType.NVarChar, 4000) { Value = row["REVIEWER_COMMENT"] });
                                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = user_name });
                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] });
                                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] });
                                    parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = row["ID_PHA"] });
                                    parameters.Add(new SqlParameter("@RESPONDER_USER_NAME", SqlDbType.NVarChar, 50) { Value = row["RESPONDER_USER_NAME"] });


                                    if (sub_software == "hazop" || sub_software == "whatif")
                                    {
                                        parameters.Add(new SqlParameter("@RAM_ACTION_SECURITY", SqlDbType.NVarChar, 50) { Value = row["RAM_ACTION_SECURITY"] });
                                        parameters.Add(new SqlParameter("@RAM_ACTION_LIKELIHOOD", SqlDbType.NVarChar, 50) { Value = row["RAM_ACTION_LIKELIHOOD"] });
                                        parameters.Add(new SqlParameter("@RAM_ACTION_RISK", SqlDbType.NVarChar, 50) { Value = row["RAM_ACTION_RISK"] });
                                    }
                                    else if (sub_software == "hra")
                                    {
                                        parameters.Add(new SqlParameter("@RESIDUAL_RISK_RATING", SqlDbType.NVarChar, 4000) { Value = row["RESIDUAL_RISK_RATING"] });
                                    }

                                    if (flow_action == "save")
                                    {
                                        parameters.Add(new SqlParameter("@REVIEWER_ACTION_TYPE", SqlDbType.Int) { Value = 1 });
                                    }
                                    else
                                    {
                                        parameters.Add(new SqlParameter("@REVIEWER_ACTION_TYPE", SqlDbType.Int) { Value = 2 });
                                        parameters.Add(new SqlParameter("@REVIEWER_ACTION_DATE", SqlDbType.DateTime) { Value = DateTime.Now });
                                    }

                                    sqlstr = $@"
                            UPDATE {table_name} SET 
                                DOCUMENT_FILE_ADMIN_NAME = @DOCUMENT_FILE_ADMIN_NAME,
                                DOCUMENT_FILE_ADMIN_PATH = @DOCUMENT_FILE_ADMIN_PATH,
                                DOCUMENT_FILE_ADMIN_SIZE = @DOCUMENT_FILE_ADMIN_SIZE,
                                ACTION_STATUS = @ACTION_STATUS,
                                {(sub_software == "hazop" || sub_software == "whatif" ? "RAM_ACTION_SECURITY = @RAM_ACTION_SECURITY, RAM_ACTION_LIKELIHOOD = @RAM_ACTION_LIKELIHOOD, RAM_ACTION_RISK = @RAM_ACTION_RISK," : "")}
                                {(sub_software == "hra" ? "RESIDUAL_RISK_RATING = @RESIDUAL_RISK_RATING," : "")}
                                REVIEWER_ACTION_TYPE = @REVIEWER_ACTION_TYPE,
                                {(flow_action != "save" ? "REVIEWER_ACTION_DATE = @REVIEWER_ACTION_DATE," : "")}
                                REVIEWER_COMMENT = @REVIEWER_COMMENT,
                                UPDATE_DATE = getdate(),
                                UPDATE_BY = @UPDATE_BY
                            WHERE 
                                SEQ = @SEQ AND 
                                ID = @ID AND 
                                ID_PHA = @ID_PHA AND 
                                RESPONDER_USER_NAME = @RESPONDER_USER_NAME";

                                    ClassConnectionDb cls_conn = new ClassConnectionDb();
                                    ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                                    if (ret != "true") throw new Exception(ret);
                                }
                            }

                            if (dsData.Tables["general"] != null)
                            {
                                DataTable dt = dsData.Tables["general"].Copy();
                                dt.AcceptChanges();

                                foreach (DataRow row in dt.Rows)
                                {
                                    if (row["action_type"]?.ToString() == "insert")
                                    {
                                        parameters = new List<SqlParameter>
                            {
                                new SqlParameter("@REVIEW_FOLOW_COMMENT", SqlDbType.NVarChar, 4000) { Value = row["REVIEW_FOLOW_COMMENT"] },
                                new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = user_name },
                                new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] },
                                new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] },
                                new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = row["ID_PHA"] }
                            };

                                        sqlstr = @"
                                UPDATE EPHA_T_GENERAL SET 
                                    REVIEW_FOLOW_COMMENT = @REVIEW_FOLOW_COMMENT,
                                    UPDATE_DATE = getdate(),
                                    UPDATE_BY = @UPDATE_BY
                                WHERE 
                                    SEQ = @SEQ AND 
                                    ID = @ID AND 
                                    ID_PHA = @ID_PHA";

                                        ClassConnectionDb cls_conn = new ClassConnectionDb();
                                        ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                                        if (ret != "true") throw new Exception(ret);
                                    }
                                }
                            }

                            if (dsData.Tables["drawingworksheet"] != null)
                            {
                                DataTable dt = dsData.Tables["drawingworksheet"].Copy();
                                dt.AcceptChanges();

                                foreach (DataRow row in dt.Rows)
                                {
                                    string action_type = row["action_type"]?.ToString() ?? "";
                                    parameters = new List<SqlParameter>();

                                    parameters.Add(new SqlParameter("@SEQ", SqlDbType.Int) { Value = row["SEQ"] });
                                    parameters.Add(new SqlParameter("@ID", SqlDbType.Int) { Value = row["ID"] });
                                    parameters.Add(new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = row["ID_PHA"] });
                                    parameters.Add(new SqlParameter("@ID_WORKSHEET", SqlDbType.Int) { Value = row["ID_WORKSHEET"] });
                                    parameters.Add(new SqlParameter("@NO", SqlDbType.Int) { Value = row["NO"] });
                                    parameters.Add(new SqlParameter("@DOCUMENT_NAME", SqlDbType.NVarChar, 4000) { Value = row["DOCUMENT_NAME"] });
                                    parameters.Add(new SqlParameter("@DOCUMENT_NO", SqlDbType.NVarChar, 4000) { Value = row["DOCUMENT_NO"] });
                                    parameters.Add(new SqlParameter("@DOCUMENT_FILE_NAME", SqlDbType.NVarChar, 4000) { Value = row["DOCUMENT_FILE_NAME"] });
                                    parameters.Add(new SqlParameter("@DOCUMENT_FILE_PATH", SqlDbType.NVarChar, 4000) { Value = row["DOCUMENT_FILE_PATH"] });
                                    //parameters.Add(new SqlParameter("@DOCUMENT_FILE_SIZE", SqlDbType.Int) { Value = row["DOCUMENT_FILE_SIZE"] });
                                    parameters.Add(new SqlParameter("@DOCUMENT_FILE_SIZE", SqlDbType.Int)
                                    {
                                        Value = row.Table.Columns.Contains("DOCUMENT_FILE_SIZE") ? ConvertToDBNull(row["DOCUMENT_FILE_SIZE"]) : DBNull.Value
                                    });
                                    parameters.Add(new SqlParameter("@DESCRIPTIONS", SqlDbType.NVarChar, 4000) { Value = row["DESCRIPTIONS"] });
                                    parameters.Add(new SqlParameter("@DOCUMENT_MODULE", SqlDbType.NVarChar, 4000) { Value = row["DOCUMENT_MODULE"] ?? "review_followup" });
                                    parameters.Add(new SqlParameter("@CREATE_BY", SqlDbType.NVarChar, 50) { Value = row["CREATE_BY"] });
                                    parameters.Add(new SqlParameter("@UPDATE_BY", SqlDbType.NVarChar, 50) { Value = row["UPDATE_BY"] });


                                    if (action_type == "insert")
                                    {
                                        sqlstr = @"
                                INSERT INTO EPHA_T_DRAWING_WORKSHEET (
                                    SEQ, ID, ID_PHA, ID_WORKSHEET, NO, DOCUMENT_NAME, DOCUMENT_NO, 
                                    DOCUMENT_FILE_NAME, DOCUMENT_FILE_PATH, DOCUMENT_FILE_SIZE, DESCRIPTIONS, 
                                    DOCUMENT_MODULE, CREATE_DATE, UPDATE_DATE, CREATE_BY, UPDATE_BY
                                ) VALUES (
                                    @SEQ, @ID, @ID_PHA, @ID_WORKSHEET, @NO, @DOCUMENT_NAME, @DOCUMENT_NO, 
                                    @DOCUMENT_FILE_NAME, @DOCUMENT_FILE_PATH, @DOCUMENT_FILE_SIZE, @DESCRIPTIONS, 
                                    @DOCUMENT_MODULE, getdate(), null, @CREATE_BY, @UPDATE_BY
                                )";
                                    }
                                    else if (action_type == "update")
                                    {
                                        sqlstr = @"
                                UPDATE EPHA_T_DRAWING_WORKSHEET SET 
                                    NO = @NO, DOCUMENT_NAME = @DOCUMENT_NAME, DOCUMENT_NO = @DOCUMENT_NO, 
                                    DOCUMENT_FILE_NAME = @DOCUMENT_FILE_NAME, DOCUMENT_FILE_PATH = @DOCUMENT_FILE_PATH, 
                                    DOCUMENT_FILE_SIZE = @DOCUMENT_FILE_SIZE, DESCRIPTIONS = @DESCRIPTIONS, 
                                    DOCUMENT_MODULE = @DOCUMENT_MODULE, UPDATE_DATE = getdate(), UPDATE_BY = @UPDATE_BY 
                                WHERE 
                                    SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA AND ID_WORKSHEET = @ID_WORKSHEET";
                                    }
                                    else if (action_type == "delete")
                                    {
                                        sqlstr = @"
                                DELETE FROM EPHA_T_DRAWING_WORKSHEET 
                                WHERE SEQ = @SEQ AND ID = @ID AND ID_PHA = @ID_PHA AND ID_WORKSHEET = @ID_WORKSHEET";
                                    }

                                    if (!string.IsNullOrEmpty(sqlstr))
                                    {
                                        ClassConnectionDb cls_conn = new ClassConnectionDb();
                                        ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                                        if (ret != "true") throw new Exception(ret);
                                    }
                                }
                            }

                            transaction.Commit();
                            ret = "true";
                        }
                    }
                }
                catch (Exception ex)
                {
                    ret = "error: " + ex.Message;
                }

                if (ret == "true" && flow_action == "submit")
                {
                    DataTable dt = dsData.Tables["managerecom"].Copy();
                    dt.AcceptChanges();

                    if (dt?.Rows.Count > 0)
                    {
                        bool bCloseAll = false;

                        string id_pha = dt.Rows[0]["ID_PHA"].ToString();

                        #region Check if all items are updated by the reviewer
                        sqlstr = @"
                SELECT SUM(a) - SUM(b) as xcount 
                FROM (
                    SELECT COUNT(1) as a, SUM(CASE WHEN ISNULL(REVIEWER_ACTION_TYPE, 0) = 2 THEN 1 ELSE 0 END) as b 
                    FROM {0} 
                    WHERE RESPONDER_USER_NAME IS NOT NULL AND ID_PHA = @ID_PHA
                ) z";

                        sqlstr = string.Format(sqlstr, table_name);

                        parameters = new List<SqlParameter>
            {
                new SqlParameter("@ID_PHA", SqlDbType.Int) { Value = id_pha }
            };

                        DataTable dtcheck = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);

                        if (dtcheck?.Rows.Count > 0 && Convert.ToInt32(dtcheck.Rows[0]["xcount"]) == 0)
                        {
                            bCloseAll = true;
                        }
                        #endregion

                        if (bCloseAll)
                        {
                            string pha_status_new = "91";
                            ret = update_status_table_now(user_name, pha_seq, pha_no_now, pha_status_new);
                            ret = update_revision_table_now(user_name, pha_seq, pha_no_now, version_now, pha_status_new, "", sub_software);
                            keep_version(user_name, ref pha_seq, ref version_now, pha_status_new, sub_software, false, false, false, true);

                            ClassEmail clsmail = new ClassEmail();
                            clsmail.MailNotificationReviewerClosedAll(pha_seq, sub_software);
                        }
                    }
                }

            }
            else { ret = "false"; msg = "No Data."; }
            return cls_json.SetJSONresult(ClassFile.refMsg(ret, msg));
        }

        #endregion followup & review followup


        //******************************************************
        #region set send email to member review

        public string set_member_review(string user_name, string id_pha, string sub_software)
        {
            string msg = "";
            string ret = "true";
            ClassJSON cls_json = new ClassJSON();
            // ตรวจสอบสิทธิ์ก่อนดำเนินการ
            if (!IsAuthorized(user_name))
            {
                return cls_json.SetJSONresult(ClassFile.refMsg("Error", "User is not authorized to perform this action."));
            }
            try
            {
                // SQL query to get the session ID
                string sqlstr = @"
            SELECT DISTINCT c.id_session 
            FROM epha_t_header a 
            INNER JOIN EPHA_T_SESSION b ON a.id = b.id_pha 
            INNER JOIN (SELECT MAX(id) AS id, id_pha FROM EPHA_T_SESSION GROUP BY id_pha) b2 ON b.id = b2.id AND b.id_pha = b2.id_pha
            INNER JOIN EPHA_T_MEMBER_TEAM c ON a.id = c.id_pha AND b.id = c.id_session
            INNER JOIN (SELECT MAX(id_session) AS id_session, id_pha FROM EPHA_T_MEMBER_TEAM GROUP BY id_pha) c2 ON c.id_session = c2.id_session AND c.id_pha = c2.id_pha
            WHERE LOWER(a.seq) = LOWER(@id_pha) AND ISNULL(b.action_to_review, 0) = 0 AND ISNULL(c.action_review, 0) = 0";

                List<SqlParameter> parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@id_pha", SqlDbType.NVarChar, 50) { Value = id_pha });


                DataTable dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);

                if (dt?.Rows.Count > 0)
                {
                    string id_session = dt.Rows[0]["id_session"]?.ToString() ?? "";

                    using (var conn = new SqlConnection(ClassConnectionDb.ConnectionString()))
                    {
                        conn.Open();
                        using (var transaction = conn.BeginTransaction())
                        {
                            try
                            {
                                // Update ACTION_REVIEW
                                sqlstr = @"
                            UPDATE EPHA_T_MEMBER_TEAM 
                            SET ACTION_REVIEW = 1, DATE_REVIEW = GETDATE() 
                            WHERE ID_PHA = @id_pha AND ID_SESSION = @id_session";

                                parameters = new List<SqlParameter>
                        {
                            new SqlParameter("@id_pha", SqlDbType.Int) { Value = int.Parse(id_pha) },
                            new SqlParameter("@id_session", SqlDbType.Int) { Value = int.Parse(id_session) }
                        };
                                ClassConnectionDb cls_conn = new ClassConnectionDb();
                                ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                                if (ret != "true") throw new Exception(ret);

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
                else
                {
                    ret = "No data found to update.";
                }
            }
            catch (Exception ex)
            {
                ret = "error: " + ex.Message;
            }

            return cls_json.SetJSONresult(ClassFile.refMsg(ret, msg, ""));
        }

        #endregion set send email to member review
    }
}
