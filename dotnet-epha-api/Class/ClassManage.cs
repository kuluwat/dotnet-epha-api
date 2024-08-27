using dotnet_epha_api.Class;
using dotnet6_epha_api.Class;
using Model;
using Newtonsoft.Json;
using System.Data;
using System.Data.SqlClient;

namespace Class
{
    public class ClassManage
    {
        string sqlstr = "";
        string jsper = "";
        ClassFunctions cls = new ClassFunctions();
        ClassJSON cls_json = new ClassJSON();
        ClassConnectionDb cls_conn = new ClassConnectionDb();
         
        #region home task 
        public string DocumentCopy(ManageDocModel param)
        {
            DataSet dsData = new DataSet();
            string userName = param.user_name?.ToString() ?? "";
            string subSoftware = param.sub_software?.ToString() ?? "";
            string phaNo = param.pha_no?.ToString() ?? "";
            string phaSeq = param.pha_seq?.ToString() ?? "";

            string phaNoNow = "";
            string versionNow = "";
            string seqHeaderNow = phaSeq;
            string phaStatusNow = "11";
            string phaSubSoftware = subSoftware;
            string ret = "";
            string msg = "";

            // pha_sub_software  from epha_t_header  
            cls = new ClassFunctions();
            using (var conn = new SqlConnection(ClassConnectionDb.ConnectionString()))
            {
                conn.Open();
                sqlstr = "select distinct pha_sub_software from epha_t_header where seq = @seq";
                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = phaSeq }
                };

                DataTable dt = new DataTable();
                using (var cmd = new SqlCommand(sqlstr, conn))
                {
                    cmd.Parameters.AddRange(parameters.ToArray());
                    using (var da = new SqlDataAdapter(cmd))
                    {
                        da.Fill(dt);
                    }
                }

                if (dt?.Rows.Count > 0)
                {
                    phaSubSoftware = dt.Rows[0]["pha_sub_software"]?.ToString() ?? "";
                }

                //Copy seqHeaderNow => New Seq
                ClassHazopSet clsHazopSet = new ClassHazopSet();
                clsHazopSet.keep_version(userName, ref seqHeaderNow, ref versionNow, phaStatusNow, phaSubSoftware, false, false, false, false);

                versionNow = "1";
                if (!string.IsNullOrEmpty(seqHeaderNow))
                {
                    string yearNow = DateTime.Now.Year.ToString();
                    if (Convert.ToInt64(yearNow) > 2500)
                    {
                        yearNow = (Convert.ToInt64(yearNow) - 543).ToString();
                    }
                    ClassHazop getPhaNo = new ClassHazop();
                    phaNoNow = getPhaNo.get_pha_no(phaSubSoftware, yearNow);

                    string userNameChkSqlStr = cls.ChkSqlStr(userName, 100);
                    string phaNoChkSqlStr = cls.ChkSqlStr(phaNoNow, 100);
                    string requestUserDisplayName = "";

                    ClassLogin clsLogin = new ClassLogin();
                    DataTable dtUser = clsLogin.dataUserRole(userName);
                    if (dtUser?.Rows.Count > 0)
                    {
                        requestUserDisplayName = cls.ChkSqlStr(dtUser.Rows[0]["user_displayname"]?.ToString() ?? "", 4000);
                    }

                    using (var transaction = conn.BeginTransaction())
                    {
                        try
                        {
                            // pha_request_by,request_user_name,request_user_displayname
                            sqlstr = @"update epha_t_header set flow_mail_to_member = null, pha_version = 1 , pha_version_text = 'A', pha_version_desc = 'Issued for Review'
                                       , update_date = null,  update_by = null, create_date = getdate()
                                       , create_by = @userName, pha_request_by = @userName, request_user_name = @userName, request_user_displayname = @requestUserDisplayName
                                       , pha_status = @phaStatusNow, pha_no  = @phaNo where id = @seqHeaderNow";

                            parameters = new List<SqlParameter>();
                            parameters.Add(new SqlParameter("@requestUserDisplayName", SqlDbType.VarChar, 4000) { Value = requestUserDisplayName });
                            parameters.Add(new SqlParameter("@phaStatusNow", SqlDbType.VarChar, 50) { Value = phaStatusNow });
                            parameters.Add(new SqlParameter("@phaNo", SqlDbType.VarChar, 100) { Value = phaNoNow });
                            parameters.Add(new SqlParameter("@userName", SqlDbType.VarChar, 100) { Value = userName });
                            parameters.Add(new SqlParameter("@seqHeaderNow", SqlDbType.VarChar, 50) { Value = seqHeaderNow });

                            ClassConnectionDb cls_conn = new ClassConnectionDb();
                            ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                            if (ret != "true") { goto NextLine; }

                            sqlstr = @"update epha_t_member_team set action_review = null, date_review = null, comment = null
                                       , update_date = null, update_by = null, create_date = getdate()
                                       , create_by = @userName where id_pha = @seqHeaderNow";

                            parameters = new List<SqlParameter>();
                            parameters.Add(new SqlParameter("@userName", SqlDbType.VarChar, 100) { Value = userName });
                            parameters.Add(new SqlParameter("@seqHeaderNow", SqlDbType.VarChar, 50) { Value = seqHeaderNow });

                              cls_conn = new ClassConnectionDb();
                            ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                            if (ret != "true") { goto NextLine; }

                            sqlstr = @"update epha_t_approver set action_review = null, date_review = null, comment = null, approver_action_type = null, action_status = null
                                       , update_date = null, update_by = null, create_date = getdate()
                                       , create_by = @userName where id_pha = @seqHeaderNow";

                            parameters = new List<SqlParameter>();
                            parameters.Add(new SqlParameter("@userName", SqlDbType.VarChar, 100) { Value = userName });
                            parameters.Add(new SqlParameter("@seqHeaderNow", SqlDbType.VarChar, 50) { Value = seqHeaderNow });
                              cls_conn = new ClassConnectionDb();
                            ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                            if (ret != "true") { goto NextLine; }

                            string[] tables = { "epha_t_general", "epha_t_functional_audition", "epha_t_node", "epha_t_node_drawing", "epha_t_node_guide_words",
                                                "epha_t_list", "epha_t_list_drawing", "epha_t_table1_hazard", "epha_t_table1_subareas",
                                                "epha_t_table2_tasks", "epha_t_table2_workers", "epha_t_table2_descriptions"};

                            foreach (string table in tables)
                            {
                                sqlstr = $"update {table} set update_date = null, update_by = null, create_date = getdate(), create_by = @userName where id_pha = @seqHeaderNow";
                                parameters = new List<SqlParameter>();
                                parameters.Add(new SqlParameter("@userName", SqlDbType.VarChar, 100) { Value = userName });
                                parameters.Add(new SqlParameter("@seqHeaderNow", SqlDbType.VarChar, 50) { Value = seqHeaderNow });
                                cls_conn = new ClassConnectionDb();
                                ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                                if (ret != "true") { goto NextLine; }
                            }

                            if (phaSubSoftware == "hazop")
                            {
                                sqlstr = @"update epha_t_node_worksheet set responder_action_type = null, responder_action_date = null, responder_receivesd_date = null
                                           , responder_comment = null, reviewer_action_type = null, reviewer_action_date = null, reviewer_comment = null
                                           , update_date = null, update_by = null, create_date = getdate()
                                           , create_by = @userName where id_pha = @seqHeaderNow";
                            }
                            else if (phaSubSoftware == "whatif")
                            {
                                sqlstr = @"update epha_t_list_worksheet set responder_action_type = null, responder_action_date = null, responder_receivesd_date = null
                                           , responder_comment = null, reviewer_action_type = null, reviewer_action_date = null, reviewer_comment = null
                                           , update_date = null, update_by = null, create_date = getdate()
                                           , create_by = @userName where id_pha = @seqHeaderNow";
                            }
                            else if (phaSubSoftware == "jsea")
                            {
                                sqlstr = @"update epha_t_tasks_worksheet set responder_action_type = null, responder_action_date = null, responder_receivesd_date = null
                                           , responder_comment = null, reviewer_action_type = null, reviewer_action_date = null, reviewer_comment = null
                                           , update_date = null, update_by = null, create_date = getdate()
                                           , create_by = @userName where id_pha = @seqHeaderNow";
                            }
                            else if (phaSubSoftware == "hra")
                            {
                                sqlstr = @"update epha_t_table3_worksheet set responder_action_type = null, responder_action_date = null, responder_receivesd_date = null
                                           , responder_comment = null, reviewer_action_type = null, reviewer_action_date = null, reviewer_comment = null
                                           , update_date = null, update_by = null, create_date = getdate()
                                           , create_by = @userName where id_pha = @seqHeaderNow";
                            }

                            parameters = new List<SqlParameter>();
                            parameters.Add(new SqlParameter("@userName", SqlDbType.VarChar, 100) { Value = userName });
                            parameters.Add(new SqlParameter("@seqHeaderNow", SqlDbType.VarChar, 50) { Value = seqHeaderNow });
                            cls_conn = new ClassConnectionDb();
                            ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
                            if (ret != "true") { goto NextLine; }

                        NextLine:;
                            if (ret == "true")
                            {
                                msg = "";
                                transaction.Commit();
                            }
                            else
                            {
                                msg = ret;
                                ret = "false";
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

            return cls_json.SetJSONresult(ClassFile.refMsgSave(ret, msg, (seqHeaderNow == phaSeq ? "" : seqHeaderNow), seqHeaderNow, phaNoNow, phaStatusNow));
        }

        public string DocumentCancel(ManageDocModel param)
        {
            DataSet dsData = new DataSet();
            string userName = param.user_name?.ToString() ?? "";
            string subSoftware = param.sub_software?.ToString() ?? "";
            string phaNo = param.pha_no?.ToString() ?? "";
            string phaSeq = param.pha_seq?.ToString() ?? "";
            string phaStatusComment = param.pha_status_comment?.ToString() ?? "";

            string phaSubSoftware = subSoftware;
            string ret = "";
            string msg = "";

            cls = new ClassFunctions();
            using (var conn = new SqlConnection(ClassConnectionDb.ConnectionString()))
            {
                conn.Open();
                using (var transaction = conn.BeginTransaction())
                {
                    try
                    {
                        sqlstr = @"update epha_t_header set pha_status = 81, pha_status_comment = @phaStatusComment, update_date = getdate(), update_by = @userName where id = @phaSeq";
                        var parameters = new List<SqlParameter>
                        {
                            new SqlParameter("@phaStatusComment", SqlDbType.VarChar, 4000) { Value = phaStatusComment },
                            new SqlParameter("@userName", SqlDbType.VarChar, 100) { Value = userName },
                            new SqlParameter("@phaSeq", SqlDbType.VarChar, 50) { Value = phaSeq }
                        };

                        cls_conn = new ClassConnectionDb();
                        ret = cls_conn.ExecuteNonQuerySQLTrans(sqlstr, parameters, conn, transaction);
               
                        if (ret == "true")
                        {
                            msg = "";
                            transaction.Commit();
                        }
                        else
                        {
                            msg = ret;
                            ret = "false";
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

            return cls_json.SetJSONresult(ClassFile.refMsgSave(ret, msg, "", "", "", ""));
        }
        #endregion home task 
    }
}
