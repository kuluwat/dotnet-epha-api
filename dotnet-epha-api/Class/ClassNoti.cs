using dotnet6_epha_api.Class;
using System.Data;
using System.Data.SqlClient;
namespace Class
{ 
    public class ClassNoti
    {
        string sqlstr = "";  

        public DataTable DataDailyByActionRequired(string user_name, string seq, string sub_software, Boolean group_by_user, Boolean home_task)
        {
            DataTable dt = new DataTable(); 
            var parameters = new List<SqlParameter>();

            #region get data 
            sqlstr = @"
        select a.id as id_pha, a.pha_status,
               isnull(nw.responder_user_name,'') as user_name, emp.user_displayname, emp.user_email,
               isnull(a.pha_request_by,'') as user_name_ori,
               0 as id_action, nw.responder_action_date as user_action_date,
               1 as action_sort,
               0 as task, upper(a.pha_sub_software) as pha_type,
               'Recommendation Closing' as action_required, a.pha_no as document_number, g.pha_request_name as document_title,
               nw.recommendations_no as rev_def, a.pha_version_text + '('+  a.pha_version_desc +')' as rev, emp_ori.user_displayname as originator,
               format(nw.responder_receivesd_date,'dd MMM yyyy') as receivesd, format(nw.estimated_start_date,'dd MMM yyyy') as due_date,
               nw.responder_action_date as action_date,
               isnull(datediff(day, getdate(),case when nw.estimated_start_date >= getdate() then nw.estimated_start_date else getdate() end),0) as remaining,
               emp_conso.user_displayname as consolidator
        from epha_t_header a
        inner join EPHA_T_GENERAL g on a.id = g.id_pha
        inner join (select max(id) as id_session, id_pha from EPHA_T_SESSION group by id_pha) s on a.id = s.id_pha
        inner join (select max(id_session) as id_session, id_pha from EPHA_T_MEMBER_TEAM group by id_pha) t2 on a.id = t2.id_pha and s.id_session = t2.id_session
        inner join EPHA_T_NODE_WORKSHEET nw on a.id = nw.id_pha
        left join VW_EPHA_PERSON_DETAILS emp on lower(nw.responder_user_name) = lower(emp.user_name)
        left join VW_EPHA_PERSON_DETAILS emp_ori on lower(a.pha_request_by) = lower(emp_ori.user_name)
        left join VW_EPHA_PERSON_DETAILS emp_conso on lower(nw.responder_user_name) = lower(emp_conso.user_name)
        where nw.responder_user_name is not null and nw.estimated_start_date is not null and nw.responder_action_date is null
          and a.seq in (select max(seq) from epha_t_header group by pha_no) and a.pha_status = 13

        union

        select a.id as id_pha, a.pha_status,
               isnull(a.pha_request_by,'') as user_name, emp.user_displayname, emp.user_email,
               isnull(a.pha_request_by,'') as user_name_ori,
               0 as id_action, null as user_action_date,
               3 as action_sort,
               0 as task, upper(a.pha_sub_software) as pha_type,
               'Approve' as action_required, a.pha_no as document_number, g.pha_request_name as document_title,
               nw.recommendations_no as rev_def, a.pha_version_text + '('+  a.pha_version_desc +')' as rev, emp_ori.user_displayname as originator,
               format(nw.responder_action_date,'dd MMM yyyy') as receivesd, format(g.target_end_date,'dd MMM yyyy') as due_date,
               null as action_date,
               isnull(datediff(day, getdate(),case when g.target_end_date >= getdate() then g.target_end_date else getdate() end),0) as remaining,
               emp_conso.user_displayname as consolidator
        from epha_t_header a
        inner join EPHA_T_GENERAL g on a.id = g.id_pha
        inner join EPHA_T_SESSION s on a.id = s.id_pha
        inner join (select max(id) as id_session, id_pha from EPHA_T_SESSION group by id_pha) s2 on a.id = s2.id_pha and s.id = s2.id_session
        inner join (select max(id_session) as id_session, id_pha from EPHA_T_MEMBER_TEAM group by id_pha) t2 on a.id = t2.id_pha and s2.id_session = t2.id_session
        inner join EPHA_T_NODE_WORKSHEET nw on a.id = nw.id_pha
        left join VW_EPHA_PERSON_DETAILS emp on lower(a.pha_request_by) = lower(emp.user_name)
        left join VW_EPHA_PERSON_DETAILS emp_ori on lower(a.pha_request_by) = lower(emp_ori.user_name)
        left join VW_EPHA_PERSON_DETAILS emp_conso on lower(nw.responder_user_name) = lower(emp_conso.user_name)
        where a.pha_status in (14) and nw.responder_user_name is not null and nw.responder_action_date is not null and lower(nw.action_status) not in ('closed')
          and a.seq in (select max(seq) from epha_t_header group by pha_no) and a.pha_status = 14

        union

        select distinct a.id as id_pha, a.pha_status,
                        isnull(emp.user_name,'') as user_name, emp.user_displayname, emp.user_email,
                        isnull(a.pha_request_by,'') as user_name_ori,
                        0 as id_action, null as user_action_date,
                        5 as action_sort,
                        0 as task, upper(a.pha_sub_software) as pha_type,
                        'Approver' as action_required, a.pha_no as document_number, g.pha_request_name as document_title,
                        nw.recommendations_no as rev_def, a.pha_version_text + '('+  a.pha_version_desc +')' as rev, emp_ori.user_displayname as originator,
                        format(nw.responder_action_date,'dd MMM yyyy') as receivesd, format(g.target_end_date,'dd MMM yyyy') as due_date,
                        null as action_date,
                        isnull(datediff(day, getdate(),case when g.target_end_date >= getdate() then g.target_end_date else getdate() end),0) as remaining,
                        '' as consolidator
        from epha_t_header a
        inner join EPHA_T_GENERAL g on a.id = g.id_pha
        inner join EPHA_T_SESSION s on a.id = s.id_pha
        inner join (select max(id) as id_session, id_pha from EPHA_T_SESSION group by id_pha) s2 on a.id = s2.id_pha and s.id = s2.id_session
        inner join (select max(id_session) as id_session, id_pha from EPHA_T_MEMBER_TEAM group by id_pha) t2 on a.id = t2.id_pha and s2.id_session = t2.id_session
        inner join EPHA_T_NODE_WORKSHEET nw on a.id = nw.id_pha
        inner join EPHA_T_APPROVER ta2 on a.id = ta2.id_pha and t2.id_pha = ta2.id_pha and t2.id_session = ta2.id_session and isnull(ta2.approver_action_type,0) <> 2
        left join VW_EPHA_PERSON_DETAILS emp on lower(ta2.user_name) = lower(emp.user_name)
        left join VW_EPHA_PERSON_DETAILS emp_ori on lower(a.pha_request_by) = lower(emp_ori.user_name)
        left join VW_EPHA_PERSON_DETAILS emp_conso on lower(nw.responder_user_name) = lower(emp_conso.user_name)
        where a.pha_status in (21) and a.request_approver = 1
          and a.seq in (select max(seq) from epha_t_header group by pha_no)";

            if (!home_task)
            {
                sqlstr += @"
            union

            select a.id as id_pha, a.pha_status,
                   isnull(emp.user_name,'') as user_name, emp.user_displayname, emp.user_email,
                   isnull(a.pha_request_by,'') as user_name_ori,
                   0 as id_action, null as user_action_date,
                   6 as action_sort,
                   0 as task, upper(a.pha_sub_software) as pha_type,
                   'Approver Approve' as action_required, a.pha_no as document_number, g.pha_request_name as document_title,
                   nw.recommendations_no as rev_def, a.pha_version_text + '('+  a.pha_version_desc +')' as rev, emp_ori.user_displayname as originator,
                   format(nw.responder_action_date,'dd MMM yyyy') as receivesd, format(g.target_end_date,'dd MMM yyyy') as due_date,
                   null as action_date,
                   isnull(datediff(day, getdate(),case when g.target_end_date >= getdate() then g.target_end_date else getdate() end),0) as remaining,
                   emp_conso.user_displayname as consolidator
            from epha_t_header a
            inner join EPHA_T_GENERAL g on a.id = g.id_pha
            inner join EPHA_T_SESSION s on a.id = s.id_pha
            inner join (select max(id) as id_session, id_pha from EPHA_T_SESSION group by id_pha) s2 on a.id = s2.id_pha and s.id = s2.id_session
            inner join (select max(id_session) as id_session, id_pha from EPHA_T_MEMBER_TEAM group by id_pha) t2 on a.id = t2.id_pha and s2.id_session = t2.id_session
            inner join EPHA_T_NODE_WORKSHEET nw on a.id = nw.id_pha
            inner join EPHA_T_APPROVER ta2 on a.id = ta2.id_pha and t2.id_pha = ta2.id_pha and t2.id_session = ta2.id_session and isnull(ta2.approver_action_type,0) <> 2
            left join VW_EPHA_PERSON_DETAILS emp on lower(ta2.user_name) = lower(emp.user_name)
            left join VW_EPHA_PERSON_DETAILS emp_ori on lower(a.pha_request_by) = lower(emp_ori.user_name)
            left join VW_EPHA_PERSON_DETAILS emp_conso on lower(nw.responder_user_name) = lower(emp_conso.user_name)
            where a.pha_status in (21) and a.request_approver = 1
              and a.seq in (select max(seq) from epha_t_header group by pha_no)

            union

            select a.id as id_pha, a.pha_status,
                   isnull(a.pha_request_by,'') as user_name, emp.user_displayname, emp.user_email,
                   isnull(a.pha_request_by,'') as user_name_ori,
                   0 as id_action, null as user_action_date,
                   9 as action_sort,
                   0 as task, upper(a.pha_sub_software) as pha_type,
                   'For Info' as action_required, a.pha_no as document_number, g.pha_request_name as document_title,
                   nw.recommendations_no as rev_def, a.pha_version_text + '('+  a.pha_version_desc +')' as rev, emp_ori.user_displayname as originator,
                   format(nw.responder_action_date,'dd MMM yyyy') as receivesd, format(g.target_end_date,'dd MMM yyyy') as due_date,
                   null as action_date,
                   isnull(datediff(day, getdate(),case when g.target_end_date >= getdate() then g.target_end_date else getdate() end),0) as remaining,
                   emp_conso.user_displayname as consolidator
            from epha_t_header a
            inner join EPHA_T_GENERAL g on a.id = g.id_pha
            inner join EPHA_T_SESSION s on a.id = s.id_pha
            inner join (select max(id) as id_session, id_pha from EPHA_T_SESSION group by id_pha) s2 on a.id = s2.id_pha and s.id = s2.id_session
            inner join (select max(id_session) as id_session, id_pha from EPHA_T_MEMBER_TEAM group by id_pha) t2 on a.id = t2.id_pha and s2.id_session = t2.id_session
            inner join EPHA_T_NODE_WORKSHEET nw on a.id = nw.id_pha
            left join VW_EPHA_PERSON_DETAILS emp on lower(a.pha_request_by) = lower(emp.user_name)
            left join VW_EPHA_PERSON_DETAILS emp_ori on lower(a.pha_request_by) = lower(emp_ori.user_name)
            left join VW_EPHA_PERSON_DETAILS emp_conso on lower(nw.responder_user_name) = lower(emp_conso.user_name)
            where a.pha_status in (91)
              and a.seq in (select max(seq) from epha_t_header group by pha_no)";
            }

            if (group_by_user)
            {
                sqlstr = "select distinct t.user_name, t.user_displayname, t.user_email from (" + sqlstr + ") t where 1=1";

                if (!string.IsNullOrEmpty(user_name))
                {
                    sqlstr += " and lower(t.user_name) = lower(@user_name)";
                    parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name });
                }

                if (!string.IsNullOrEmpty(seq))
                {
                    sqlstr += " and t.id_pha = @seq";
                    parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq });
                }

                sqlstr += " order by t.user_name";
            }
            else
            {
                if (home_task)
                {
                    sqlstr = "select distinct id_pha, pha_status, user_name, user_displayname, user_email, user_name_ori, id_action, user_action_date, action_sort, task, pha_type, action_required, document_number, document_title, rev, originator, receivesd, due_date, action_date, remaining, consolidator from (" + sqlstr + ") t where 1=1";
                }
                else
                {
                    sqlstr = "select distinct t.* from (" + sqlstr + ") t where 1=1";
                }

                if (!string.IsNullOrEmpty(user_name))
                {
                    sqlstr += " and lower(t.user_name) = lower(@user_name)";
                    parameters.Add(new SqlParameter("@user_name", SqlDbType.VarChar, 50) { Value = user_name });
                }

                if (!string.IsNullOrEmpty(seq))
                {
                    sqlstr += " and t.id_pha = @seq";
                    parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq });
                }

                sqlstr += " order by t.user_name, t.action_sort, t.document_number, t.rev";
            }

            // Replace table names for different sub_software
            if (sub_software == "jsea")
            {
                sqlstr = sqlstr.Replace("EPHA_T_NODE_WORKSHEET", "EPHA_T_TASKS_WORKSHEET");
            }
            else if (sub_software == "whatif")
            {
                sqlstr = sqlstr.Replace("EPHA_T_NODE_WORKSHEET", "EPHA_T_LIST_WORKSHEET");
            }
            else if (sub_software == "hra")
            {
                sqlstr = sqlstr.Replace("EPHA_T_NODE_WORKSHEET", @"
                   select ws.id, ws.id_pha, ws.id_tasks, ws.action_status,
                   wk.user_name as responder_user_name, ws.recommendations_no, ws.recommendations,
                   ws.responder_receivesd_date, ws.responder_action_date, ws.estimated_start_date
            from vw_epha_max_seq_by_pha_no smx
            inner join epha_t_table3_worksheet ws on smx.id_pha = ws.id_pha
            inner join epha_t_table2_workers wk on smx.id_pha = wk.id_pha and wk.id_tasks = ws.id_tasks");
            }
             
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);

            #endregion get data 

            return dt;
        }
         
        public DataTable DataDailyByActionRequired_Responder(string seq, string sub_software, Boolean group_by_user, string seq_worksheet_list)
        {
            DataTable dt = new DataTable(); 
            var parameters = new List<SqlParameter>();

            #region get data 
            //Recommendation Closing -> Waiting Follow Up --> 13: Waiting Follow Up  
            sqlstr = @"
        select distinct a.id as id_pha, a.pha_status,
                        isnull(nw.responder_user_name, '') as user_name, emp.user_displayname, emp.user_email,
                        isnull(a.pha_request_by, '') as user_name_ori,
                        0 as id_action, nw.responder_action_date as user_action_date,
                        1 as action_sort,
                        0 as task, upper(a.pha_sub_software) as pha_type,
                        'Recommendation Closing' as action_required, a.pha_no as document_number, g.pha_request_name as document_title,
                        nw.recommendations_no as rev_def, a.pha_version_text + '(' + a.pha_version_desc + ')' as rev, emp_ori.user_displayname as originator,
                        format(nw.responder_receivesd_date, 'dd MMM yyyy') as receivesd, format(nw.estimated_start_date, 'dd MMM yyyy') as due_date,
                        nw.responder_action_date as action_date,
                        isnull(datediff(day, getdate(), case when nw.estimated_start_date >= getdate() then nw.estimated_start_date else getdate() end), 0) as remaining,
                        emp_conso.user_displayname as consolidator,
                        nw.seq as seq_worksheet_list
        from epha_t_header a
        inner join EPHA_T_GENERAL g on a.id = g.id_pha
        inner join (select max(id) as id_session, id_pha from EPHA_T_SESSION group by id_pha) s on a.id = s.id_pha
        left join (select max(id_session) as id_session, id_pha from EPHA_T_MEMBER_TEAM group by id_pha) t2 on a.id = t2.id_pha and s.id_session = t2.id_session
        inner join EPHA_T_NODE_WORKSHEET nw on a.id = nw.id_pha
        left join VW_EPHA_PERSON_DETAILS emp on lower(nw.responder_user_name) = lower(emp.user_name)
        left join VW_EPHA_PERSON_DETAILS emp_ori on lower(a.pha_request_by) = lower(emp_ori.user_name)
        left join VW_EPHA_PERSON_DETAILS emp_conso on lower(nw.responder_user_name) = lower(emp_conso.user_name)
        where nw.responder_user_name is not null and nw.estimated_start_date is not null and nw.responder_action_date is null
          and a.seq in (select max(seq) from epha_t_header group by pha_no) and a.pha_status = 13";

            // Filter by seq_worksheet_list if provided
            if (!string.IsNullOrEmpty(seq_worksheet_list))
            {
                sqlstr += " and nw.seq in (" + seq_worksheet_list + ")";
            }

            if (group_by_user)
            {
                sqlstr = "select distinct t.user_name, t.user_displayname, t.user_email from (" + sqlstr + ") t where 1=1";

                if (!string.IsNullOrEmpty(seq))
                {
                    sqlstr += " and t.id_pha = @seq";
                    parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq });
                }

                sqlstr += " order by t.user_name";
            }
            else
            {
                sqlstr = "select distinct t.* from (" + sqlstr + ") t where 1=1";

                if (!string.IsNullOrEmpty(seq))
                {
                    sqlstr += " and t.id_pha = @seq";
                    parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq });
                }

                sqlstr += " order by t.user_name, t.action_sort, t.document_number, t.rev";
            }

            // Replace table name for different sub_software
            if (sub_software == "jsea")
            {
                sqlstr = sqlstr.Replace("EPHA_T_NODE_WORKSHEET", "EPHA_T_TASKS_WORKSHEET");
            }
            else if (sub_software == "whatif")
            {
                sqlstr = sqlstr.Replace("EPHA_T_NODE_WORKSHEET", "EPHA_T_LIST_WORKSHEET");
            }
            else if (sub_software == "hra")
            {
                sqlstr = sqlstr.Replace("EPHA_T_NODE_WORKSHEET", @"
            select ws.id, ws.id_pha, ws.id_tasks, ws.action_status,
                   wk.user_name as responder_user_name, ws.recommendations_no, ws.recommendations,
                   ws.responder_receivesd_date, ws.responder_action_date, ws.estimated_start_date
            from vw_epha_max_seq_by_pha_no smx
            inner join epha_t_table3_worksheet ws on smx.id_pha = ws.id_pha
            inner join epha_t_table2_workers wk on smx.id_pha = wk.id_pha and wk.id_tasks = ws.id_tasks");
            }
             
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);

            #endregion get data 

            return dt;
        }

        public DataTable DataDailyByActionRequired_TeammMember_TA2eMOC(string user_name, string seq, string sub_software, Boolean group_by_user)
        {
            DataTable dt = new DataTable();  
            var parameters = new List<SqlParameter>();

            #region get data  
            sqlstr = @"
        SELECT DISTINCT a.id AS id_pha, a.pha_status,
                        ISNULL(t.user_name, '') AS user_name, emp.user_displayname, emp.user_email,
                        ISNULL(a.pha_request_by, '') AS user_name_ori,
                        t.seq AS id_action, t.date_review AS user_action_date,
                        2 AS action_sort, 0 AS task, UPPER(a.pha_sub_software) AS pha_type,
                        'Reviewer' AS action_required, a.pha_no AS document_number, g.pha_request_name AS document_title,
                        a.pha_version_text + '(' + a.pha_version_desc + ')' AS rev, emp_ori.user_displayname AS originator,
                        FORMAT(s.date_to_review, 'dd MMM yyyy') AS receivesd, FORMAT(DATEADD(day, 5, s.date_to_review), 'dd MMM yyyy') AS due_date,
                        t.date_review AS action_date,
                        ISNULL(DATEDIFF(day, GETDATE(), CASE WHEN (DATEADD(day, 5, s.date_to_review)) >= GETDATE() THEN DATEADD(day, 5, s.date_to_review) ELSE GETDATE() END), 0) AS remaining,
                        emp_conso.user_displayname AS consolidator
        FROM epha_t_header a
        INNER JOIN EPHA_T_GENERAL g ON a.id = g.id_pha
        INNER JOIN EPHA_T_SESSION s ON a.id = s.id_pha
        INNER JOIN (SELECT MAX(id) AS id_session, id_pha FROM EPHA_T_SESSION GROUP BY id_pha) s2 ON a.id = s2.id_pha AND s.id = s2.id_session
        INNER JOIN EPHA_T_MEMBER_TEAM t ON a.id = t.id_pha AND s2.id_session = t.id_session
        INNER JOIN (SELECT MAX(id_session) AS id_session, id_pha FROM EPHA_T_MEMBER_TEAM GROUP BY id_pha) t2 ON a.id = t2.id_pha AND s2.id_session = t2.id_session
        LEFT JOIN VW_EPHA_PERSON_DETAILS emp ON LOWER(t.user_name) = LOWER(emp.user_name)
        LEFT JOIN VW_EPHA_PERSON_DETAILS emp_ori ON LOWER(a.pha_request_by) = LOWER(emp_ori.user_name)
        LEFT JOIN VW_EPHA_PERSON_DETAILS emp_conso ON LOWER(t.user_name) = LOWER(emp_conso.user_name)
        WHERE t.user_name IS NOT NULL
          AND t.pha_status >= 12";

            if (!string.IsNullOrEmpty(seq))
            {
                sqlstr += " AND t.id_pha = @seq";
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq });
            }

            if (group_by_user)
            {
                sqlstr = "SELECT DISTINCT t.user_name, t.user_displayname, t.user_email FROM (" + sqlstr + ") t ORDER BY t.user_name";
            }
            else
            {
                sqlstr = "SELECT DISTINCT t.* FROM (" + sqlstr + ") t ORDER BY t.user_name, t.action_sort, t.document_number, t.rev";
            }

            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);

            #endregion get data

            return dt;
        }

        public DataTable DataDailyByActionRequired_TeammMember(string user_name, string seq, string sub_software, Boolean group_by_user)
        {
            DataTable dt = new DataTable();  
            var parameters = new List<SqlParameter>();

            #region get data  
            sqlstr = @"
        SELECT DISTINCT a.id AS id_pha, a.pha_status,
                        ISNULL(t.user_name, '') AS user_name, emp.user_displayname, emp.user_email,
                        ISNULL(a.pha_request_by, '') AS user_name_ori,
                        t.seq AS id_action, t.date_review AS user_action_date,
                        2 AS action_sort, 0 AS task, UPPER(a.pha_sub_software) AS pha_type,
                        'Reviewer' AS action_required, a.pha_no AS document_number, g.pha_request_name AS document_title,
                        a.pha_version_text + '(' + a.pha_version_desc + ')' AS rev, emp_ori.user_displayname AS originator,
                        FORMAT(GETDATE(), 'dd MMM yyyy') AS receivesd, FORMAT(DATEADD(day, 5, s.meeting_date), 'dd MMM yyyy') AS due_date,
                        t.date_review AS action_date,
                        ISNULL(DATEDIFF(day, GETDATE(), CASE WHEN (DATEADD(day, 5, s.meeting_date)) >= GETDATE() THEN DATEADD(day, 5, s.meeting_date) ELSE GETDATE() END), 0) AS remaining,
                        emp_conso.user_displayname AS consolidator
        FROM epha_t_header a
        INNER JOIN EPHA_T_GENERAL g ON a.id = g.id_pha
        INNER JOIN EPHA_T_SESSION s ON a.id = s.id_pha
        INNER JOIN (SELECT MAX(id) AS id_session, id_pha FROM EPHA_T_SESSION GROUP BY id_pha) s2 ON a.id = s2.id_pha AND s.id = s2.id_session
        INNER JOIN EPHA_T_MEMBER_TEAM t ON a.id = t.id_pha AND s2.id_session = t.id_session
        INNER JOIN (SELECT MAX(id_session) AS id_session, id_pha FROM EPHA_T_MEMBER_TEAM GROUP BY id_pha) t2 ON a.id = t2.id_pha AND s2.id_session = t2.id_session
        LEFT JOIN VW_EPHA_PERSON_DETAILS emp ON LOWER(t.user_name) = LOWER(emp.user_name)
        LEFT JOIN VW_EPHA_PERSON_DETAILS emp_ori ON LOWER(a.pha_request_by) = LOWER(emp_ori.user_name)
        LEFT JOIN VW_EPHA_PERSON_DETAILS emp_conso ON LOWER(t.user_name) = LOWER(emp_conso.user_name)
        WHERE t.user_name IS NOT NULL
          AND t.pha_status >= 12";

            if (!string.IsNullOrEmpty(seq))
            {
                sqlstr += " AND t.id_pha = @seq";
                parameters.Add(new SqlParameter("@seq", SqlDbType.VarChar, 50) { Value = seq });
            }

            if (group_by_user)
            {
                sqlstr = "SELECT DISTINCT t.user_name, t.user_displayname, t.user_email FROM (" + sqlstr + ") t ORDER BY t.user_name";
            }
            else
            {
                sqlstr = "SELECT DISTINCT t.* FROM (" + sqlstr + ") t ORDER BY t.user_name, t.action_sort, t.document_number, t.rev";
            }

            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);

            #endregion get data

            return dt;
        }

        public DataTable DataDailyByActionRequired_ReviewApprove(string id_pha, string responder_user_name, string sub_software, Boolean group_by_user, Boolean responder_close_all)
        {
            DataTable dt = new DataTable(); 
            var parameters = new List<SqlParameter>();

            #region get data  
            sqlstr = @"
        SELECT a.id AS id_pha, a.pha_status,
               ISNULL(nw.responder_user_name, '') AS user_name, emp_conso.user_displayname, emp_conso.user_email,
               ISNULL(a.pha_request_by, '') AS user_name_ori,
               nw.seq AS id_action, NULL AS user_action_date,
               4 AS action_sort, 0 AS task, UPPER(a.pha_sub_software) AS pha_type,
               'Review & Approve' AS action_required, a.pha_no AS document_number, g.pha_request_name AS document_title,
               nw.recommendations_no AS rev_def, a.pha_version_text + '(' + a.pha_version_desc + ')' AS rev, emp_ori.user_displayname AS originator,
               FORMAT(nw.responder_action_date, 'dd MMM yyyy') AS receivesd, FORMAT(g.target_end_date, 'dd MMM yyyy') AS due_date,
               NULL AS action_date,
               ISNULL(DATEDIFF(day, GETDATE(), CASE WHEN g.target_end_date >= GETDATE() THEN g.target_end_date ELSE GETDATE() END), 0) AS remaining,
               emp_conso.user_displayname AS consolidator
        FROM epha_t_header a
        INNER JOIN EPHA_T_GENERAL g ON a.id = g.id_pha
        INNER JOIN EPHA_T_SESSION s ON a.id = s.id_pha
        INNER JOIN (SELECT MAX(id) AS id_session, id_pha FROM EPHA_T_SESSION GROUP BY id_pha) s2 ON a.id = s2.id_pha AND s.id = s2.id_session
        INNER JOIN (SELECT MAX(id_session) AS id_session, id_pha FROM EPHA_T_MEMBER_TEAM GROUP BY id_pha) t2 ON a.id = t2.id_pha AND s2.id_session = t2.id_session
        INNER JOIN EPHA_T_NODE_WORKSHEET nw ON a.id = nw.id_pha
        LEFT JOIN VW_EPHA_PERSON_DETAILS emp ON LOWER(a.pha_request_by) = LOWER(emp.user_name)
        LEFT JOIN VW_EPHA_PERSON_DETAILS emp_ori ON LOWER(a.pha_request_by) = LOWER(emp_ori.user_name)
        LEFT JOIN VW_EPHA_PERSON_DETAILS emp_conso ON LOWER(nw.responder_user_name) = LOWER(emp_conso.user_name)
        WHERE nw.responder_user_name IS NOT NULL AND nw.responder_action_date IS NOT NULL AND LOWER(nw.action_status) NOT IN ('closed')
          AND a.seq IN (SELECT MAX(seq) FROM epha_t_header GROUP BY pha_no)";

            if (responder_close_all)
            {
                sqlstr += " AND a.pha_status IN (14)";
            }
            else
            {
                sqlstr += " AND a.pha_status IN (13)";
            }

            if (group_by_user)
            {
                sqlstr = "SELECT DISTINCT t.user_name, t.user_displayname, t.user_email FROM (" + sqlstr + ") t WHERE 1=1";
                if (!string.IsNullOrEmpty(responder_user_name))
                {
                    sqlstr += " AND LOWER(t.user_name) = LOWER(@responder_user_name)";
                    parameters.Add(new SqlParameter("@responder_user_name", SqlDbType.VarChar, 50) { Value = responder_user_name });
                }
                if (!string.IsNullOrEmpty(id_pha))
                {
                    sqlstr += " AND LOWER(t.id_pha) = LOWER(@id_pha)";
                    parameters.Add(new SqlParameter("@id_pha", SqlDbType.VarChar, 50) { Value = id_pha });
                }
                sqlstr += " ORDER BY t.user_name";
            }
            else
            {
                sqlstr = "SELECT DISTINCT t.* FROM (" + sqlstr + ") t WHERE 1=1";
                if (!string.IsNullOrEmpty(responder_user_name))
                {
                    sqlstr += " AND LOWER(t.user_name) = LOWER(@responder_user_name)";
                    parameters.Add(new SqlParameter("@responder_user_name", SqlDbType.VarChar, 50) { Value = responder_user_name });
                }
                if (!string.IsNullOrEmpty(id_pha))
                {
                    sqlstr += " AND LOWER(t.id_pha) = LOWER(@id_pha)";
                    parameters.Add(new SqlParameter("@id_pha", SqlDbType.VarChar, 50) { Value = id_pha });
                }
                sqlstr += " ORDER BY t.user_name, t.action_sort, t.document_number, t.rev";
            }

            // Replace the node worksheet based on sub_software
            if (sub_software == "jsea")
            {
                sqlstr = sqlstr.Replace("EPHA_T_NODE_WORKSHEET", "EPHA_T_TASKS_WORKSHEET");
            }
            else if (sub_software == "whatif")
            {
                sqlstr = sqlstr.Replace("EPHA_T_NODE_WORKSHEET", "EPHA_T_LIST_WORKSHEET");
            }
            else if (sub_software == "hra")
            {
                sqlstr = sqlstr.Replace("EPHA_T_NODE_WORKSHEET", @"
            SELECT ws.id, ws.id_pha, ws.id_tasks, ws.action_status,
                   wk.user_name AS responder_user_name, ws.recommendations_no, ws.recommendations,
                   ws.responder_receivesd_date, ws.responder_action_date, ws.estimated_start_date
            FROM vw_epha_max_seq_by_pha_no smx
            INNER JOIN epha_t_table3_worksheet ws ON smx.id_pha = ws.id_pha
            INNER JOIN epha_t_table2_workers wk ON smx.id_pha = wk.id_pha AND wk.id_tasks = ws.id_tasks");
            }
             
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);

            #endregion get data

            return dt;
        }

        public DataTable DataDailyByActionRequired_Closed(string id_pha, string sub_software, Boolean group_by_user, Boolean task_noti)
        {
            DataTable dt = new DataTable(); 
            var parameters = new List<SqlParameter>();

            #region get data  
            sqlstr = @"
        SELECT a.id AS id_pha, a.pha_status,
               ISNULL(a.pha_request_by, '') AS user_name, emp.user_displayname, emp.user_email,
               ISNULL(a.pha_request_by, '') AS user_name_ori,
               nw.seq AS id_action, NULL AS user_action_date,
               6 AS action_sort, 0 AS task, UPPER(a.pha_sub_software) AS pha_type,
               'For Info' AS action_required, a.pha_no AS document_number, g.pha_request_name AS document_title,
               nw.recommendations_no AS rev_def, a.pha_version_text + '(' + a.pha_version_desc + ')' AS rev, emp_ori.user_displayname AS originator,
               FORMAT(nw.responder_action_date, 'dd MMM yyyy') AS receivesd, FORMAT(g.target_end_date, 'dd MMM yyyy') AS due_date,
               NULL AS action_date,
               ISNULL(DATEDIFF(day, GETDATE(), CASE WHEN g.target_end_date >= GETDATE() THEN g.target_end_date ELSE GETDATE() END), 0) AS remaining,
               emp_conso.user_displayname AS consolidator
        FROM epha_t_header a
        INNER JOIN EPHA_T_GENERAL g ON a.id = g.id_pha
        INNER JOIN EPHA_T_SESSION s ON a.id = s.id_pha
        INNER JOIN (SELECT MAX(id) AS id_session, id_pha FROM EPHA_T_SESSION GROUP BY id_pha) s2 ON a.id = s2.id_pha AND s.id = s2.id_session
        INNER JOIN (SELECT MAX(id_session) AS id_session, id_pha FROM EPHA_T_MEMBER_TEAM GROUP BY id_pha) t2 ON a.id = t2.id_pha AND s2.id_session = t2.id_session
        INNER JOIN EPHA_T_NODE_WORKSHEET nw ON a.id = nw.id_pha
        LEFT JOIN VW_EPHA_PERSON_DETAILS emp ON LOWER(a.pha_request_by) = LOWER(emp.user_name)
        LEFT JOIN VW_EPHA_PERSON_DETAILS emp_ori ON LOWER(a.pha_request_by) = LOWER(emp_ori.user_name)
        LEFT JOIN VW_EPHA_PERSON_DETAILS emp_conso ON LOWER(nw.responder_user_name) = LOWER(emp_conso.user_name)
        WHERE a.pha_status IN (91) AND nw.responder_user_name IS NOT NULL AND nw.reviewer_action_date IS NOT NULL
          AND a.seq IN (SELECT MAX(seq) FROM epha_t_header GROUP BY pha_no)";

            if (task_noti)
            {
                sqlstr += @"
            UNION
            SELECT a.id AS id_pha, a.pha_status,
                   ISNULL(a.pha_request_by, '') AS user_name, emp.user_displayname, emp.user_email,
                   ISNULL(a.pha_request_by, '') AS user_name_ori,
                   nw.seq AS id_action, NULL AS user_action_date,
                   6 AS action_sort, 0 AS task, UPPER(a.pha_sub_software) AS pha_type,
                   'For Info' AS action_required, a.pha_no AS document_number, g.pha_request_name AS document_title,
                   nw.recommendations_no AS rev_def, a.pha_version_text + '(' + a.pha_version_desc + ')' AS rev, emp_ori.user_displayname AS originator,
                   FORMAT(nw.responder_action_date, 'dd MMM yyyy') AS receivesd, FORMAT(g.target_end_date, 'dd MMM yyyy') AS due date,
                   NULL AS action_date,
                   ISNULL(DATEDIFF(day, GETDATE(), CASE WHEN g.target_end_date >= GETDATE() THEN g.target_end_date ELSE GETDATE() END), 0) AS remaining,
                   emp_conso.user_displayname AS consolidator
            FROM epha_t_header a
            INNER JOIN EPHA_T_GENERAL g ON a.id = g.id_pha
            INNER JOIN EPHA_T_SESSION s ON a.id = s.id_pha
            INNER JOIN (SELECT MAX(id) AS id_session, id_pha FROM EPHA_T_SESSION GROUP BY id_pha) s2 ON a.id = s2.id_pha AND s.id = s2.id_session
            INNER JOIN (SELECT MAX(id_session) AS id_session, id_pha FROM EPHA_T_MEMBER_TEAM GROUP BY id_pha) t2 ON a.id = t2.id_pha AND s2.id_session = t2.id_session
            INNER JOIN EPHA_T_NODE_WORKSHEET nw ON a.id = nw.id_pha
            LEFT JOIN VW_EPHA_PERSON_DETAILS emp ON LOWER(a.pha_request_by) = LOWER(emp.user_name)
            LEFT JOIN VW_EPHA_PERSON_DETAILS emp_ori ON LOWER(a.pha_request_by) = LOWER(emp_ori.user_name)
            LEFT JOIN VW_EPHA_PERSON_DETAILS emp_conso ON LOWER(nw.responder_user_name) = LOWER(emp_conso.user_name)
            WHERE a.pha_status IN (91) AND nw.responder_user_name IS NOT NULL AND nw.reviewer_action_date IS NOT NULL
              AND a.seq IN (SELECT MAX(seq) FROM epha_t_header GROUP BY pha_no)";
            }

            if (group_by_user)
            {
                sqlstr = "SELECT DISTINCT t.user_name, t.user_displayname, t.user_email FROM (" + sqlstr + ") t WHERE 1=1";
                if (!string.IsNullOrEmpty(id_pha))
                {
                    sqlstr += " AND LOWER(t.id_pha) = LOWER(@id_pha)";
                    parameters.Add(new SqlParameter("@id_pha", SqlDbType.VarChar, 50) { Value = id_pha });
                }
                sqlstr += " ORDER BY t.user_name";
            }
            else
            {
                sqlstr = "SELECT DISTINCT t.* FROM (" + sqlstr + ") t WHERE 1=1";
                if (!string.IsNullOrEmpty(id_pha))
                {
                    sqlstr += " AND LOWER(t.id_pha) = LOWER(@id_pha)";
                    parameters.Add(new SqlParameter("@id_pha", SqlDbType.VarChar, 50) { Value = id_pha });
                }
                sqlstr += " ORDER BY t.user_name, t.action_sort, t.document_number, t.rev";
            }

            // Replace the node worksheet based on sub_software
            if (sub_software == "jsea")
            {
                sqlstr = sqlstr.Replace("EPHA_T_NODE_WORKSHEET", "EPHA_T_TASKS_WORKSHEET");
            }
            else if (sub_software == "whatif")
            {
                sqlstr = sqlstr.Replace("EPHA_T_NODE_WORKSHEET", "EPHA_T_LIST_WORKSHEET");
            }
            else if (sub_software == "hra")
            {
                sqlstr = sqlstr.Replace("EPHA_T_NODE_WORKSHEET", @"
            SELECT ws.id, ws.id_pha, ws.id_tasks, ws.action_status,
                   wk.user_name AS responder_user_name, ws.recommendations_no, ws.recommendations,
                   ws.responder_receivesd_date, ws.responder_action_date, ws.estimated_start_date
            FROM vw_epha_max_seq_by_pha_no smx
            INNER JOIN epha_t_table3_worksheet ws ON smx.id_pha = ws.id_pha
            INNER JOIN epha_t_table2_workers wk ON smx.id_pha = wk.id_pha AND wk.id_tasks = ws.id_tasks");
            }
             
            dt = ClassConnectionDb.ExecuteAdapterSQLDataTable(sqlstr, parameters);

            #endregion get data

            return dt;
        }
         
    }
}
