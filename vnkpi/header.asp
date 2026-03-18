<%
' ============================================================
' header.asp — Navigation bar + global CSS + user switcher
' Must be included AFTER: conn.asp, auth_check.asp, functions.asp
' Set activePage variable before including this file.
'
' User switcher card:
'   - Shown to managers only (curIsManager = True)
'   - Loads all active users from DB
'   - Matches prototype .user-btn CSS exactly
'   - Green border/color for managers, blue for staff
'   - Active button = currently viewed user (curUserID)
'   - Posts to switch_user.asp with TargetUserID + ReturnPage
' ============================================================
%>
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Vinova KPI System v5.0</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;background:#f3f4f6;padding:20px}
.container{max-width:1600px;margin:0 auto}
.header{background:linear-gradient(135deg,#667eea,#764ba2);color:white;padding:30px;border-radius:12px;margin-bottom:20px;box-shadow:0 4px 12px rgba(0,0,0,0.15)}
.card{background:white;padding:25px;margin:15px 0;border-radius:10px;box-shadow:0 2px 8px rgba(0,0,0,0.1)}
.btn{padding:12px 24px;margin:5px;border:none;border-radius:8px;cursor:pointer;font-size:14px;font-weight:600;transition:all 0.3s;display:inline-block;text-decoration:none}
.btn-primary{background:#3b82f6;color:white}.btn-primary:hover{background:#2563eb;transform:translateY(-2px)}
.btn-success{background:#10b981;color:white}.btn-success:hover{background:#059669;transform:translateY(-2px)}
.btn-danger{background:#ef4444;color:white}.btn-danger:hover{background:#dc2626;transform:translateY(-2px)}
.btn-warning{background:#f59e0b;color:white}.btn-warning:hover{background:#d97706;transform:translateY(-2px)}
.btn-secondary{background:#6b7280;color:white}.btn-secondary:hover{background:#4b5563}
.tabs{display:flex;gap:10px;margin:20px 0;overflow-x:auto;background:white;padding:10px;border-radius:10px;flex-wrap:wrap}
.tab{padding:12px 20px;border-radius:6px;cursor:pointer;white-space:nowrap;transition:all 0.3s;font-weight:500;text-decoration:none;color:#374151;display:inline-block}
.tab:hover{background:#f3f4f6}
.tab.active{background:#3b82f6;color:white;font-weight:700;box-shadow:0 2px 8px rgba(59,130,246,0.3)}
.tab.mgr-only{color:#7c3aed}
.tab.mgr-only.active{background:#7c3aed;color:white;box-shadow:0 2px 8px rgba(124,58,237,0.3)}
.progress{background:#e5e7eb;height:30px;border-radius:15px;overflow:hidden;margin:15px 0;position:relative}
.progress-bar{background:linear-gradient(90deg,#3b82f6,#8b5cf6);height:100%;transition:width 0.8s ease;display:flex;align-items:center;justify-content:center;color:white;font-weight:700}
table{width:100%;border-collapse:collapse;margin-top:15px}
th,td{padding:12px;text-align:left;border-bottom:1px solid #e5e7eb;font-size:14px}
th{background:#f9fafb;font-weight:700;color:#374151;font-size:12px;text-transform:uppercase;position:sticky;top:0}
tr:hover{background:#f9fafb}
.badge{padding:6px 12px;border-radius:20px;font-size:11px;font-weight:700;text-transform:uppercase;display:inline-block}
.tier-icon{font-size:60px;text-align:center;margin:20px 0;animation:float 3s ease-in-out infinite}
@keyframes float{0%,100%{transform:translateY(0)}50%{transform:translateY(-10px)}}
.stat-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:20px;margin:20px 0}
.stat-box{background:linear-gradient(135deg,#667eea,#764ba2);color:white;padding:20px;border-radius:10px;text-align:center;transition:transform 0.3s}
.stat-box:hover{transform:translateY(-5px)}
.stat-label{font-size:12px;opacity:0.9;text-transform:uppercase;letter-spacing:1px}
.stat-value{font-size:36px;font-weight:700;margin:10px 0}
.modal{display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.6);z-index:1000;align-items:center;justify-content:center}
.modal.active{display:flex}
.modal-content{background:white;padding:30px;border-radius:12px;max-width:600px;width:90%;max-height:90vh;overflow-y:auto;animation:slideIn 0.3s ease}
@keyframes slideIn{from{transform:translateY(-50px);opacity:0}to{transform:translateY(0);opacity:1}}
.form-group{margin:20px 0}
.form-group label{display:block;font-weight:600;margin-bottom:8px;color:#374151}
.form-group input,.form-group select,.form-group textarea{width:100%;padding:12px;border:2px solid #e5e7eb;border-radius:8px;font-size:14px;transition:border-color 0.3s}
.form-group input:focus,.form-group select:focus,.form-group textarea:focus{outline:none;border-color:#3b82f6;box-shadow:0 0 0 3px rgba(59,130,246,0.1)}
.jt--2{background:#7f1d1d;color:#fff;padding:6px 12px;border-radius:6px;font-size:11px;font-weight:700;display:inline-block}
.jt--1{background:#991b1b;color:#fff;padding:6px 12px;border-radius:6px;font-size:11px;font-weight:700;display:inline-block}
.jt-0 {background:#1e3a5f;color:#fff;padding:6px 12px;border-radius:6px;font-size:11px;font-weight:700;display:inline-block}
.jt-1 {background:#fee2e2;color:#991b1b;padding:6px 12px;border-radius:6px;font-size:11px;font-weight:700;display:inline-block}
.jt-2 {background:#fed7aa;color:#9a3412;padding:6px 12px;border-radius:6px;font-size:11px;font-weight:700;display:inline-block}
.jt-3 {background:#fef3c7;color:#92400e;padding:6px 12px;border-radius:6px;font-size:11px;font-weight:700;display:inline-block}
.jt-41{background:#dbeafe;color:#1e40af;padding:6px 12px;border-radius:6px;font-size:11px;font-weight:700;display:inline-block}
.jt-42{background:#f3e8ff;color:#6b21a8;padding:6px 12px;border-radius:6px;font-size:11px;font-weight:700;display:inline-block}
.jt-43{background:#fce7f3;color:#9f1239;padding:6px 12px;border-radius:6px;font-size:11px;font-weight:700;display:inline-block}
.jt-5 {background:#e0e7ff;color:#3730a3;padding:6px 12px;border-radius:6px;font-size:11px;font-weight:700;display:inline-block}
.overdue-row{background:#fee2e2 !important;animation:blink 2s infinite}
.target-today-row{background:#fff7ed !important;animation:pulse-orange 2s infinite}
@keyframes blink{0%,100%{opacity:1}50%{opacity:0.7}}
@keyframes pulse-orange{0%,100%{background:#fff7ed}50%{background:#fed7aa}}
.early-finish{background:#d1fae5 !important}
.private-badge{background:#dc2626;color:white;padding:4px 8px;border-radius:4px;font-size:10px;font-weight:700;animation:pulse 2s infinite}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:0.6}}
.shift-badge{background:#f59e0b;color:white;padding:4px 8px;border-radius:4px;font-size:10px;font-weight:700;margin-left:5px}
.alert{padding:15px;border-radius:8px;margin:15px 0;border-left:4px solid}
.alert-info{background:#dbeafe;color:#1e40af;border-left-color:#3b82f6}
.alert-warning{background:#fef3c7;color:#92400e;border-left-color:#f59e0b}
.alert-success{background:#d1fae5;color:#065f46;border-left-color:#10b981}
.alert-danger{background:#fee2e2;color:#991b1b;border-left-color:#dc2626}
.problem-card{border-left:4px solid #3b82f6;padding:15px;margin:10px 0;background:#f9fafb;border-radius:8px;transition:all 0.3s}
.problem-card:hover{box-shadow:0 4px 12px rgba(0,0,0,0.1);transform:translateX(5px)}
.problem-urgent{border-left-color:#dc2626;background:#fee2e2}
.problem-solved{border-left-color:#10b981;background:#d1fae5}
.kpi-adjust-history{max-height:400px;overflow-y:auto;margin-top:20px}
.close-btn{float:right;font-size:24px;cursor:pointer;color:#6b7280;line-height:1}
.close-btn:hover{color:#dc2626}

/* -- User switcher buttons — exact prototype .user-btn CSS -- */
.user-btn{margin:5px;padding:10px 20px;border:2px solid #3b82f6;background:white;color:#3b82f6;border-radius:8px;cursor:pointer;font-weight:600;transition:all 0.3s}
.user-btn:hover{background:#3b82f6;color:white;transform:scale(1.05)}
.user-btn.active{background:#3b82f6;color:white}
.user-btn.mgr-btn{border-color:#10b981;color:#10b981}
.user-btn.mgr-btn:hover{background:#10b981;color:white}
.user-btn.mgr-btn.active{background:#10b981;color:white}
</style>
</head>
<body>
<div class="container">

<!-- -- Header banner — exact prototype -- -->
<div class="header">
    <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px">
        <div>
            <h1 style="font-size:36px;margin-bottom:10px;font-weight:700">&#x1F3AF; Vinova Production System v5.0</h1>
            <p style="opacity:0.95;font-size:16px">Complete Job Tracking &bull; Fair KPI &bull; Meeting Mode &bull; Problem Management &bull; All Features Working</p>
        </div>
        <div style="text-align:right">
            <div style="font-size:15px;font-weight:600;margin-bottom:6px">
                <% If curIsManager Then %>
                &#x1F451; <%= safeStr(curRealName) %> <span style="background:rgba(255,255,255,0.2);padding:2px 10px;border-radius:12px;font-size:12px">Manager</span>
                <% Else %>
                &#x1F464; <%= safeStr(curUserName) %> <span style="background:rgba(255,255,255,0.2);padding:2px 10px;border-radius:12px;font-size:12px"><%= safeStr(curDept) %></span>
                <% End If %>
            </div>
            <a href="logout.asp" class="btn btn-secondary" style="padding:8px 18px;font-size:13px;margin:0">&#x1F6AA; Sign out</a>
        </div>
    </div>
</div>

<!-- ========================================================
     USER SWITCHER — manager only, matches prototype card
     Shows current "logged in as" line + all active users
     Clicking a user POSTs to switch_user.asp and reloads
======================================================== -->
<% If curIsManager Then %>
<div class="card">
    <%
    ' Show who we are currently viewing as
    Dim viewingLabel
    If curUserID = curRealUserID Then
        viewingLabel = safeStr(curRealName) & " (Manager - yourself)"
    Else
        viewingLabel = safeStr(curUserName) & " (" & safeStr(curDept) & ")"
    End If
    %>
    <h3 style="margin-bottom:15px">&#x1F464; Logged in as: <span style="color:#3b82f6"><%= viewingLabel %></span></h3>

    <div style="display:flex;flex-wrap:wrap;gap:10px">
        <%
        ' Load all active users from DB grouped: managers first, then staff
        ' Matches prototype order: managers (green), then staff (blue)
        Dim rsUsers
        Set rsUsers = objConn.Execute( _
            "SELECT UserID, Username, IsManager, ISNULL(Dept, Department) AS Dept " & _
            "FROM " & DB_SCHEMA & "Users " & _
            "WHERE IsActive = 1 " & _
            "ORDER BY IsManager DESC, Username ASC")

        Do While Not rsUsers.EOF
            Dim uID      : uID      = CLng(rsUsers("UserID"))
            Dim uName    : uName    = CStr(rsUsers("Username"))
            Dim uIsMgr   : uIsMgr   = CBool(rsUsers("IsManager"))
            Dim uDept    : uDept    = CStr(rsUsers("Dept"))
            Dim isActive : isActive = (uID = curUserID)

            ' Managers get green buttons, staff get blue — matches prototype
            Dim btnClass
            If uIsMgr Then
                btnClass = "user-btn mgr-btn" & iif(isActive, " active", "")
            Else
                btnClass = "user-btn" & iif(isActive, " active", "")
            End If

            ' Icon matches prototype:  for managers,  for staff
            Dim btnIcon
            If uIsMgr Then btnIcon = "&#x1F454;" Else btnIcon = "&#x1F468;&#x200D;&#x1F4BC;"

            ' Each button is a small form POSTing to switch_user.asp
            ' ReturnPage sends back to the current tab
        %>
        <form method="POST" action="switch_user.asp" style="display:inline;margin:0;padding:0">
            <input type="hidden" name="TargetUserID" value="<%= uID %>">
            <input type="hidden" name="ReturnPage"   value="<%= activePage %>.asp">
            <button type="submit" class="<%= btnClass %>">
                <%= btnIcon %> <%= safeStr(uName) %><% If uIsMgr Then %> (Manager)<% End If %>
            </button>
        </form>
        <% rsUsers.MoveNext : Loop %>
        <% rsUsers.Close : Set rsUsers = Nothing %>
    </div>
</div>
<% End If %>

<!-- -- Tab navigation — icons match prototype exactly -- -->
<div class="tabs">
    <a href="dashboard.asp"    class="tab <%= iif(activePage="dashboard", "active", "") %>">&#x1F4CA; Dashboard</a>
    <a href="myjobs.asp"       class="tab <%= iif(activePage="myjobs",    "active", "") %>">&#x1F4CB; My Jobs</a>
    <a href="assign_job.asp"   class="tab <%= iif(activePage="assignjob", "active", "") %>">&#x2795; Assign Job</a>
    <a href="assigned_list.asp" class="tab <%= iif(activePage="assigned", "active", "") %>">&#x1F4D1; My Assigned Jobs</a>
    <a href="approve_list.asp" class="tab <%= iif(activePage="approve",   "active", "") %>">&#x2705; Approve Submitted</a>
    <% If curIsManager Then %>
    <a href="team.asp"         class="tab mgr-only <%= iif(activePage="team",     "active", "") %>">&#x1F465; Team</a>
    <a href="problems.asp"     class="tab mgr-only <%= iif(activePage="problems", "active", "") %>">&#x1F4DD; Problems</a>
    <a href="meeting.asp"      class="tab mgr-only <%= iif(activePage="meeting",  "active", "") %>">&#x1F3AF; Meeting</a>
    <% End If %>
</div>
<script>
    (function () {
        function vinovaHeartbeat() {
            var xhr = new XMLHttpRequest();
            // The timestamp 't=' ensures the browser doesn't cache the request
            xhr.open('GET', 'ping.asp?t=' + new Date().getTime(), true);
            xhr.onreadystatechange = function () {
                if (xhr.readyState === 4) {
                    // Optional: You could log success here if debugging
                }
            };
            xhr.send();
        }

        // Run every 10 minutes
        setInterval(vinovaHeartbeat, 1800000);
    })();
</script>
