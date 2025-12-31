<style>
/* แถบรายงานบรรทัดใหม่ใต้ Navbar */
.reports-sub-bar {
    display: none;
    width:90%;
    background: #f8fafc;
    border-top: 1px solid #e2e8f0;
    padding: 0.4rem 0;
    text-align: center;
    border-radius:10px;
    margin-top:0px;
}

.reports-sub-bar.active {
    display: block;
}

.sub-bar-content {
    max-width: 1200px;
    margin: 0 auto;
}

.sub-bar-content a {
    display: inline-block;
    margin: 0 1rem;
    //padding: 0.75rem 1.5rem;
    padding: 0.3rem 0.6rem;
    background: white;
    color: #1e3a8a;
    text-decoration: none;
    border-radius: 30px;
    font-weight: 600;
    font-size: 15px;
    transition: all 0.1s;
    box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    border: 1px solid #e2e8f0;
}

.sub-bar-content a:hover {
    background: #8dceff; !important;
    color: blue;
    transform: translateY(-3px);
    border: 1px solid #8888ff;
}

.sub-bar-content a.sub-active {
    background: #1e3a8a;
    color: white;
}

.sub-bar-content a.sub-disabled {
    color: #94a3b8;
    background: #f1f5f9;
    cursor: not-allowed;
}

/* ปุ่ม Reports ใน Navbar */
.reports-btn.active {
    background: #1e3a8a;
    color: white;
}
</style>

<div class="navbar" style="white-space:nowrap;">
    <a href="dashboard.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/dashboard.asp", "active", "") %>">
        <div class="icon">&#128065;</div>
        Dashboard
    </a>
    <a href="createjob.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/createjob.asp", "active", "") %>">
        <div class="icon">+</div>
        Create Job
    </a>
    <% If request.cookies("IsManager") Then %>
    <a href="meetingmode.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/meetingmode.asp", "active", "") %>">
        <div class="icon">&#127919;</div>
        Meeting Mode
    </a>
    <% End If %>
    <a href="alljobs.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/alljobs.asp", "active", "") %>">
        <div class="icon">&#8635;</div>
        Job WIP
    </a>
    <a href="completed.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/completed.asp", "active", "") %>">
        <div class="icon">&#10004;</div>
        Job Completed
    </a>
    <a id="reportsToggle" style="cursor:pointer;" class="reports-btn <%= IIf(InStr(LCase(Request.ServerVariables("SCRIPT_NAME")), "rep") > 0, "active", "") %>">
        <div class="icon">&#128196;</div>
        Report
    </a>
     <a href="fair_kpi.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/fair_kpi.asp", "active", "") %>">
        <div class="icon">&#10004;</div>
        Fair KPI System
    </a>
     <a href="view_problemtracking.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/view_problemtracking.asp", "active", "") %>">
        <div class="icon">&#10004;</div>
        Problem Tracking
    </a>
 
    <a href="logout.asp" style="" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/logout.asp", "active", "") %>">
        <div align=center class="icon" astyle="width:15px;">&#10140;</div>
        Logout
    </a>

<% If request.cookies("IsManager") or true Then %>
<div style="margin-top:-11px;height:1.05cm;width:90%;">
<div id="reportsSubBar" class="reports-sub-bar">
    <div class="sub-bar-content">
        <a href="rep_kpi_ranking.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/rep_kpi_ranking.asp", "sub-active", "") %>">
            KPI Ranking Dashboard
        </a>
        <a href="report_jobs_by_type.asp" class="<%= IIf(LCase(Request.ServerVariables("SCRIPT_NAME")) = "/report_jobs_by_type.asp", "sub-active", "") %>">
            Jobs by Type Summary
        </a>
        <!-- เพิ่มรายงานอื่นได้ที่นี่ -->
        <a href="#" class="sub-disabled">report 3</a>
        <a href="#" class="sub-disabled">report 4</a>
    </div>
</div>
</div>
<% End If %>

<script language=javascript>
    document.addEventListener('DOMContentLoaded', function () {
        var toggle = document.getElementById('reportsToggle');
        var subBar = document.getElementById('reportsSubBar');

        if (toggle && subBar) {
            toggle.addEventListener('click', function (e) {
                e.preventDefault();
                subBar.classList.toggle('active');
                toggle.classList.toggle('active');
            });

            document.addEventListener('click', function (e) {
                if (!toggle.contains(e.target) && !subBar.contains(e.target)) {
                    subBar.classList.remove('active');
                    toggle.classList.remove('active');
                }
            });
        }
    });
<% if Instr(LCase(Request.ServerVariables("SCRIPT_NAME")),"rep")>0 then  %>
  document.getElementById('reportsSubBar').style.display='block';
<% end if %>
</script>


</div>

