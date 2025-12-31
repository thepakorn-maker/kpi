<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="th">
<head>
    <meta charset="windows-874">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>KPI Ranking Report - Project Management Dashboard</title>
    <link rel="stylesheet" href="css/style.css">
    <style>
        .report-container {
            padding: 1.5rem;
        }
        .report-title {
            font-size: 28px;
            font-weight: 700;
            text-align: center;
            margin: 2rem 0;
            color: #1e3a8a;
        }
        .report-subtitle {
            text-align: center;
            color: #6b7280;
            margin-bottom: 3rem;
        }
        .ranking-table {
            width: 90%;
            max-width: 900px;
            margin: 0 auto;
            border-collapse: collapse;
            background: white;
            border-radius: 12px;
            overflow: hidden;
            box-shadow: 0 8px 20px rgba(0,0,0,0.1);
        }
        .ranking-table th {
            background: #1e3a8a;
            color: white;
            padding: 1.2rem;
            text-align: center;
            font-weight: 600;
            font-size: 16px;
        }
        .ranking-table td {
            padding: 1.2rem;
            text-align: center;
            border-bottom: 1px solid #e5e7eb;
        }
        .ranking-table tr:hover {
            background: #f0f9ff;
        }
        .rank-number {
            font-size: 28px;
            font-weight: 700;
            color: #f59e0b;
        }
        .kpi-excellent { color: #16a34a; font-weight: 700; font-size: 24px; }
        .kpi-good { color: #3b82f6; font-weight: 700; font-size: 24px; }
        .kpi-warning { color: #f59e0b; font-weight: 700; font-size: 24px; }
        .kpi-critical { color: #dc2626; font-weight: 700; font-size: 24px; }
        .legend {
            text-align: center;
            margin: 3rem 0;
            font-size: 16px;
        }
        .legend span {
            display: inline-block;
            margin: 0 1rem;
        }
    </style>
</head>
<body>
    <!--#include file="includes/include.asp" -->
    <!--#include file="includes/header.asp" -->
    <!--#include file="includes/navbar.asp" -->

    <% If not(request.cookies("IsManager")) and false Then Response.Redirect "dashboard.asp" %>

    <div class="report-container">
        <div class="report-title">
            &#128202; KPI Ranking Dashboard
        </div>
        <p class="report-subtitle">
            อันดับคะแนน KPI ของทีมทั้ง 14 คน (ข้อมูลเรียลไทม์ ณ วันที่ <%= FormatDateTime(Date(), 1) %>)
        </p>

        <table class="ranking-table">
            <thead>
                <tr>
                    <th style="width:15%">อันดับ</th>
                    <th style="width:45%; text-align:left; padding-left: 3rem;">ชื่อ - ตำแหน่ง</th>
                    <th style="width:20%">KPI Score</th>
                    <th style="width:20%">สถานะ</th>
                </tr>
            </thead>
            <tbody>
                <%
                Dim sqlRanking
                sqlRanking = "SELECT UserID, Name, Department, KPIScore FROM Users ORDER BY KPIScore DESC"
                
                Dim rsRanking
                Set rsRanking = conn.Execute(sqlRanking)
                
                Dim rank : rank = 0
                Do While Not rsRanking.EOF
                    rank = rank + 1
                    
                    Dim kpiClass, statusText
                    If rsRanking("KPIScore") >= 95 Then
                        kpiClass = "kpi-excellent"
                        statusText = "ดีมาก"
                    ElseIf rsRanking("KPIScore") >= 90 Then
                        kpiClass = "kpi-good"
                        statusText = "ดี"
                    ElseIf rsRanking("KPIScore") >= 80 Then
                        kpiClass = "kpi-warning"
                        statusText = "ควรปรับปรุง"
                    Else
                        kpiClass = "kpi-critical"
                        statusText = "ต้องปรับปรุงด่วน"
                    End If
                %>
                <tr>
                    <td class="rank-number"><%= rank %></td>
                    <td style="text-align:left; padding-left: 3rem;">
                        <strong><%= rsRanking("Name") %></strong><br>
                        <small style="color:#6b7280;"><%= rsRanking("Department") %></small>
                    </td>
                    <td class="<%= kpiClass %>"><%= rsRanking("KPIScore") %></td>
                    <td><%= statusText %></td>
                </tr>
                <%
                    rsRanking.MoveNext
                Loop
                rsRanking.Close
                Set rsRanking = Nothing
                %>
            </tbody>
        </table>
<div class="legend">
<strong>ระดับคะแนน KPI:</strong><br>
<span class="kpi-excellent">95+</span>&#8594;ดีมาก
<span class="kpi-good">90-94</span> &#8594;ดี
<span class="kpi-warning">80-89</span>&#8594;ควรปรับปรุง
<span class="kpi-critical">ต่ำกว่า 80</span> &#8594; ต้องปรับปรุงด่วน
</div>

    </div>

    <!-- Switch User -->
    <!--#include file="switch_user.asp" -->

    <% Call CloseConnection() %>
</body>
</html>