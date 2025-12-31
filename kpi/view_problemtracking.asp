<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="includes/include.asp" -->
<!--#include file="db/connect.asp" -->

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="windows-874">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Vinova v5.0 - Problem Tracking</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        /* Optional simple scrollbar styling */
        .card-scroll::-webkit-scrollbar { width: 6px; }
        .card-scroll::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 3px; }
    </style>
</head>
<body class="bg-gray-100 font-sans text-gray-800">

    <%
    ' ==========================================
    ' DATABASE CONFIGURATION
    ' ==========================================
    'Dim conn, connString
    'connString = "Provider=SQLNCLI11;Server=YOUR_SERVER_NAME;Database=vndb;Uid=YOUR_USERNAME;Pwd=YOUR_PASSWORD;"
    'Set conn = Server.CreateObject("ADODB.Connection")
    'conn.Open connString

    ' ==========================================
    ' DATA RETRIEVAL
    ' ==========================================
    Dim rsProblems, pSQL
    pSQL = "SELECT p.ProblemID, p.UserID, u.Name as UserName, p.Title, p.Category, p.Severity, p.Description, p.Points, p.RaisedDate, p.Status, p.JobID " & _
           "FROM [vn].[ProblemTracking] p " & _
           "INNER JOIN [vn].[Users] u ON p.UserID = u.UserID " & _
           "ORDER BY p.RaisedDate DESC"
    Set rsProblems = conn.Execute(pSQL)
    
    Dim problemsList
    problemsList = Array()
    
    Do While Not rsProblems.EOF
        Dim pID, pUserID, pUserName, pTitle, pCategory, pSeverity, pDesc, pPoints, pDate, pStatus, pColor, pCatDisplay, pJobID
        
        pID = rsProblems("ProblemID")
        pUserID = rsProblems("UserID")
        pUserName = rsProblems("UserName")
        pTitle = rsProblems("Title")
        pCategory = rsProblems("Category")
        pSeverity = rsProblems("Severity")
        pDesc = rsProblems("Description")
        pPoints = rsProblems("Points")
        pDate = rsProblems("RaisedDate")
        pStatus = rsProblems("Status")
        pJobID = rsProblems("JobID")
        
        ' Logic for Colors/Icons
        If InStr(pCategory, "Quality") > 0 Then 
          pCatDisplay = "&#10024; Quality" : pColor = "red"
        ElseIf InStr(pCategory, "Communication") > 0 Then 
          pCatDisplay = "&#128483; Communication" : pColor = "orange"
        ElseIf InStr(pCategory, "Time") > 0 Then 
          pCatDisplay = "&#9203; Time Management" : pColor = "orange"
        ElseIf InStr(pCategory, "Skills") > 0 Then 
          pCatDisplay = "&#128170; Skills" : pColor = "blue"
        ElseIf InStr(pCategory, "Recognition") > 0 Then 
          pCatDisplay = "&#11088; Recognition" : pColor = "yellow"
        Else 
            pCatDisplay = pCategory 
            If UCase(pSeverity) = "HIGH" Then 
              pColor = "red" 
            Else 
            If UCase(pSeverity) = "MEDIUM" Then 
              pColor = "orange" 
            Else 
              pColor = "blue" 
            End If 
            End If 
        End If

        ReDim Preserve problemsList(UBound(problemsList) + 1)
        Set problemsList(UBound(problemsList)) = Server.CreateObject("Scripting.Dictionary")
        problemsList(UBound(problemsList)).Add "id", pID
        problemsList(UBound(problemsList)).Add "user", pUserName
        problemsList(UBound(problemsList)).Add "title", pTitle
        problemsList(UBound(problemsList)).Add "category", pCatDisplay
        problemsList(UBound(problemsList)).Add "severity", UCase(pSeverity)
        problemsList(UBound(problemsList)).Add "color", pColor
        problemsList(UBound(problemsList)).Add "desc", pDesc
        problemsList(UBound(problemsList)).Add "points", pPoints
        problemsList(UBound(problemsList)).Add "date", pDate
        problemsList(UBound(problemsList)).Add "status", pStatus
        problemsList(UBound(problemsList)).Add "jobid", pJobID
        
        rsProblems.MoveNext
    Loop
    rsProblems.Close
    
    ' Counts
    Dim pTotal, pScheduled, pResolved, idx
    pTotal = UBound(problemsList) + 1
    
    Dim rsCounts, sqlCounts
    Set rsCounts = conn.Execute("SELECT COUNT(CASE WHEN Status='Pending' THEN 1 END) as P, COUNT(CASE WHEN Status='Resolved' THEN 1 END) as R FROM [vn].[ProblemTracking]")
    pScheduled = rsCounts("P")
    pResolved = rsCounts("R")
    rsCounts.Close
    %>
    
    <div id="app">
        <!-- Header -->
        <div class="bg-gradient-to-r from-blue-600 to-purple-600 text-white p-6 shadow-lg">
            <div class="max-w-6xl mx-auto">
                <h1 class="text-3xl font-bold">Vinova v5.0 - New Features Demo</h1>
                <p class="mt-2 text-blue-100">ASP Classic Integration - Problem Tracking</p>
            </div>
        </div>
        
        <!-- Navigation -->
        <div class="max-w-6xl mx-auto p-4">
            <div class="bg-white rounded-lg shadow-lg p-2 flex gap-2 mb-6">
                <a href="fair_kpi.asp" class="flex-1 py-3 px-4 rounded-lg font-medium bg-gray-100 text-gray-700 hover:bg-gray-200 text-center md:text-base text-sm">&#128202; Fair KPI</a>
                <a href="problem_list.asp" class="flex-1 py-3 px-4 rounded-lg font-medium bg-purple-600 text-white cursor-default text-center md:text-base text-sm">&#128196; Problem Tracking</a>
            </div>
            
            <div id="content">
                <!-- Stats Cards -->
                <div class="grid md:grid-cols-3 gap-4 mb-6">
                    <div class="bg-red-50 rounded-lg p-6 text-center border border-red-100">
                        <div class="text-sm text-gray-600">Pending</div><div class="text-4xl font-bold text-red-600"><%= pTotal %></div>
                    </div>
                    <div class="bg-blue-50 rounded-lg p-6 text-center border border-blue-100">
                        <div class="text-sm text-gray-600">Scheduled</div><div class="text-4xl font-bold text-blue-600"><%= pScheduled %></div>
                    </div>
                    <div class="bg-green-50 rounded-lg p-6 text-center border border-green-100">
                        <div class="text-sm text-gray-600">Resolved</div><div class="text-4xl font-bold text-green-600"><%= pResolved %></div>
                    </div>
                </div>
                
                <!-- Meeting Agenda -->
                <div class="bg-white rounded-xl shadow-lg p-6 mb-6">
                    <div class="flex flex-col md:flex-row justify-between items-center mb-6 gap-4">
                        <div>
                            <h2 class="text-2xl font-bold">Meeting Agenda</h2>
                            <p class="text-sm text-gray-500"><%= Year(Now) %>/<%= Month(Now) %>/<%= Day(Now) %></p>
                        </div>
                        <button onclick="window.print()" class="px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition">&#128196; Print</button>
                    </div>
                    
                    <div class="space-y-6">
                        <div>
                            <h3 class="font-semibold text-red-600 mb-3">&#9888; URGENT</h3>
                            <% For idx = 0 To UBound(problemsList) 
                            If problemsList(idx)("severity") = "HIGH" Then 
                              Call RenderProblemCard(problemsList(idx), 1) 
                            End If 
                            Next %>
                        </div>
                        <div>
                            <h3 class="font-semibold text-orange-600 mb-3">&#128196; PERFORMANCE</h3>
                            <% For idx = 0 To UBound(problemsList) 
                              If problemsList(idx)("severity") = "MEDIUM" Then 
                                Call RenderProblemCard(problemsList(idx), 2) 
                              End If 
                              Next %>
                        </div>
                        <div>
                            <h3 class="font-semibold text-indigo-600 mb-3">&#128170; TRAINING</h3>
                            <% For idx = 0 To UBound(problemsList) 
                              If InStr(problemsList(idx)("category"), "Skills") > 0 Then 
                                Call RenderProblemCard(problemsList(idx), 4) 
                              End If 
                              Next %>
                        </div>
                        <div>
                            <h3 class="font-semibold text-yellow-600 mb-3">&#11088; RECOGNITION</h3>
                            <% For idx = 0 To UBound(problemsList) 
                              If InStr(problemsList(idx)("category"), "Recognition") > 0 Then 
                                Call RenderProblemCard(problemsList(idx), 5) 
                              End If 
                              Next %>
                        </div>
                    </div>
                </div>
                
                <!-- Action Buttons -->
                <div class="flex gap-4 mt-6">
                    <a href="new_problemtracking.asp" class="w-full md:w-auto px-6 py-3 bg-blue-600 text-white rounded-lg hover:bg-blue-700 font-medium shadow-md transition">+ Add New Problem</a>
                </div>
            </div>
        </div>
    </div>

    <%
    Sub RenderProblemCard(f_problem, f_index)
        Dim f_user, f_title, f_category, f_desc, f_color, f_severity, f_points, f_date, f_suggested, f_pointsClass, f_status, f_jobid, f_id
        f_id = f_problem("id")
        f_user = f_problem("user")
        f_title = f_problem("title")
        f_category = f_problem("category")
        f_desc = f_problem("desc")
        f_color = f_problem("color")
        f_severity = f_problem("severity")
        f_points = f_problem("points")
        f_date = f_problem("date")
        f_status = f_problem("status")
        f_jobid = f_problem("jobid")
        
        If Not IsNull(f_date) Then f_date = Day(f_date) & "/" & Month(f_date) & "/" & Year(f_date) Else f_date = "N/A" End If
        
        f_suggested = ""
        If Not IsNull(f_points) And f_points <> 0 Then
            If f_points > 0 Then f_pointsClass = "text-green-600" : f_suggested = "+" & f_points & " pts" Else f_pointsClass = "text-red-600" : f_suggested = f_points & " pts" End If
        End If
    %>
        <div class="border-2 border-<%= f_color %>-200 bg-white rounded-lg p-4 mb-3 shadow-sm hover:shadow-md transition">
            <div class="flex justify-between items-start gap-4">
                <div class="flex-1">
                    <div class="flex items-center gap-2 mb-2">
                        <span class="font-bold text-gray-600"><%= f_index %>.</span>
                        <div>
                            <div class="font-bold text-base md:text-lg"><%= f_user %> - <%= f_title %></div>
                            <div class="text-xs md:text-sm text-gray-600">Category: <%= f_category %></div>
                        </div>
                    </div>
                    <div class="text-xs md:text-sm text-gray-700 ml-6 mb-2"><%= f_desc %></div>
                    <div class="flex flex-wrap items-center gap-2 md:gap-4 ml-6 text-xs text-gray-500">
                        <span>Severity: <strong class="text-<%= f_color %>-600"><%= f_severity %></strong></span>
                        <span>Raised: <%= f_date %></span>
                        <% If f_suggested <> "" Then %><span class="font-semibold <%= f_pointsClass %>"><%= f_suggested %></span><% End If %>
                        <span>Status: <%= f_status %></span>
                    </div>
                </div>
                
                <div class="flex flex-col md:flex-row gap-2 shrink-0">
                    <!-- EDIT BUTTON: Links to problem_edit.asp -->
                    <a href="edit_problemtracking.asp?id=<%= f_id %>" class="px-3 py-1 bg-indigo-50 text-indigo-600 border border-indigo-200 text-xs md:text-sm rounded hover:bg-indigo-100 transition font-medium text-center">
                        &#9998; Edit
                    </a>
                    
                    <button onclick="alert('Discuss modal placeholder')" class="px-3 py-1 bg-blue-600 text-white text-xs md:text-sm rounded hover:bg-blue-700 transition">Discuss</button>
                </div>
            </div>
        </div>
    <%
    End Sub

    conn.Close
    Set conn = Nothing
    %>
</body>
</html>