<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="windows-874">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Vinova v5.0 - Problem Tracking (ASP Classic)</title>
    <script src="/js/tailwind.js"></script>
</head>
<body class="bg-gray-100">
<!--#include file="includes/include.asp" -->
<!--#include file="db/connect.asp" -->

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
    
    ' We need to fetch problems.
    ' The example file had static data with categories like " Quality".
    ' We will map the Database 'Category' (varchar) to these visual styles.
    
    Dim rsProblems, pSQL
    pSQL = "SELECT p.ProblemID, p.UserID, u.Name as UserName, p.Title, p.Category, p.Severity, p.Description, p.Points, p.RaisedDate, p.Status " & _
           "FROM [vn].[ProblemTracking] p " & _
           "INNER JOIN [vn].[Users] u ON p.UserID = u.UserID " & _
           "ORDER BY p.RaisedDate DESC"
           
    Set rsProblems = conn.Execute(pSQL)
    
    ' Process Data into a list for easier handling
    Dim problemsList
    problemsList = Array()
    
    Do While Not rsProblems.EOF
        Dim pID, pUserID, pUserName, pTitle, pCategory, pSeverity, pDesc, pPoints, pDate, pStatus, pColor, pCatDisplay
        
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
        
        ' Logic to match Demo Visuals based on Data
        ' Map text categories to display icons/colors from example
        ' Logic to match Demo Visuals based on Data
        ' Map text categories to display icons/colors from example
        If InStr(pCategory, "Quality") > 0 Then
            pCatDisplay = "&#10024; Quality"
            pColor = "red"
        ElseIf InStr(pCategory, "Communication") > 0 Then
            pCatDisplay = "&#128483; Communication"
            pColor = "orange"
        ElseIf InStr(pCategory, "Time") > 0 Then
            pCatDisplay = "&#9203; Time Management"
            pColor = "orange"
        ElseIf InStr(pCategory, "Skills") > 0 Then
            pCatDisplay = "&#128170; Skills"
            pColor = "blue"
        ElseIf InStr(pCategory, "Recognition") > 0 Then
            pCatDisplay = "&#11088; Recognition"
            pColor = "yellow"
        Else
            ' Fallback for other categories
            pCatDisplay = pCategory
            If UCase(pSeverity) = "HIGH" Then pColor = "red"
            If UCase(pSeverity) = "MEDIUM" Then pColor = "orange"
            If UCase(pSeverity) = "LOW" Then pColor = "blue"
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
        
        rsProblems.MoveNext
    Loop
    rsProblems.Close
    
    ' Count totals for the top cards
    Dim pTotal, pScheduled, pResolved, idx
    pTotal = 0 ' Pending this meeting (All open in DB)
    pScheduled = 3 ' Mock value as per example
    pResolved = 12 ' Mock value as per example
    
    ' In real app, calculate these:
    ' pTotal = Count of problems where Status='Pending'
    ' pResolved = Count of problems where Status='Resolved'
    
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
                <button onclick="location.href='fair_kpi.asp'" class="flex-1 py-3 px-4 rounded-lg font-medium bg-gray-100 text-gray-700 hover:bg-gray-200">
                    &#128202; NEW: Fair KPI System
                </button>
                <button disabled class="flex-1 py-3 px-4 rounded-lg font-medium bg-blue-600 text-white cursor-default opacity-100">
                    &#128196; NEW: Problem Tracking
                </button>
            </div>
            
            <div id="content">
                
                <!-- Problem Overview -->
                <div class="grid md:grid-cols-3 gap-4 mb-6">
                    <div class="bg-red-50 rounded-lg p-6">
                        <div class="text-sm text-gray-600">Pending (This Meeting)</div>
                        <div class="text-4xl font-bold text-red-600"><%= UBound(problemsList) + 1 %></div>
                    </div>
                    <div class="bg-blue-50 rounded-lg p-6">
                        <div class="text-sm text-gray-600">Scheduled (Next)</div>
                        <div class="text-4xl font-bold text-blue-600"><%= pScheduled %></div>
                    </div>
                    <div class="bg-green-50 rounded-lg p-6">
                        <div class="text-sm text-gray-600">Resolved (Month)</div>
                        <div class="text-4xl font-bold text-green-600"><%= pResolved %></div>
                    </div>
                </div>
                
                <!-- Meeting Agenda -->
                <div class="bg-white rounded-xl shadow-lg p-6">
                    <div class="flex justify-between items-center mb-6">
                        <div>
                            <h2 class="text-2xl font-bold">Today's Meeting Agenda</h2>
                            <p class="text-sm text-gray-500"><%= Year(Now) %>/<%= Month(Now) %>/<%= Day(Now) %></p>
                        </div>
                        <button onclick="window.print()" class="px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700">
                            &#128196; Print Agenda
                        </button>
                    </div>
                    
                    <div class="space-y-6">
                        <div>
                            <h3 class="font-semibold text-red-600 mb-3">&#9888; URGENT (Discuss First):</h3>
                            <%
                            ' Filter High Severity
                            For idx = 0 To UBound(problemsList)
                                If problemsList(idx)("severity") = "HIGH" Then
                                    Call RenderProblemCard(problemsList(idx), 1)
                                End If
                            Next
                            %>
                        </div>
                        
                        <div>
                            <h3 class="font-semibold text-orange-600 mb-3">&#128196; PERFORMANCE REVIEW:</h3>
                            <%
                            ' Filter Medium Severity
                            For idx = 0 To UBound(problemsList)
                                If problemsList(idx)("severity") = "MEDIUM" Then
                                    Call RenderProblemCard(problemsList(idx), 2)
                                End If
                            Next
                            %>
                        </div>
                        
                        <div>
                            <h3 class="font-semibold text-indigo-600 mb-3">&#128170; TRAINING NEEDS:</h3>
                            <%
                            ' Filter Skills Category
                            For idx = 0 To UBound(problemsList)
                                If InStr(problemsList(idx)("category"), "Skills") > 0 Then
                                    Call RenderProblemCard(problemsList(idx), 4)
                                End If
                            Next
                            %>
                        </div>
                        
                        <div>
                            <h3 class="font-semibold text-yellow-600 mb-3">&#11088; RECOGNITION:</h3>
                            <%
                            ' Filter Recognition Category
                            For idx = 0 To UBound(problemsList)
                                If InStr(problemsList(idx)("category"), "Recognition") > 0 Then
                                    Call RenderProblemCard(problemsList(idx), 5)
                                End If
                            Next
                            %>
                        </div>
                    </div>
                    
                    <div class="mt-6 pt-6 border-t text-sm text-gray-600">
                        Estimated Time: <%= (UBound(problemsList) + 1) * 9 %> minutes
                    </div>
                </div>
                
                <!-- Action Buttons -->
                <div class="flex gap-4 mt-6">
                    <button onclick="location.href='new_problemtracking.asp';" class="px-6 py-3 bg-blue-600 text-white rounded-lg hover:bg-blue-700 font-medium">
                        + Add New Problem
                    </button>
                    <button onclick="alert('In production: Shows analytics dashboard')" class="px-6 py-3 bg-gray-200 text-gray-700 rounded-lg hover:bg-gray-300 font-medium">
                        View Analytics
                    </button>
                </div>
            </div>
        </div>
    </div>

    <%
    ' ==========================================
    ' HELPER SUBROUTINE
    ' ==========================================
    ' Note: We use 'f_' prefix for function parameters to avoid scope conflicts
    
    Sub RenderProblemCard(f_problem, f_index)
        Dim f_user, f_title, f_category, f_desc, f_color, f_severity, f_points, f_date, f_suggested
        Dim f_pointsClass
        
        f_user = f_problem("user")
        f_title = f_problem("title")
        f_category = f_problem("category")
        f_desc = f_problem("desc")
        f_color = f_problem("color")
        f_severity = f_problem("severity")
        f_points = f_problem("points")
        f_date = f_problem("date")
        
        ' Format Date
        If Not IsNull(f_date) Then
            f_date = Day(f_date) & "/" & Month(f_date) & "/" & Year(f_date)
        Else
            f_date = "N/A"
        End If
        
        ' Points styling
        f_suggested = ""
        If Not IsNull(f_points) And f_points <> 0 Then
            If f_points > 0 Then
                f_pointsClass = "text-green-600"
                f_suggested = "+" & f_points & " pts"
            Else
                f_pointsClass = "text-red-600"
                f_suggested = f_points & " pts"
            End If
        End If
    %>
        <div class="border-2 border-<%= f_color %>-200 bg-white rounded-lg p-4 mb-3">
            <div class="flex justify-between items-start">
                <div class="flex-1">
                    <div class="flex items-center gap-2 mb-2">
                        <span class="font-bold text-gray-600"><%= f_index %>.</span>
                        <div>
                            <div class="font-bold"><%= f_user %> - <%= f_title %></div>
                            <div class="text-sm text-gray-600">Category: <%= f_category %></div>
                        </div>
                    </div>
                    <div class="text-sm text-gray-700 ml-6 mb-2"><%= f_desc %></div>
                    <div class="flex items-center gap-4 ml-6 text-xs text-gray-500">
                        <span>Severity: <strong class="text-<%= f_color %>-600"><%= f_severity %></strong></span>
                        <span>Raised: <%= f_date %></span>
                        <% If f_suggested <> "" Then %>
                            <span class="font-semibold <%= f_pointsClass %>">Suggested: <%= f_suggested %></span>
                        <% End If %>
                    </div>
                </div>
                <div class="flex gap-2">
                    <button onclick="alert('In production: Opens discussion modal for <%= f_user %>')" class="px-3 py-1 bg-blue-600 text-white text-sm rounded hover:bg-blue-700">
                        Discuss
                    </button>
                    <button class="px-3 py-1 bg-gray-200 text-gray-700 text-sm rounded hover:bg-gray-300">
                        Skip
                    </button>
                </div>
            </div>
        </div>
    <%
    End Sub

    ' Cleanup
    conn.Close
    Set conn = Nothing
    %>
</body>
</html>