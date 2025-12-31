<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="includes/include.asp" -->
<!--#include file="db/connect.asp" -->

<%
' ==========================================
' HANDLE FORM SUBMISSION (INSERT LOGIC)
' ==========================================
Dim action, frmUserID, frmTitle, frmCategory, frmSeverity, frmDesc, frmPoints, frmJobID, sqlInsert

action = Request.Form("action")

If action = "add_problem" Then
    ' 1. Get Form Data
    frmUserID   = Request.Form("UserID")
    frmTitle    = Request.Form("Title")
    frmCategory = Request.Form("Category")
    frmSeverity = Request.Form("Severity")
    frmDesc     = Request.Form("Description")
    frmPoints   = Request.Form("Points")
    frmJobID    = Request.Form("JobID")
    
    ' 2. Validation (Basic)
    If frmTitle = "" Then
        ' Simple error handling (In production, redirect back with error msg)
        Response.Write("<script>alert('Title is required'); window.history.back();</script>")
        Response.End
    End If

    ' 3. Prepare SQL
    ' Note: Points is an Int, JobID is an Int. We convert empty strings to NULL for Integers.
    sqlInsert = "INSERT INTO [vn].[ProblemTracking] " & _
                "([UserID], [Title], [Category], [Severity], [Description], [Points], [RaisedDate], [Status], [JobID]) " & _
                "VALUES (" & _
                " " & frmUserID & ", " & _
                " '" & Replace(frmTitle, "'", "''") & "', " & _
                " '" & frmCategory & "', " & _
                " '" & frmSeverity & "', " & _
                " '" & Replace(frmDesc, "'", "''") & "', " & _
                " " & frmPoints & ", " & _
                " GETDATE(), " & _
                " 'Pending', " & _
                " " & (IIf(frmJobID="", "NULL", frmJobID)) & " " & _
                ");"
    
    ' 4. Execute SQL
    conn.Execute sqlInsert
    
    ' 5. Redirect to refresh the page and show new data
    Response.Redirect("problem_tracking.asp")
    Response.End
End If
%>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="windows-874">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Vinova v5.0 - Problem Tracking (ASP Classic)</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        /* Simple custom style to toggle the Modal */
        #addModal { display: none; }
        #addModal.open { display: flex; }
        
        /* Scrollbar for the Modal */
        .modal-content::-webkit-scrollbar { width: 6px; }
        .modal-content::-webkit-scrollbar-track { background: #f1f1f1; }
        .modal-content::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 3px; }
        .modal-content::-webkit-scrollbar-thumb:hover { background: #94a3b8; }
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
    ' DATA RETRIEVAL (EXISTING LOGIC)
    ' ==========================================
    
    Dim rsProblems, pSQL
    pSQL = "SELECT p.ProblemID, p.UserID, u.Name as UserName, p.Title, p.Category, p.Severity, p.Description, p.Points, p.RaisedDate, p.Status, p.JobID " & _
           "FROM [vn].[ProblemTracking] p " & _
           "INNER JOIN [vn].[Users] u ON p.UserID = u.UserID " & _
           "ORDER BY p.RaisedDate DESC"
           
    Set rsProblems = conn.Execute(pSQL)
    
    ' Process Data into a list
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
        
        ' Map Logic (As Requested)
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
    
    ' Counts
    Dim pTotal, pScheduled, pResolved, idx
    pTotal = UBound(problemsList) + 1
    
    Dim rsResolved
    Set rsResolved = conn.Execute("SELECT COUNT(*) as ResolvedCount FROM [vn].[ProblemTracking] WHERE Status = 'Resolved'")
    If Not rsResolved.EOF Then pResolved = rsResolved("ResolvedCount") Else pResolved = 0
    rsResolved.Close
    
    Dim rsScheduled
    Set rsScheduled = conn.Execute("SELECT COUNT(*) as ScheduledCount FROM [vn].[ProblemTracking] WHERE Status = 'Pending'")
    If Not rsScheduled.EOF Then pScheduled = rsScheduled("ScheduledCount") Else pScheduled = 0
    rsScheduled.Close
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
                        <div class="text-4xl font-bold text-red-600"><%= pTotal %></div>
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
                <div class="bg-white rounded-xl shadow-lg p-6 mb-6">
                    <div class="flex justify-between items-center mb-6">
                        <div>
                            <h2 class="text-2xl font-bold">Today's Meeting Agenda</h2>
                            <p class="text-sm text-gray-500"><%= Year(Now) %>/<%= Month(Now) %>/<%= Day(Now) %></p>
                        </div>
                        <button onclick="window.print()" class="px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition">
                            &#128196; Print Agenda
                        </button>
                    </div>
                    
                    <div class="space-y-6">
                        <!-- Render Logic Loop -->
                        <div>
                            <h3 class="font-semibold text-red-600 mb-3">&#9888; URGENT (Discuss First):</h3>
                            <% For idx = 0 To UBound(problemsList) 
                            If problemsList(idx)("severity") = "HIGH" Then 
                              Call RenderProblemCard(problemsList(idx), 1) 
                            End If 
                            Next %>
                        </div>
                        <div>
                            <h3 class="font-semibold text-orange-600 mb-3">&#128196; PERFORMANCE REVIEW:</h3>
                            <% For idx = 0 To UBound(problemsList) 
                            If problemsList(idx)("severity") = "MEDIUM" Then 
                              Call RenderProblemCard(problemsList(idx), 2) 
                            End If 
                            Next %>
                        </div>
                        <div>
                            <h3 class="font-semibold text-indigo-600 mb-3">&#128170; TRAINING NEEDS:</h3>
                            <% For idx = 0 To UBound(problemsList) 
                            If InStr(problemsList(idx)("category"), "Skills") > 0 Then 
                              Call RenderProblemCard(problemsList(idx), 4) 
                            End If 
                            Next %>
                        </div>
                        <div>
                            <h3 class="font-semibold text-yellow-600 mb-3">&#11088; RECOGNITION:</h3>
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
                    <!-- MODAL TRIGGER BUTTON -->
                    <button onclick="toggleModal('addModal', true)" class="px-6 py-3 bg-blue-600 text-white rounded-lg hover:bg-blue-700 font-medium shadow-md transition transform hover:-translate-y-0.5">
                        + Add New Problem
                    </button>
                    <button onclick="alert('In production: Shows analytics dashboard')" class="px-6 py-3 bg-gray-200 text-gray-700 rounded-lg hover:bg-gray-300 font-medium">
                        View Analytics
                    </button>
                </div>
            </div>
        </div>
    </div>

    <!-- ========================================== -->
    <!-- MODAL: ADD NEW PROBLEM                    -->
    <!-- ========================================== -->
    <div id="addModal" class="fixed inset-0 bg-black bg-opacity-50 z-50 items-center justify-center backdrop-blur-sm transition-opacity duration-300">
        <div class="bg-white rounded-2xl shadow-2xl w-full max-w-lg mx-4 overflow-hidden transform transition-all duration-300 scale-100">
            
            <!-- Modal Header -->
            <div class="bg-gradient-to-r from-blue-600 to-purple-600 px-6 py-4 flex justify-between items-center">
                <h3 class="text-xl font-bold text-white flex items-center gap-2">
                    <span>&#128196;</span> Add New Problem
                </h3>
                <button onclick="toggleModal('addModal', false)" class="text-white hover:text-gray-200 focus:outline-none text-2xl">&times;</button>
            </div>
            
            <!-- Modal Body (Form) -->
            <div class="p-6 max-h-[80vh] overflow-y-auto modal-content">
                <form action="problem_tracking.asp" method="POST">
                    <input type="hidden" name="action" value="add_problem">
                    
                    <!-- User ID (Hardcoded to 2 for demo, or use Session) -->
                    <input type="hidden" name="UserID" value="2">

                    <div class="space-y-4">
                        
                        <!-- Title Input -->
                        <div>
                            <label class="block text-sm font-medium text-gray-700 mb-1">Title</label>
                            <input type="text" name="Title" required class="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition" placeholder="e.g. Late Delivery">
                        </div>

                        <!-- Two Columns -->
                        <div class="grid grid-cols-2 gap-4">
                            <!-- Category Select -->
                            <div>
                                <label class="block text-sm font-medium text-gray-700 mb-1">Category</label>
                                <div class="relative">
                                    <select name="Category" class="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none appearance-none bg-white">
                                        <option value="Quality">Quality</option>
                                        <option value="Communication">Communication</option>
                                        <option value="Time Management">Time Management</option>
                                        <option value="Skills">Skills</option>
                                        <option value="Recognition">Recognition</option>
                                        <option value="Other">Other</option>
                                    </select>
                                    <div class="pointer-events-none absolute inset-y-0 right-0 flex items-center px-2 text-gray-700">
                                        <svg class="fill-current h-4 w-4" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20"><path d="M9.293 12.95l.707.707L15.657 8l-1.414-1.414L10 10.828 5.757 6.586 4.343 8z"/></svg>
                                    </div>
                                </div>
                            </div>

                            <!-- Severity Select -->
                            <div>
                                <label class="block text-sm font-medium text-gray-700 mb-1">Severity</label>
                                <div class="relative">
                                    <select name="Severity" class="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none appearance-none bg-white">
                                        <option value="LOW">Low</option>
                                        <option value="MEDIUM" selected>Medium</option>
                                        <option value="HIGH">High</option>
                                    </select>
                                    <div class="pointer-events-none absolute inset-y-0 right-0 flex items-center px-2 text-gray-700">
                                        <svg class="fill-current h-4 w-4" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20"><path d="M9.293 12.95l.707.707L15.657 8l-1.414-1.414L10 10.828 5.757 6.586 4.343 8z"/></svg>
                                    </div>
                                </div>
                            </div>
                        </div>

                        <!-- Description Textarea -->
                        <div>
                            <label class="block text-sm font-medium text-gray-700 mb-1">Description</label>
                            <textarea name="Description" rows="3" class="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition" placeholder="Provide details about the issue..."></textarea>
                        </div>

                        <!-- Points & JobID Row -->
                        <div class="grid grid-cols-2 gap-4">
                            <div>
                                <label class="block text-sm font-medium text-gray-700 mb-1">Points (Optional)</label>
                                <input type="number" name="Points" value="0" class="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none">
                            </div>
                            <div>
                                <label class="block text-sm font-medium text-gray-700 mb-1">Job ID (Optional)</label>
                                <input type="number" name="JobID" class="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none" placeholder="e.g. 101">
                            </div>
                        </div>

                    </div>
                    
                    <!-- Modal Footer -->
                    <div class="mt-8 flex justify-end gap-3">
                        <button type="button" onclick="toggleModal('addModal', false)" class="px-4 py-2 text-gray-600 bg-gray-100 rounded-lg hover:bg-gray-200 transition">Cancel</button>
                        <button type="submit" class="px-6 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 shadow-md transition">Save Problem</button>
                    </div>
                </form>
            </div>
        </div>
    </div>

    <script>
        // Simple function to toggle Modal visibility
        function toggleModal(modalID, show) {
            var modal = document.getElementById(modalID);
            if (show) {
                modal.classList.add('open');
            } else {
                modal.classList.remove('open');
            }
        }
    </script>

    <%
    ' ==========================================
    ' HELPER SUBROUTINE (RENDER CARD)
    ' ==========================================
    Sub RenderProblemCard(f_problem, f_index)
        Dim f_user, f_title, f_category, f_desc, f_color, f_severity, f_points, f_date, f_suggested, f_pointsClass
        f_user = f_problem("user")
        f_title = f_problem("title")
        f_category = f_problem("category")
        f_desc = f_problem("desc")
        f_color = f_problem("color")
        f_severity = f_problem("severity")
        f_points = f_problem("points")
        f_date = f_problem("date")
        
        If Not IsNull(f_date) Then f_date = Day(f_date) & "/" & Month(f_date) & "/" & Year(f_date) Else f_date = "N/A" End If
        
        f_suggested = ""
        If Not IsNull(f_points) And f_points <> 0 Then
            If f_points > 0 Then f_pointsClass = "text-green-600" : f_suggested = "+" & f_points & " pts" Else f_pointsClass = "text-red-600" : f_suggested = f_points & " pts" End If
        End If
    %>
        <div class="border-2 border-<%= f_color %>-200 bg-white rounded-lg p-4 mb-3 shadow-sm hover:shadow-md transition">
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
                        <% If f_suggested <> "" Then %><span class="font-semibold <%= f_pointsClass %>">Suggested: <%= f_suggested %></span><% End If %>
                    </div>
                </div>
                <div class="flex gap-2">
                    <button onclick="alert('In production: Opens discussion modal for <%= f_user %>')" class="px-3 py-1 bg-blue-600 text-white text-sm rounded hover:bg-blue-700 transition">Discuss</button>
                    <button class="px-3 py-1 bg-gray-200 text-gray-700 text-sm rounded hover:bg-gray-300 transition">Skip</button>
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