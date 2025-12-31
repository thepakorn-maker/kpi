<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="includes/include.asp" -->
<!--#include file="db/connect.asp" -->

<%
' ==========================================
' HANDLE FORM SUBMISSION (INSERT)
' ==========================================
'Dim conn, connString, action, sqlCmd
'connString = "Provider=SQLNCLI11;Server=YOUR_SERVER_NAME;Database=vndb;Uid=YOUR_USERNAME;Pwd=YOUR_PASSWORD;"
'Set conn = Server.CreateObject("ADODB.Connection")
'conn.Open connString

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim frmUserID, frmTitle, frmCategory, frmSeverity, frmDesc, frmPoints, frmJobID
    
    frmUserID   = 2 ' Hardcoded for demo. Use Session("UserID") in production.
    frmTitle    = Request.Form("Title")
    frmCategory = Request.Form("Category")
    frmSeverity = Request.Form("Severity")
    frmDesc     = Request.Form("Description")
    frmPoints   = Request.Form("Points")
    frmJobID    = Request.Form("JobID")
    
    If frmTitle = "" Then
        Response.Write("Error: Title is required.")
        Response.End
    End If

    sqlCmd = "INSERT INTO [vn].[ProblemTracking] ([UserID], [Title], [Category], [Severity], [Description], [Points], [RaisedDate], [Status], [JobID]) VALUES (" & _
             frmUserID & ", '" & Replace(frmTitle, "'", "''") & "', '" & frmCategory & "', '" & frmSeverity & "', '" & Replace(frmDesc, "'", "''") & "', " & frmPoints & ", GETDATE(), 'Pending', " & (IIf(frmJobID="", "NULL", frmJobID)) & ");"
    
    conn.Execute sqlCmd
    conn.Close
    Set conn = Nothing
    
    ' Redirect back to list
    Response.Redirect "problem_list.asp"
    Response.End
End If
%>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="windows-874">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Add Problem</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-gray-100 font-sans text-gray-800 p-4 md:p-8">

    <div class="max-w-2xl mx-auto bg-white rounded-xl shadow-lg overflow-hidden">
        <div class="bg-gradient-to-r from-blue-600 to-purple-600 px-6 py-4">
            <h1 class="text-xl font-bold text-white">&#128196; Add New Problem</h1>
            <p class="text-blue-100 text-sm">Create a new tracking record</p>
        </div>
        
        <div class="p-6">
            <form method="POST" action="problem_add.asp" class="space-y-6">
                
                <div>
                    <label class="block text-sm font-medium text-gray-700 mb-1">Title</label>
                    <input type="text" name="Title" required class="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none">
                </div>

                <div class="grid md:grid-cols-2 gap-6">
                    <div>
                        <label class="block text-sm font-medium text-gray-700 mb-1">Category</label>
                        <select name="Category" class="w-full px-4 py-2 border border-gray-300 rounded-lg bg-white">
                            <option value="Quality">Quality</option>
                            <option value="Communication">Communication</option>
                            <option value="Time Management">Time Management</option>
                            <option value="Skills">Skills</option>
                            <option value="Recognition">Recognition</option>
                            <option value="Other">Other</option>
                        </select>
                    </div>
                    <div>
                        <label class="block text-sm font-medium text-gray-700 mb-1">Severity</label>
                        <select name="Severity" class="w-full px-4 py-2 border border-gray-300 rounded-lg bg-white">
                            <option value="LOW">Low</option>
                            <option value="MEDIUM" selected>Medium</option>
                            <option value="HIGH">High</option>
                        </select>
                    </div>
                </div>

                <div>
                    <label class="block text-sm font-medium text-gray-700 mb-1">Description</label>
                    <textarea name="Description" rows="4" class="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none"></textarea>
                </div>

                <div class="grid md:grid-cols-2 gap-6">
                    <div>
                        <label class="block text-sm font-medium text-gray-700 mb-1">Points</label>
                        <input type="number" name="Points" value="0" class="w-full px-4 py-2 border border-gray-300 rounded-lg">
                    </div>
                    <div>
                        <label class="block text-sm font-medium text-gray-700 mb-1">Job ID</label>
                        <input type="number" name="JobID" class="w-full px-4 py-2 border border-gray-300 rounded-lg">
                    </div>
                </div>

                <div class="flex items-center gap-4 pt-4">
                    <button type="submit" class="flex-1 bg-blue-600 text-white py-3 rounded-lg hover:bg-blue-700 font-medium">Save Problem</button>
                    <a href="view_problemtracking.asp" class="flex-1 text-center bg-gray-200 text-gray-700 py-3 rounded-lg hover:bg-gray-300 font-medium">Cancel</a>
                </div>
            </form>
        </div>
    </div>

</body>
</html>