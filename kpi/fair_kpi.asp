<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
' Option Explicit requires all variables to be declared.
' We handle variable declarations carefully to avoid "Name Redefined" errors.
%>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="windows-874">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Vinova v5.0 - New Features Demo (ASP Classic)</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-gray-100">
<!--#include file="includes/include.asp" -->
<!--#include file="db/connect.asp" -->

    <%
    ' ==========================================
    ' DATABASE CONFIGURATION
    ' ==========================================
    'Dim conn, connString
    ' Update this connection string to match your SQL Server environment
    'connString = "Provider=SQLNCLI11;Server=YOUR_SERVER_NAME;Database=vndb;Uid=YOUR_USERNAME;Pwd=YOUR_PASSWORD;"
    
    'Set conn = Server.CreateObject("ADODB.Connection")
    'conn.Open connString

    ' ==========================================
    ' HELPER FUNCTIONS
    ' ==========================================

    Function GetTier(pct)
        ' Variables used inside functions are local to that function.
        ' They do not conflict with variables outside.
        Dim tierName, tierIcon, tierColor
        If pct >= 95 Then
            tierName = "PLATINUM"
            tierIcon = "&#127942;"
            tierColor = "bg-yellow-500"
        ElseIf pct >= 85 Then
            tierName = "GOLD"
            tierIcon = "&#129351;"
            tierColor = "bg-yellow-400"
        ElseIf pct >= 70 Then
            tierName = "SILVER"
            tierIcon = "&#129352;"
            tierColor = "bg-gray-400"
        ElseIf pct >= 60 Then
            tierName = "BRONZE"
            tierIcon = "&#129353;"
            tierColor = "bg-orange-400"
        Else
            tierName = "NEEDS IMPROVEMENT"
            tierIcon = "&#9888;"
            tierColor = "bg-red-500"
        End If
        ' Return an array (simulating the previous structure)
        GetTier = Array(tierName, tierIcon, tierColor)
    End Function

    Function GetStars(pct)
        Dim s, count, i, output
        s = "&#11088;"
        output = ""
        count = 0
        
        If pct >= 95 Then
            count = 5
        ElseIf pct >= 85 Then
            count = 4
        ElseIf pct >= 70 Then
            count = 3
        ElseIf pct >= 60 Then
            count = 2
        Else
            count = 1
        End If
        
        For i = 1 To count
            output = output & s
        Next
        
        GetStars = output
    End Function

    ' ==========================================
    ' DATA RETRIEVAL
    ' ==========================================

    ' 1. Fetch Team Data (Aggregating stats from Jobs table)
    Dim rsTeam, teamSQL
    teamSQL = "SELECT " & _
              "  u.UserID, u.Name, u.Department, " & _
              "  COUNT(j.JobID) as TotalJobs, " & _
              "  SUM(CASE WHEN j.Status = 'Completed' AND j.CompletedDate <= j.TargetDate THEN 1 ELSE 0 END) as OnTimeJobs, " & _
              "  SUM(CASE WHEN j.Status = 'Completed' AND j.CompletedDate > j.TargetDate THEN 1 ELSE 0 END) as LateJobs, " & _
              "  SUM(CASE WHEN j.Status = 'Pending' AND j.TargetDate < GETDATE() THEN 1 ELSE 0 END) as OverdueCount " & _
              "FROM [vn].[Users] u " & _
              "LEFT JOIN [vn].[Jobs] j ON u.UserID = j.AssignedTo " & _
              "GROUP BY u.UserID, u.Name, u.Department " & _
              "ORDER BY OnTimeJobs DESC"
teamSQL = "SELECT " & _
              "  u.UserID, u.Name, u.Department, " & _
              "  COUNT(j.JobID) as TotalJobs, " & _
              "  SUM(CASE " & _
              "      WHEN j.Status = 'Completed' AND j.CompletedDate <= j.TargetDate THEN 1 " & _
              "      WHEN j.Status = 'Pending' AND j.TargetDate >= CAST(GETDATE() AS DATE) THEN 1 " & _
              "      ELSE 0 " & _
              "  END) as OnTimeJobs, " & _
              "  SUM(CASE " & _
              "      WHEN j.Status = 'Completed' AND j.CompletedDate > j.TargetDate THEN 1 " & _
              "      WHEN j.Status = 'Pending' AND j.TargetDate < CAST(GETDATE() AS DATE) THEN 1 " & _
              "      ELSE 0 " & _
              "  END) as LateJobs, " & _
              "  SUM(CASE WHEN j.Status = 'Pending' AND j.TargetDate < CAST(GETDATE() AS DATE) THEN 1 ELSE 0 END) as OverdueCount " & _
              "FROM [vn].[Users] u " & _
              "LEFT JOIN [vn].[Jobs] j ON u.UserID = j.AssignedTo " & _
              "GROUP BY u.UserID, u.Name, u.Department " & _
              "ORDER BY OnTimeJobs DESC"

    Set rsTeam = conn.Execute(teamSQL)

    ' 2. Process Data into an Array
    Dim teamList
    teamList = Array()
    Dim totalPctSum, memberCount
    totalPctSum = 0
    memberCount = 0
        Dim uID, uName, uDept, tJobs, oTime, lJobs, ovDue, pct, streak

    Do While Not rsTeam.EOF
        
        uID = rsTeam("UserID")
        uName = rsTeam("Name")
        uDept = rsTeam("Department")
        tJobs = rsTeam("TotalJobs")
        
        If IsNull(tJobs) Or tJobs = 0 Then tJobs = 0 Else tJobs = CInt(tJobs)
        If IsNull(rsTeam("OnTimeJobs")) Then oTime = 0 Else oTime = CInt(rsTeam("OnTimeJobs"))
        If IsNull(rsTeam("LateJobs")) Then lJobs = 0 Else lJobs = CInt(rsTeam("LateJobs"))
        If IsNull(rsTeam("OverdueCount")) Then ovDue = 0 Else ovDue = CInt(rsTeam("OverdueCount"))

        ' Calculate Percentage (Handle division by zero)
        If tJobs > 0 Then
            pct = Round((oTime / tJobs) * 100)
        Else
            pct = 0 ' No jobs yet
        End If

        ' Simulate Streak
        If ovDue = 0 And tJobs > 2 And pct >= 70 Then
            streak = 3 
        ElseIf ovDue = 0 Then
            streak = 0
        Else
            streak = 0
        End If

        ' Add to List
        ReDim Preserve teamList(UBound(teamList) + 1)
        Set teamList(UBound(teamList)) = Server.CreateObject("Scripting.Dictionary")
        teamList(UBound(teamList)).Add "name", uName
        teamList(UBound(teamList)).Add "dept", uDept
        teamList(UBound(teamList)).Add "pct", pct
        teamList(UBound(teamList)).Add "jobs", tJobs
        teamList(UBound(teamList)).Add "onTime", oTime
        teamList(UBound(teamList)).Add "late", lJobs
        teamList(UBound(teamList)).Add "streak", streak

        totalPctSum = totalPctSum + pct
        memberCount = memberCount + 1

        rsTeam.MoveNext
    Loop
    rsTeam.Close

    Dim teamAvg
    If memberCount > 0 Then
        teamAvg = Round(totalPctSum / memberCount)
    Else
        teamAvg = 0
    End If

    ' Sort the list by Percentage Descending
    Dim i, j, temp
    For i = 0 To UBound(teamList) - 1
        For j = i + 1 To UBound(teamList)
            If teamList(i)("pct") < teamList(j)("pct") Then
                Set temp = teamList(i)
                Set teamList(i) = teamList(j)
                Set teamList(j) = temp
            End If
        Next
    Next

    ' Set Current User (Simulating login)
    Dim currentUser
    Dim targetIndex
    If UBound(teamList) >= 2 Then targetIndex = 8 Else targetIndex = 0
    
    Set currentUser = teamList(targetIndex)
    
    ' ==========================================
    ' HTML RENDERING
    ' ==========================================
    %>
    
    <div id="app">
        <!-- Header -->
        <div class="bg-gradient-to-r from-blue-600 to-purple-600 text-white p-6 shadow-lg">
            <div class="max-w-6xl mx-auto">
                <h1 class="text-3xl font-bold">Vinova v5.0 - New Features Demo</h1>
                <p class="mt-2 text-blue-100">ASP Classic Integration - Fair KPI System</p>
            </div>
        </div>
        
        <!-- Navigation -->
        <div class="max-w-6xl mx-auto p-4">
            <div class="bg-white rounded-lg shadow-lg p-2 flex gap-2 mb-6">
                <button disabled class="flex-1 py-3 px-4 rounded-lg font-medium bg-blue-600 text-white cursor-default opacity-100">
                    &#128202; NEW: Fair KPI System
                </button>
                <button onclick="alert('Problem Tracking module would go here.')" class="flex-1 py-3 px-4 rounded-lg font-medium bg-gray-100 text-gray-700 hover:bg-gray-200">
                    &#128196; NEW: Problem Tracking
                </button>
            </div>
            
            <div id="content">
                <%
                ' ==========================================
                ' KPI RENDERING LOGIC
                ' ==========================================
                ' NOTE: We do NOT use 'Dim' here. The variables are implicitly available 
                ' because we are not using Option Explicit, or if we add it, these would 
                ' need to be declared at the very top. By removing Dim here, we fix the
               ' "Name redefined" error caused by Function/Scope overlap in previous versions.
                
                'Dim uTier, uStars, uVsAvg, uVsTop, uPct, uName, uDept, uJobs, uOnTime, uLate, uStreak
                
                uName = currentUser("name")
                uDept = currentUser("dept")
                uPct = currentUser("pct")
                uJobs = currentUser("jobs")
                uOnTime = currentUser("onTime")
                uLate = currentUser("late")
                uStreak = currentUser("streak")
                
                uTier = GetTier(uPct)
                uStars = GetStars(uPct)
                
                uVsAvg = uPct - teamAvg
                uVsTop = uPct - teamList(0)("pct")
                %>
                
                <!-- Team Member View -->
                <div class="bg-white rounded-xl shadow-lg p-6 mb-6">
                    <div class="flex justify-between items-start mb-6">
                        <div>
                            <h2 class="text-2xl font-bold">My Performance Dashboard</h2>
                            <p class="text-gray-600">Viewing as: <strong><%= uName %></strong> (<%= uDept %>)</p>
                        </div>
                        <div class="text-5xl"><%= uTier(1) %></div>
                    </div>
                    
                    <div class="grid md:grid-cols-2 gap-6">
                        <!-- Left: Main Stats -->
                        <div>
                            <div class="bg-gradient-to-br from-blue-50 to-blue-100 rounded-xl p-6">
                                <div class="flex justify-between mb-2">
                                    <h3 class="text-lg font-semibold">On-Time Rate</h3>
                                    <span class="text-xl"><%= uStars %></span>
                                </div>
                                <div class="text-5xl font-bold text-blue-600 mb-2"><%= uPct %>%</div>
                                <div class="mb-3">
                                    <span class="text-sm">Performance Tier: </span>
                                    <span class="font-bold"><%= uTier(1) %> <%= uTier(0) %></span>
                                </div>
                                <div class="w-full bg-gray-200 rounded-full h-5 mb-2">
                                    <div class="<%= uTier(2) %> h-5 rounded-full" style="width: <%= uPct %>%"></div>
                                </div>
                                <div class="text-xs text-gray-600">Bonus eligible: 5%</div>
                            </div>
                            
                            <div class="bg-gray-50 rounded-xl p-6 mt-4">
                                <h4 class="font-semibold mb-4">Detailed Stats</h4>
                                <div class="space-y-2">
                                    <div class="flex justify-between">
                                        <span>Total Jobs:</span>
                                        <span class="font-bold"><%= uJobs %></span>
                                    </div>
                                    <div class="flex justify-between">
                                        <span>&#8226; On-Time:</span>
                                        <span class="font-bold text-green-600"><%= uOnTime %> &#10004;</span>
                                    </div>
                                    <div class="flex justify-between">
                                        <span>&#8226; Late:</span>
                                        <span class="font-bold text-orange-600"><%= uLate %> &#9888;</span>
                                    </div>
                                    <div class="border-t pt-2 flex justify-between">
                                        <span>Current Streak:</span>
                                        <span class="font-bold text-blue-600"><%= uStreak %> jobs &#128293;</span>
                                    </div>
                                </div>
                            </div>
                        </div>
                        
                        <!-- Right: Comparison -->
                        <div>
                            <div class="bg-gradient-to-br from-purple-50 to-purple-100 rounded-xl p-6">
                                <h4 class="font-semibold mb-4">Team Comparison</h4>
                                <div class="space-y-4">
                                    <div>
                                        <div class="flex justify-between mb-1">
                                            <span class="text-sm">You</span>
                                            <span class="text-sm font-bold"><%= uPct %>%</span>
                                        </div>
                                        <div class="w-full bg-gray-200 rounded-full h-2">
                                            <div class="bg-blue-600 h-2 rounded-full" style="width: <%= uPct %>%"></div>
                                        </div>
                                    </div>
                                    <div>
                                        <div class="flex justify-between mb-1">
                                            <span class="text-sm">Team Average</span>
                                            <span class="text-sm font-bold"><%= teamAvg %>%</span>
                                        </div>
                                        <div class="w-full bg-gray-200 rounded-full h-2">
                                            <div class="bg-gray-600 h-2 rounded-full" style="width: <%= teamAvg %>%"></div>
                                        </div>
                                    </div>
                                    <div>
                                        <div class="flex justify-between mb-1">
                                            <span class="text-sm">Top Performer</span>
                                            <span class="text-sm font-bold"><%= teamList(0)("pct") %>%</span>
                                        </div>
                                        <div class="w-full bg-gray-200 rounded-full h-2">
                                            <div class="bg-yellow-500 h-2 rounded-full" style="width: <%= teamList(0)("pct") %>%"></div>
                                        </div>
                                    </div>
                                </div>
                                
                                <div class="mt-4 pt-4 border-t border-purple-200 space-y-2">
                                    <div class="flex justify-between">
                                        <span class="font-medium">Status:</span>
                                        <span class="font-bold <%= IIf(uVsAvg >= 0, "text-green-600", "text-orange-600") %>">
                                            <%= IIf(uVsAvg > 0, "+", "") & uVsAvg %>% vs average <%= IIf(uVsAvg >= 0, "&#8593;", "&#8595;") %>
                                        </span>
                                    </div>
                                    <div class="flex justify-between">
                                        <span class="font-medium">Gap to top:</span>
                                        <span class="font-bold text-purple-600"><%= Abs(uVsTop) %>%</span>
                                    </div>
                                </div>
                            </div>
                            
                            <div class="bg-blue-50 border-l-4 border-blue-500 p-4 rounded mt-4">
                                <p class="text-sm text-blue-800">
                                    <strong>Good job!</strong> You're on track. Focus on improving to reach Gold tier! &#128170;
                                </p>
                            </div>
                        </div>
                    </div>
                </div>
                
                <!-- Manager View - Full Leaderboard -->
                <div class="bg-white rounded-xl shadow-lg p-6">
                    <div class="flex justify-between items-center mb-6">
                        <div>
                            <h2 class="text-2xl font-bold">Team Performance Leaderboard</h2>
                            <p class="text-sm text-gray-500">&#128065; Managers Only - Full Rankings</p>
                        </div>
                        <div class="text-right">
                            <div class="text-sm text-gray-600">Team Average</div>
                            <div class="text-3xl font-bold text-blue-600"><%= teamAvg %>%</div>
                        </div>
                    </div>
                    
                    <div class="space-y-3">
                        <%
                        Dim mIdx, mPct, mOnTime, mTotalJobs, mName, mDept, mStreak, mTierArr, mMedal, mVsAvg
                        
                        For mIdx = 0 To UBound(teamList)
                            mName = teamList(mIdx)("name")
                            mDept = teamList(mIdx)("dept")
                            mPct = teamList(mIdx)("pct")
                            mOnTime = teamList(mIdx)("onTime")
                            mTotalJobs = teamList(mIdx)("jobs")
                            mStreak = teamList(mIdx)("streak")
                            
                            mTierArr = GetTier(mPct)
                            mVsAvg = mPct - teamAvg
                            
                            ' Medal logic
                            If mIdx = 0 Then
                                mMedal = "&#129351;"
                            ElseIf mIdx = 1 Then
                                mMedal = "&#129352;"
                            ElseIf mIdx = 2 Then
                                mMedal = "&#129353;"
                            Else
                                mMedal = ""
                            End If
                        %>
                            <div class="p-4 rounded-lg border-2 <%= IIf(mIdx < 3, "border-yellow-300 bg-yellow-50", "border-gray-200 bg-gray-50") %>">
                                <div class="flex items-center justify-between">
                                    <div class="flex items-center gap-3 flex-1">
                                        <div class="text-xl font-bold text-gray-400 w-8"><%= mIdx + 1 %>.</div>
                                        <% If mMedal <> "" Then %>
                                            <div class="text-2xl"><%= mMedal %></div>
                                        <% End If %>
                                        <div>
                                            <div class="font-bold text-lg"><%= mName %></div>
                                            <div class="text-sm text-gray-600"><%= mDept %></div>
                                        </div>
                                    </div>
                                    
                                    <div class="flex items-center gap-4">
                                        <div class="text-right">
                                            <div class="text-2xl font-bold text-blue-600"><%= mPct %>%</div>
                                            <div class="text-xs text-gray-500">(<%= mOnTime %>/<%= mTotalJobs %>)</div>
                                        </div>
                                        <div class="text-3xl"><%= mTierArr(1) %></div>
                                    </div>
                                </div>
                                <div class="mt-2 flex justify-end gap-4 text-xs">
                                    <span class="text-gray-600">Streak: <strong><%= mStreak %></strong></span>
                                    <span class="font-bold <%= IIf(mVsAvg >= 0, "text-green-600", "text-red-600") %>">
                                        <%= IIf(mVsAvg > 0, "+", "") & mVsAvg %>% vs avg
                                    </span>
                                </div>
                            </div>
                        <%
                        Next
                        %>
                    </div>
                    
                    <div class="mt-6 p-4 bg-purple-50 rounded-lg">
                        <h3 class="font-semibold text-purple-900 mb-3">&#128161; Team Insights</h3>
                        <div class="grid md:grid-cols-2 gap-4 text-sm">
                            <div>
                                <div class="font-semibold text-green-700 mb-1">&#127942; Excellent:</div>
                                <% 
                                For k = 0 To 1
                                    If k <= UBound(teamList) Then
                                %>
                                        <div>&#8226; <%= teamList(k)("name") %> (<%= teamList(k)("pct") %>%)</div>
                                <%
                                    End If
                                Next
                                %>
                            </div>
                            <div>
                                <div class="font-semibold text-orange-700 mb-1">&#9888; Needs Support:</div>
                                <%
                                Dim lowIdx
                                lowIdx = UBound(teamList)
                                %>
                                <div>&#8226; <%= teamList(lowIdx)("name") %> (<%= teamList(lowIdx)("pct") %>%) - Schedule 1-on-1</div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <%
    ' Cleanup
    conn.Close
    Set conn = Nothing
    %>
</body>
</html>