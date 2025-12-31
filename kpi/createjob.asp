<%@ Language=VBScript %>
<%
UserID=Session("UserID")
%>
<!DOCTYPE html>
<html lang="th-TH">
<head>
    <meta charset="windows-874">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Create Job - Project Management Dashboard</title>
    <link rel="stylesheet" href="css/style.css">
    <script>
        // คำนวณวันจันทร์ถัดไปสำหรับ Type 5
        function setNextMonday() {
            var targetDate = document.getElementById('targetDate');
            //alert(document.getElementById('jobType').value);
            if (document.getElementById('jobType').value == '5') {
                var today = new Date();
                var day = today.getDay();
                var diff = (day === 1) ? 7 : (8 - day); // ถ้าวันนี้จันทร์ -> จันทร์หน้า
                today.setDate(today.getDate() + diff);
                var dd = today.getDate();
                var mm = today.getMonth() + 1; // เดือนเริ่มจาก 0
                var yyyy = today.getFullYear();

                // เติมเลข 0 ข้างหน้า ถ้าน้อยกว่า 10
                if (dd < 10) dd = '0' + dd;
                if (mm < 10) mm = '0' + mm;

                // กำหนดค่าในรูปแบบ dd/mm/yyyy
                //targetDate.disabled = false;
                //alert(today);
                //targetDate.value = today.toISOString().split('T')[0];
                $("#targetDate").datepicker("option", "disabled", true);
                targetDate.disabled = false;
                targetDate.value = dd + '/' + mm + '/' + yyyy;

                targetDate.readOnly = true;

            } else {
                $("#targetDate").datepicker("option", "disabled", false);
                targetDate.disabled = false;

                //alert(document.getElementById('jobType').value);
            }
        }
    </script>
</head>
<body>
    <!--#include file="includes/include.asp" -->
    <!--#include file="includes/header.asp" -->
    <!--#include file="includes/navbar.asp" -->
<%
Function DateToISO(d)
    If IsDate(d) Then DateToISO = Year(d) & "-" & Right("0" & Month(d),2) & "-" & Right("0" & Day(d),2)
End Function
%>
    <link rel="stylesheet"
          href="https://cdn.jsdelivr.net/npm/flatpickr/dist/flatpickr.min.css">

    <!-- flatpickr JS -->
    <script src="https://cdn.jsdelivr.net/npm/flatpickr"></script>
<link rel="stylesheet" href="/js/jquery-ui.css">
 <script src="https://code.jquery.com/jquery-3.7.1.js"></script>
  <script src="/js/jquery-ui.js"></script>
<style>
ainput.input 
{
 awidth:50% !important;   
 aheight:30px;
 apadding-left:5px;
}
input::placeholder {
            color: #bbbbbb;        /* Light gray */
            font-style: italic;    /* Optional: make it italic */
            opacity: 1;            /* Ensures color is not faded in some browsers */
        }

        /* For better browser compatibility (older versions) */
        input::-webkit-input-placeholder { color: #bbbbbb; }
        input:-moz-placeholder { color: #bbbbbb; }           /* Firefox 18- */
        input::-moz-placeholder { color: #bbbbbb; }          /* Firefox 19+ */
        input:-ms-input-placeholder { color: #bbbbbb; }      /* IE 10+ */
</style>
    <div class="user-card" style="max-width: 800px; margin: 2rem auto;">
        <h2 style="font-size: 24px; margin-bottom: 1.5rem; text-align: center;">Create New Job</h2>

        <form method="post" action="process_createjob.asp">
            <div style="margin-bottom: 1.5rem;">
                <label style="display: block; margin-bottom: 0.5rem; font-weight: 600;">Job Title</label>
                <input type="text" name="title" placeholder="Enter job description..." required style="width: 100%; padding: 12px; border-radius: 8px; border: 1px solid #d1d5db; font-size: 16px;">
            </div>

            <div style="margin-bottom: 1.5rem; padding: 1rem; background: #fdf2f8; border-radius: 8px;">
                <label>
                    <input type="checkbox" name="isPrivate" value="1">
                    <strong style="color: #be123c;"> Assign to Myself (Private)</strong>
                </label>
                <div style="font-size: 14px; color: #6b7280; margin-top: 0.5rem;">
                    This job will be private and only visible to you. Management and others cannot see it. KPI penalties still apply to your score.
                </div>
            </div>

            <div style="margin-bottom: 1.5rem;">
                <label style="display: block; margin-bottom: 0.5rem; font-weight: 600;">Job Type</label>
                <select name="typeID" id="jobType" onchange="setNextMonday()" required style="width: 100%; padding: 12px; border-radius: 8px; border: 1px solid #d1d5db; font-size: 16px;">
                    <%
                    Dim rsTypes
                    Set rsTypes = conn.Execute("SELECT TypeID, TypeName, ShiftsAllowed, PenaltyPerDay, IsAutoMonday FROM JobTypes ORDER BY TypeID")
                    Do While Not rsTypes.EOF
                        Dim shiftsText
                        If IsNull(rsTypes("ShiftsAllowed")) Then
                            shiftsText = "Unlimited shifts"
                        Else
                            shiftsText = rsTypes("ShiftsAllowed") & " shifts"
                        End If
                        %>
                        <option value="<%= rsTypes("TypeID") %>">
                            Type <%= rsTypes("TypeID") %> (<%= rsTypes("TypeName") %>) | <%= shiftsText %> | <%= rsTypes("PenaltyPerDay") %> pts/day
                        </option>
                        <%
                        rsTypes.MoveNext
                    Loop
                    rsTypes.Close
                    %>
                </select>

                <div style="margin-top: 0.75rem; padding: 1rem; background: #eff6ff; border-radius: 8px; font-size: 14px;">
                    <strong>Rules:</strong> Type selected above<br>
                    &#2022; Shifts Allowed: <%= shiftsText %> (dynamic จาก selection)<br>
                    &#2022; Penalty: <%'=rsTypes("PenaltyPerDay") %> points per day late
                </div>
            </div>

            <div style="margin-bottom: 1.5rem;">
                <label style="display: block; margin-bottom: 0.5rem; font-weight: 600;">Assign To</label>
                <select name="assignedTo" id="assignedTo" onchange="updateLoadCheck(this.value);" required style="width: 90%; padding: 12px; border-radius: 8px; border: 1px solid #d1d5db; font-size: 16px;">
                    <option value="">Select team member...</option>
                    <%
                    Dim rsUsers
                    Set rsUsers = conn.Execute("SELECT UserID, Name, Department FROM Users ORDER BY Name")
                    Do While Not rsUsers.EOF
                        %>
                        <option value="<%= rsUsers("UserID") %>"><%= rsUsers("Name") %> - <%= rsUsers("Department") %></option>
                        <%
                        rsUsers.MoveNext
                    Loop
                    rsUsers.Close
                    %>
                </select>

<!-- Load Check - แสดงจำนวนงานค้างของผู้รับ (เหมือนไฟล์ตัวอย่าง) -->
    <div id="loadCheck" style="margin-top: 0.75rem; padding: 1rem; background: #fefce8; border-radius: 8px; border-left: 6px solid #f59e0b; font-weight: 600;">
        <div style="font-size: 16px; color: #92400e;">
            เลือกผู้รับงานเพื่อดูจำนวนงานค้าง
        </div>
    </div>

<script language=javascript>
    function updateLoadCheck(userID) {
        var loadDiv = document.getElementById('loadCheck');

        if (userID === '' || userID === null) {
            
            loadDiv.innerHTML = '<div style="font-size: 16px; color: #92400e;">เลือกผู้รับงานเพื่อดูจำนวนงานค้าง</div>';
            return;
        }

        // AJAX ดึงจำนวนงาน Pending (ไม่รวม Private ของคนอื่น)
        var xhr = new XMLHttpRequest();
        xhr.open("GET", "get_load_check.asp?userID=" + userID, true);
        xhr.onreadystatechange = function () {
            if (xhr.readyState == 4 && xhr.status == 200) {
                loadDiv.innerHTML = xhr.responseText;
            }
            else loadDiv.innerHTML = xhr.responseText;
        };
        xhr.send();
    }

    // เรียกตอนโหลดหน้า (ถ้ามีค่า default)
    window.onload = function () {
        var select = document.getElementById('assignedTo');
        if (select.value) {
            updateLoadCheck(select.value);
        }
    };

/*function updateLoadCheck() {
    var userID = document.getElementById('assignedTo').value;
    if (userID === '') {
        document.getElementById('loadCheck').innerHTML = 'เลือกผู้รับงานเพื่อดูจำนวนงานค้าง';
        return;
    }
    else {
        // AJAX ดึงจำนวนงาน Pending (ไม่รวม Private ของคนอื่น)
        var xhr = new XMLHttpRequest();
        xhr.open("GET", "get_load_check.asp?userID=" + userID, true);
        xhr.onreadystatechange = function () {
            if (xhr.readyState == 4 && xhr.status == 200) {
                document.getElementById('loadCheck').innerHTML = xhr.responseText;
            }
        };
        xhr.send();
    }
}

updateLoadCheck();
*/

</script>
            
            
            </div>

            <div style="margin-bottom: 2rem;position:relative;">
                <label style="display: block; margin-bottom: 0.5rem; font-weight: 600;">Target Date</label>
                <input type="text" name="targetDate" id="targetDate" value="<%'= DateToISO(Date()) %>" required style="height:36px;width: 100% !important; padding: 12px !important; border-radius: 8px !important; border: 1px solid #d1d5db !important; font-size: 16px !important;">
                <span id=cleardate style="position:absolute;top:50%;atransform: translateY(-50%);right:7px;cursor:pointer;" onclick="this.previousElementSibling.value='';">x</span>
            </div>

            <div style="text-align: center;">
                <button type="submit" style="background: #1e3a8a; color: white; padding: 14px 40px; border: none; border-radius: 8px; font-size: 18px; font-weight: 600; cursor: pointer;">
                    &#10004; Create Job
                </button>
            </div>
        </form>
    </div>
<script language=javascript>
    $("#targetDate").datepicker();
    $("#targetDatea").datepicker({
        beforeShow: function (input, inst) {
            setTimeout(function () {
                if (!$('#ui-datepicker-close-btn').length) {
                    var btn = $('<button>', {
                        text: 'Close',
                        id: 'ui-datepicker-close-btn',
                        click: function () {
                            $("#targetDate").datepicker("hide");
                        },
                        css: {
                            margin: '5px',
                            padding: '2px 8px',
                            cursor: 'pointer'
                        }
                    });
                    $(inst.dpDiv).prepend(btn);
                }
            }, 0);
        }
    });
    $("#targetDate").datepicker("option", "dateFormat", "d/m/yy");
    //const input = document.getElementById('targetDate');

/*input.addEventListener('search', (event) => {
  if (input.value === '') {
    // This was triggered by clicking the X clear button
    //console.log('Clear button (X) was clicked!');
    $("#targetDate").datepicker("hide");
    // Run your specific logic here (e.g., reset results, hide dropdown, etc.)
  } else {
    // This was triggered by pressing Enter (with text)
    //console.log('Search submitted:', input.value);
  }
});*/

</script>

    <!-- Switch User -->
    <div class="switch-user">
        Switch User (Demo): 
        <strong><%= Session("UserName") %> - <%= currentDepartment %></strong>
    </div>

    <% Call CloseConnection() %>
</body>
</html>