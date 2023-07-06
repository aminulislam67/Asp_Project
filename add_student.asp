<!DOCTYPE html>
<html lang="en">
<head>
    <title>Add Student</title>
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
    <style>
        body {
            padding: 20px;
            background-color: #e4f4fc;
        }
        .container {
            max-width: 500px;
            margin: 0 auto;
            padding: 20px;
            border: 3px solid #ccc;
            border-radius: 8px;
            background-color: #f7f7f7;
            box-shadow: 0px 5px 10px skyblue;
            margin-top: 20px;
        }
    </style>
</head>
<body>
<nav class="navbar navbar-light" style="background-color: #27292b;">
    <a class="navbar-brand font-weight-bold text-white" href="#">Student Management System</a>
    <div class="ml-auto">
        <button type="button" class="btn btn-info mr-2">Add Student</button>
        <button type="button" class="btn btn-success">Display</button>
    </div>
</nav>

<div class="container">
    <h2>Add Student</h2>
    <form id="addStudentForm" method="post" action="">
        <div class="form-group">
            <label for="firstName">First Name:</label>
            <input type="text" class="form-control" id="firstName" name="firstName" required>
        </div>
        <div class="form-group">
            <label for="lastName">Last Name:</label>
            <input type="text" class="form-control" id="lastName" name="lastName" required>
        </div>
        <div class="form-group">
            <label for="studentID">Student ID:</label>
            <input type="text" class="form-control" id="studentID" name="studentID" required>
        </div>
        <div class="form-group">
            <label for="email">Email:</label>
            <input type="email" class="form-control" id="email" name="email" required>
        </div>
        <div class="form-group">
            <label for="gender">Gender:</label>
            <select class="form-control" id="gender" name="gender" required>
                <option value="">Select Gender</option>
                <option value="male">Male</option>
                <option value="female">Female</option>
                <option value="other">Other</option>
            </select>
        </div>
        <div class="form-group">
            <label for="session">Session:</label>
            <input type="text" class="form-control" id="session" name="session" required>
        </div>
        <div class="form-group">
            <label for="dateOfBirth">Date of Birth:</label>
            <input type="date" class="form-control" id="dateOfBirth" name="dateOfBirth" required>
        </div>
        <button type="submit" class="btn btn-primary">Submit</button>
    </form>
</div>

<script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>

<%
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' Retrieve form data
    Dim firstName, lastName, studentID, email, gender, session, dateOfBirth
    firstName = Request.Form("firstName")
    lastName = Request.Form("lastName")
    studentID = Request.Form("studentID")
    email = Request.Form("email")
    gender = Request.Form("gender")
    session = Request.Form("session")
    dateOfBirth = Request.Form("dateOfBirth")

    ' Path to the Access database file
    Dim dbPath
    dbPath = Server.MapPath("crud_db.accdb")

    ' Connection string for Access database
    Dim connStr
    connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"

    ' Create a new connection object
    Dim conn
    Set conn = Server.CreateObject("ADODB.Connection")

    On Error Resume Next

    ' Open the database connection
    conn.Open connStr

    ' Check for connection errors
    If Err.Number <> 0 Then
        Response.Write "An error occurred while connecting to the database."
        Response.End
    End If

    On Error Goto 0

    ' Check if the student ID already exists in the database
    Dim strSQLCheck
    strSQLCheck = "SELECT COUNT(*) FROM [Students] WHERE [StudentID] = '" & Replace(studentID, "'", "''") & "'"
    Dim rsCheck
    Set rsCheck = conn.Execute(strSQLCheck)
    Dim studentCount
    studentCount = rsCheck.Fields(0).Value
    rsCheck.Close

    If studentCount > 0 Then
        Response.Write "The student ID is already registered."
    Else
        ' Prepare the SQL statement to insert data into the database
        Dim strSQLInsert
        strSQLInsert = "INSERT INTO [Students] ([FirstName], [LastName], [StudentID], [Email], [Gender], [Session], [DateOfBirth]) VALUES ('" & Replace(firstName, "'", "''") & "', '" & Replace(lastName, "'", "''") & "', '" & Replace(studentID, "'", "''") & "', '" & Replace(email, "'", "''") & "', '" & Replace(gender, "'", "''") & "', '" & Replace(session, "'", "''") & "', #" & Replace(dateOfBirth, "'", "''") & "#)"

        ' Execute the SQL insert statement
        conn.Execute strSQLInsert

        ' Check for any errors during the insert operation
        If Err.Number <> 0 Then
            Response.Write "An error occurred while saving the form data."
            Response.End
        End If

        ' Close the database connection
        conn.Close
        Set conn = Nothing

        ' Redirect to a success page or display a success message
        Response.Redirect "display.asp"
    End If
End If
%>
</body>
</html>
