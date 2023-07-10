<!DOCTYPE html>
<html lang="en">
<head>
    <title>User Login</title>
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
    <style>
        body {
            padding: 20px;
        }
        .container {
            max-width: 500px;
            margin: 0 auto;
            margin-top: 150px;
            padding: 20px;
            border: 3px solid #ccc;
            border-radius: 8px;
            background-color: #f7f7f7;
            box-shadow: 0px 5px 10px skyblue;
        }
    </style>
</head>
<body>

<nav class="navbar navbar-light" style="background-color: #e3f2fd;">
    <a class="navbar-brand font-weight-bold" href="#">Student Management System</a>
    <div class="ml-auto">
        <button type="button" class="btn btn-info mr-2">Register</button>
        <button type="button" class="btn btn-success">Login</button>
    </div>
</nav>

<div class="container">
    <h1>User Login</h1>
    <form id="loginForm" method="post" action="login.asp">
        <div class="form-group">
            <label for="email">Email:</label>
            <input type="email" class="form-control" id="email" name="email" required pattern="[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}" title="Please provide a valid email address."/>
        </div>
        <div class="form-group">
            <label for="password">Password:</label>
            <input type="password" class="form-control" id="password" name="password" required minlength="8" title="Password must have a minimum of 8 characters."/>
        </div>
        <button type="submit" class="btn btn-danger">Login</button><br>
        Don't have an account?
        <button type="button" class="btn btn-success" onclick="location.href='process_registration.asp'">Register</button>
    </form>
</div>

<script src="js/bootstrap.min.js"></script>


<%
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' Retrieve form data
    Dim email, password
    email = Request.Form("email")
    password = Request.Form("password")

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

    ' Prepare the SQL statement to check the login credentials
    Dim strSQLLogin
    strSQLLogin = "SELECT Email, Password FROM [users] WHERE [Email] = '" & Replace(email, "'", "''") & "' AND [Password] = '" & Replace(password, "'", "''") & "'"

    ' Execute the SQL select statement
    Dim rsLogin
    Set rsLogin = conn.Execute(strSQLLogin)

    ' Check if login is successful
    If Not rsLogin.EOF Then
        ' Login successful
        Response.Redirect "display.asp"
        
        Set rsLogin = Nothing
        conn.Close
        Set conn = Nothing
        
    Else
        ' Login failed
        Dim strSQLEmailCheck, strSQLPasswordCheck
        strSQLEmailCheck = "SELECT Email FROM [users] WHERE [Email] = '" & Replace(email, "'", "''") & "'"
        strSQLPasswordCheck = "SELECT Password FROM [users] WHERE [Password] = '" & Replace(password, "'", "''") & "'"

        Dim rsEmailCheck, rsPasswordCheck
        Set rsEmailCheck = conn.Execute(strSQLEmailCheck)
        Set rsPasswordCheck = conn.Execute(strSQLPasswordCheck)

        Dim emailCorrect, passwordCorrect
        emailCorrect = Not rsEmailCheck.EOF
        passwordCorrect = Not rsPasswordCheck.EOF

        If Not emailCorrect And Not passwordCorrect Then
            ' Both email and password are incorrect
            Response.Write "<h2>Email and password are incorrect.</h2>"
        ElseIf Not emailCorrect Then
            ' Email is incorrect
            Response.Write "<h2>Email is incorrect.</h2>"
        ElseIf Not passwordCorrect Then
            ' Password is incorrect
            Response.Write "<h2>Password is incorrect.</h2>"
        End If
    End If

    ' Close the database connection
    conn.Close
    Set conn = Nothing
End If
%>
</body>
</html>
