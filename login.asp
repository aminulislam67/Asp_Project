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
    <div class="container">
        <h1>User Login</h1>
        <form id="loginForm" method="post" action="login.asp">
            <div class="form-group">
                <label for="email">Email:</label>
                <input type="email" class="form-control" id="email" name="email" required>
            </div>
            <div class="form-group">
                <label for="password">Password:</label>
                <input type="password" class="form-control" id="password" name="password" required>
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
        strSQLLogin = "SELECT * FROM [users] WHERE [Email] = '" & Replace(email, "'", "''") & "' AND [Password] = '" & Replace(password, "'", "''") & "'"

        ' Execute the SQL select statement
        Dim rsLogin
        Set rsLogin = conn.Execute(strSQLLogin)

        ' Check if login is successful
        If Not rsLogin.EOF Then
            ' Login successful
            Response.Write "<h2>Login Successful</h2>"
        Else
            ' Login failed
            Response.Write "<h2>Login Failed</h2>"
        End If

        ' Close the database connection
        conn.Close
        Set conn = Nothing
    End If
    %>
</body>
</html>
