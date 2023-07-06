<!DOCTYPE html>
<html lang="en">
<head>
    <title>User Registration</title>
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">


    <style>
        body {
            padding: 20px;
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

<nav class="navbar navbar-light" style="background-color: #e3f2fd;">

        <a class="navbar-brand font-weight-bold" href="#">Student Management System</a>
        <div class="ml-auto">
            <button type="button" class="btn btn-info mr-2">Register</button>
            <button type="button" class="btn btn-success">Login</button>

    </div>
</nav>


    <div class="container">
        <h1>User Registration</h1>
        <form id="registrationForm" method="post" action="">
            <div class="form-group">
                <label for="firstName">First Name:</label>
                <input type="text" class="form-control" id="firstName" name="firstName" required>
            </div>
            <div class="form-group">
                <label for="lastName">Last Name:</label>
                <input type="text" class="form-control" id="lastName" name="lastName" required>
            </div>
            <div class="form-group">
                <label for="email">Email:</label>
                <input type="email" class="form-control" id="email" name="email" required>
                <small id="emailError" class="form-text text-danger"></small>
            </div>
            <div class="form-group">
                <label for="phone">Phone:</label>
                <input type="text" class="form-control" id="phone" name="phone" required>
            </div>
            <div class="form-group">
                <label for="password">Password:</label>
                <input type="password" class="form-control" id="password" name="password" required>
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
            <button type="submit" class="btn btn-success">Register</button><br>
            Already have an account?
            <button type="button" class="btn btn-danger" onclick="location.href='login.asp'">Login Here</button>
        </form>
    </div>

    <script src="js/bootstrap.min.js"></script>

    <%
    If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
        ' Retrieve form data
        Dim firstName, lastName, email, phone, password, gender
        firstName = Request.Form("firstName")
        lastName = Request.Form("lastName")
        email = Request.Form("email")
        phone = Request.Form("phone")
        password = Request.Form("password")
        gender = Request.Form("gender")

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

        ' Prepare the SQL statement to insert data into the database
        Dim strSQLInsert
        strSQLInsert = "INSERT INTO [users] ([FirstName], [LastName], [Email], [Phone], [Password], [Gender]) VALUES ('" & Replace(firstName, "'", "''") & "', '" & Replace(lastName, "'", "''") & "', '" & Replace(email, "'", "''") & "', '" & Replace(phone, "'", "''") & "', '" & Replace(password, "'", "''") & "', '" & Replace(gender, "'", "''") & "')"

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

        ' Redirect to the login page
        Response.Redirect "login.asp"
    End If
    %>
</body>
</html>
