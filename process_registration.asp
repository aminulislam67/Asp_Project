<!DOCTYPE html>
<html lang="en">
<head>
    <title>User Registration</title>
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
    <style>
        body {
            padding: 20px;
            background-color: #f2f2f2;
        }
         .container {
            max-width: 500px;
            margin: 0 auto;
            padding: 20px;
            border: 3px solid #ccc;
            border-radius: 8px;
            box-shadow: 0px 0px 20px skyblue;
            background-color: transparent
           
        }
    </style>
</head>
<body>
    <div class="container">
        <h1 class="mt-5">Registration Form</h1>
        <div class="form-container">
            <form id="registrationForm" method="post">
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
                <button type="submit" class="btn btn-info">Register</button><br>
                Already a user? 
                <button type="button" class="btn btn-success" onclick="location.href='login.asp'">Login Here</button>
            </form>
        </div>
    </div>

    <script src="https://code.jquery.com/jquery-3.5.1.slim.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/@popperjs/core@2.5.4/dist/umd/popper.min.js"></script>
    <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>

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
       '' Response.Write(dbPath & "<br>")

        ' Connection string for Access database
        Dim connStr
        connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"

        ' Create a new connection object
        Dim conn
        Set conn = Server.CreateObject("ADODB.Connection")

        On Error Resume Next

        ' Open the database connection
        conn.Open connStr
        'Response.Write("DB Connected" & "<br>")

        ' Check for connection errors
        If Err.Number <> 0 Then
            Response.Write "An error occurred while connecting to the database."
            Response.End
        End If

        On Error Goto 0

        ' Prepare the SQL statement to insert data into the database
        Dim strSQLInsert
        strSQLInsert = "INSERT INTO [Users] ([FirstName], [LastName], [Email], [Phone], [Password], [Gender]) VALUES ('" & Replace(firstName, "'", "''") & "', '" & Replace(lastName, "'", "''") & "', '" & Replace(email, "'", "''") & "', '" & Replace(phone, "'", "''") & "', '" & Replace(password, "'", "''") & "', '" & Replace(gender, "'", "''") & "')"

        ' Execute the SQL insert statement
        'Response.Write "SQL Statement: " & strSQLInsert & "<br>"
        conn.Execute strSQLInsert

        ' Check for any errors during the insert operation
        If Err.Number <> 0 Then
            Response.Write "An error occurred while saving the form data."
            Response.End
        End If

        ' Close the database connection
        conn.Close
        Set conn = Nothing

        ' Display the submitted data

    End If
    %>
</body>
</html>
