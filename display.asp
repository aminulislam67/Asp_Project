<!DOCTYPE html>
<html lang="en">
<head>
    <title>Display Students</title>
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
    <style>
        body {
            padding: 20px;
            background-color: #e4f4fc;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
            border: 3px solid #ccc;
            border-radius: 8px;
            background-color: #f7f7f7;
            margin-top: 20px;
        }
        .table-responsive {
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
    <h2>Display Students</h2>

    <form id="searchForm" method="get" action="">
        <div class="form-row">
            <div class="form-group col-md-6">
                <label for="searchById">Search by ID:</label>
                <input type="text" class="form-control" id="searchById" name="searchById" placeholder="Enter Student ID">
            </div>
            <div class="form-group col-md-6">
                <label for="searchByName">Search by Name:</label>
                <input type="text" class="form-control" id="searchByName" name="searchByName" placeholder="Enter Student Name">
            </div>
        </div>
        <button type="submit" class="btn btn-primary">Search</button>
    </form>

    <div class="table-responsive">
        <table class="table table-bordered">
            <thead>
                <tr>
                    <th>Student ID</th>
                    <th>First Name</th>
                    <th>Last Name</th>
                    <th>Email</th>
                    <th>Gender</th>
                    <th>Session</th>
                    <th>Date of Birth</th>
                </tr>
            </thead>
            <tbody>
                <% 
                ' Retrieve search parameters
                Dim searchById, searchByName
                searchById = Request.QueryString("searchById")
                searchByName = Request.QueryString("searchByName")

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

                ' Prepare the SQL statement to fetch student data from the database
                Dim strSQLFetch
               strSQLFetch = "SELECT * FROM [Students]"

                ' Append search conditions if provided
                If searchById <> "" Then
                    strSQLFetch = strSQLFetch & " WHERE [StudentID] = '" & Replace(searchById, "'", "''") & "'"
                ElseIf searchByName <> "" Then
                    strSQLFetch = strSQLFetch & " WHERE [FirstName] LIKE '%" & Replace(searchByName, "'", "''") & "%' OR [LastName] LIKE '%" & Replace(searchByName, "'", "''") & "%'"
                End If

                ' Execute the SQL select statement
                Dim rsFetch
                Set rsFetch = conn.Execute(strSQLFetch)

                ' Loop through the resultset and display student data
                Do Until rsFetch.EOF
                    Response.Write "<tr>"
                    Response.Write "<td>" & rsFetch("StudentID") & "</td>"
                    Response.Write "<td>" & rsFetch("FirstName") & "</td>"
                    Response.Write "<td>" & rsFetch("LastName") & "</td>"
                    Response.Write "<td>" & rsFetch("Email") & "</td>"
                    Response.Write "<td>" & rsFetch("Gender") & "</td>"
                    Response.Write "<td>" & rsFetch("Session") & "</td>"
                    Response.Write "<td>" & rsFetch("DateOfBirth") & "</td>"
                    Response.Write "</tr>"
                    rsFetch.MoveNext
                Loop

                rsFetch.Close

                ' Close the database connection
                conn.Close
                Set conn = Nothing
                %>
            </tbody>
        </table>
    </div>
</div>

<script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>

</body>
</html>
