<%
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

' Check if the ID parameter is provided
If Request.QueryString("id") <> "" Then
    ' Get the ID parameter value
    Dim userId
    userId = Request.QueryString("id")
    
    ' Retrieve the user record based on the ID
    Dim strSQLSelect
    strSQLSelect = "SELECT * FROM [Students] WHERE ID=" & userId
    
    ' Execute the SQL select statement
    Dim rs
    Set rs = conn.Execute(strSQLSelect)
    
    ' Check if the user record exists
    If Not rs.EOF Then
        ' Display the edit form with the user details
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
            
            ' Prepare the SQL statement to update the user record
            Dim strSQLUpdate
            strSQLUpdate = "UPDATE [Students] SET [FirstName]='" & Replace(firstName, "'", "''") & "', [LastName]='" & Replace(lastName, "'", "''") & "', [StudentID]='" & Replace(studentID, "'", "''") & "', [Email]='" & Replace(email, "'", "''") & "', [Gender]='" & Replace(gender, "'", "''") & "', [Session]='" & Replace(session, "'", "''") & "', [DateOfBirth]=#" & Replace(dateOfBirth, "'", "''") & "# WHERE ID=" & userId
            
            ' Execute the SQL update statement
            conn.Execute(strSQLUpdate)
            
            ' Check for any errors during the update operation
            If Err.Number <> 0 Then
                Response.Write "An error occurred while updating the user record."
                Response.End
            Else
                ' Redirect to show.asp after saving changes
                Response.Redirect "display.asp"
            End If
        Else
            ' Render the edit form
            %>
            <!DOCTYPE html>
            <html>
            <head>
                <title>Edit User</title>
                <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
                <style>
                    body {
                        padding: 20px;
                    }
                    .container {
                        max-width: 500px;
                        margin: 0 auto;
                    }
                </style>
            </head>
            <body>
                <div class="container mt-5">
                    <h1>Edit User</h1>
                    <form id="editForm" method="post">
                        <input type="hidden" name="id" value="<%= rs("ID") %>">
                        <div class="form-group">
                            <label for="firstName">First Name:</label>
                            <input type="text" class="form-control" id="firstName" name="firstName" value="<%= rs("FirstName") %>" required>
                        </div>
                        <div class="form-group">
                            <label for="lastName">Last Name:</label>
                            <input type="text" class="form-control" id="lastName" name="lastName" value="<%= rs("LastName") %>" required>
                        </div>
                        <div class="form-group">
                            <label for="studentID">Student ID:</label>
                            <input type="text" class="form-control" id="studentID" name="studentID" value="<%= rs("StudentID") %>" required>
                        </div>
                        <div class="form-group">
                            <label for="email">Email:</label>
                            <input type="email" class="form-control" id="email" name="email" value="<%= rs("Email") %>" required>
                        </div>
                        <div class="form-group">
                            <label for="gender">Gender:</label>
                            <select class="form-control" id="gender" name="gender" required>
                                <option value="male" <%= IIf(rs("Gender") = "male", "selected", "") %>>Male</option>
                                <option value="female" <%= IIf(rs("Gender") = "female", "selected", "") %>>Female</option>
                                <option value="other" <%= IIf(rs("Gender") = "other", "selected", "") %>>Other</option>
                            </select>
                        </div>
                        <div class="form-group">
                            <label for="session">Session:</label>
                            <input type="text" class="form-control" id="session" name="session" value="<%= rs("Session") %>" required>
                        </div>
                        <div class="form-group">
                            <label for="dateOfBirth">Date of Birth:</label>
                            <input type="date" class="form-control" id="dateOfBirth" name="dateOfBirth" value="<%= rs("DateOfBirth") %>" required>
                        </div>
                        <button type="submit" class="btn btn-primary">Save Changes</button>
                        <button type="button" class="btn btn-secondary" onclick="location.href='display.asp'">Cancel</button>
                    </form>
                </div>
            </body>
            </html>
            <%
        End If
    Else
        ' User record not found
        Response.Write "User record not found."
    End If
    
    rs.Close
Else
    ' ID parameter not provided
    Response.Write "Invalid request."
End If

' Close the database connection
conn.Close
Set conn = Nothing
%>