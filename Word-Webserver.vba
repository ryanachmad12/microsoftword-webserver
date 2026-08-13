Option Explicit

' Server Config 

Public Const SERVER_HOST As String = "0.0.0.0"
Public Const SERVER_PORT As Long = 8081
Public Const WEB_ROOT As String = "C:\www"

Private Const SERVER_URL_HOST As String = "127.0.0.1"
Private Const POLL_SECONDS As Long = 1


' Server State

Private ServerRunning As Boolean
Private SocketOpen As Boolean
Private WinsockReady As Boolean
Private PollScheduled As Boolean

Private RequestCount As Long
Private NextPollTime As Date

Private InfoTable As Table
Private LogTable As Table

#If VBA7 Then
    Private ServerSocket As LongPtr
#Else
    Private ServerSocket As Long
#End If

' Winsock API Declarations

#If VBA7 Then

Private Declare PtrSafe Function WSAStartup Lib "ws2_32.dll" ( _
    ByVal wVersionRequested As Integer, _
    ByRef lpWSAData As WSAData) As Long

Private Declare PtrSafe Function WSACleanup Lib "ws2_32.dll" () As Long

Private Declare PtrSafe Function socket Lib "ws2_32.dll" ( _
    ByVal af As Long, _
    ByVal socketType As Long, _
    ByVal protocol As Long) As LongPtr

Private Declare PtrSafe Function closesocket Lib "ws2_32.dll" ( _
    ByVal s As LongPtr) As Long

Private Declare PtrSafe Function bind Lib "ws2_32.dll" ( _
    ByVal s As LongPtr, _
    ByRef name As SOCKADDR_IN, _
    ByVal namelen As Long) As Long

Private Declare PtrSafe Function listen Lib "ws2_32.dll" ( _
    ByVal s As LongPtr, _
    ByVal backlog As Long) As Long

Private Declare PtrSafe Function accept Lib "ws2_32.dll" ( _
    ByVal s As LongPtr, _
    ByRef addr As SOCKADDR_IN, _
    ByRef addrlen As Long) As LongPtr

Private Declare PtrSafe Function recv Lib "ws2_32.dll" ( _
    ByVal s As LongPtr, _
    ByRef buf As Any, _
    ByVal length As Long, _
    ByVal flags As Long) As Long

Private Declare PtrSafe Function send Lib "ws2_32.dll" ( _
    ByVal s As LongPtr, _
    ByRef buf As Any, _
    ByVal length As Long, _
    ByVal flags As Long) As Long

Private Declare PtrSafe Function ioctlsocket Lib "ws2_32.dll" ( _
    ByVal s As LongPtr, _
    ByVal cmd As Long, _
    ByRef argp As Long) As Long

Private Declare PtrSafe Function WSAGetLastError Lib "ws2_32.dll" () As Long

Private Declare PtrSafe Function htons Lib "ws2_32.dll" ( _
    ByVal hostshort As Integer) As Integer

#Else

Private Declare Function WSAStartup Lib "ws2_32.dll" ( _
    ByVal wVersionRequested As Integer, _
    ByRef lpWSAData As WSAData) As Long

Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long

Private Declare Function socket Lib "ws2_32.dll" ( _
    ByVal af As Long, _
    ByVal socketType As Long, _
    ByVal protocol As Long) As Long

Private Declare Function closesocket Lib "ws2_32.dll" ( _
    ByVal s As Long) As Long

Private Declare Function bind Lib "ws2_32.dll" ( _
    ByVal s As Long, _
    ByRef name As SOCKADDR_IN, _
    ByVal namelen As Long) As Long

Private Declare Function listen Lib "ws2_32.dll" ( _
    ByVal s As Long, _
    ByVal backlog As Long) As Long

Private Declare Function accept Lib "ws2_32.dll" ( _
    ByVal s As Long, _
    ByRef addr As SOCKADDR_IN, _
    ByRef addrlen As Long) As Long

Private Declare Function recv Lib "ws2_32.dll" ( _
    ByVal s As Long, _
    ByRef buf As Any, _
    ByVal length As Long, _
    ByVal flags As Long) As Long

Private Declare Function send Lib "ws2_32.dll" ( _
    ByVal s As Long, _
    ByRef buf As Any, _
    ByVal length As Long, _
    ByVal flags As Long) As Long

Private Declare Function ioctlsocket Lib "ws2_32.dll" ( _
    ByVal s As Long, _
    ByVal cmd As Long, _
    ByRef argp As Long) As Long

Private Declare Function WSAGetLastError Lib "ws2_32.dll" () As Long

Private Declare Function htons Lib "ws2_32.dll" ( _
    ByVal hostshort As Integer) As Integer

#End If

' Types for Winsock structures

Private Type WSAData
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To 256) As Byte
    szSystemStatus(0 To 128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
#If VBA7 Then
    lpVendorInfo As LongPtr
#Else
    lpVendorInfo As Long
#End If
End Type

Private Type SOCKADDR_IN
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero(0 To 7) As Byte
End Type

Private Const AF_INET As Long = 2
Private Const SOCK_STREAM As Long = 1
Private Const IPPROTO_TCP As Long = 6
Private Const FIONBIO As Long = &H8004667E
Private Const INVALID_SOCKET As Long = -1

' SetupDashboard: Initialize the dashboard, create tables for title, info, buttons, and log, and start the server.

Public Sub SetupDashboard()

    Dim doc As Document
    Dim titleTable As Table
    Dim buttonTable As Table
    Dim r As Range

    On Error GoTo ErrorHandler

    Set doc = ActiveDocument

    ServerRunning = False
    PollScheduled = False
    RequestCount = 0

    CloseServer

    EnsureWebRoot

    doc.Content.Delete

    Set InfoTable = Nothing
    Set LogTable = Nothing

' Title table: Create a table for the dashboard title and format it with a distinct background color.

    Set titleTable = AddTable(doc, 1, 1)

    With titleTable.Cell(1, 1)

        .Range.text = _
            "SERVER MANAGEMENT" & vbCr & _
            "LOCAL WEB SERVER CONTROL PANEL"

        .Shading.BackgroundPatternColor = _
            RGB(31, 78, 121)

        .VerticalAlignment = _
            wdCellAlignVerticalCenter

        With .Range
            .Font.name = "Calibri"
            .Font.Size = 16
            .Font.Bold = True
            .Font.Color = wdColorWhite
            .ParagraphFormat.Alignment = _
                wdAlignParagraphCenter
        End With

    End With

' Info table: Create a table to display server status, engine, host, port, web root, URL, request count, and last action.

    Set InfoTable = AddTable(doc, 8, 2)

    SetInfo 1, "SERVER STATUS", "STOPPED"
    SetInfo 2, "SERVER ENGINE", "WINWORD.EXE / WINSOCK"
    SetInfo 3, "SERVER HOST", SERVER_HOST
    SetInfo 4, "SERVER PORT", CStr(SERVER_PORT)
    SetInfo 5, "WEB ROOT", WEB_ROOT
    SetInfo 6, "SERVER URL", GetServerURL()
    SetInfo 7, "REQUEST COUNT", "0"
    SetInfo 8, "LAST ACTION", "Dashboard initialized."

    FormatInfoTable

' Button table: Create a table with buttons for starting, stopping, opening the web server, and clearing the log.

    Set buttonTable = AddTable(doc, 2, 4)

    buttonTable.Cell(1, 1).Range.text = "START"
    buttonTable.Cell(1, 2).Range.text = "STOP"
    buttonTable.Cell(1, 3).Range.text = "WEB"
    buttonTable.Cell(1, 4).Range.text = "CLEAR LOG"

    FormatHeader buttonTable.Cell(1, 1)
    FormatHeader buttonTable.Cell(1, 2)
    FormatHeader buttonTable.Cell(1, 3)
    FormatHeader buttonTable.Cell(1, 4)

' Start button: Create a button to start the server and format it with a distinct color.

    Set r = buttonTable.Cell(2, 1).Range
    r.End = r.End - 1
    r.text = ""
    r.Collapse wdCollapseStart

    ActiveDocument.Fields.Add _
        Range:=r, _
        Type:=wdFieldMacroButton, _
        text:="StartServer START SERVER", _
        PreserveFormatting:=False

    FormatButton _
        buttonTable.Cell(2, 1), _
        RGB(31, 78, 121)

' Stop button: Create a macro button to stop the server and format it.

    Set r = buttonTable.Cell(2, 2).Range
    r.End = r.End - 1
    r.text = ""
    r.Collapse wdCollapseStart

    ActiveDocument.Fields.Add _
        Range:=r, _
        Type:=wdFieldMacroButton, _
        text:="StopServer STOP SERVER", _
        PreserveFormatting:=False

    FormatButton _
        buttonTable.Cell(2, 2), _
        RGB(192, 0, 0)

' Web button: Create a button to open the web server URL in the default browser.

    Set r = buttonTable.Cell(2, 3).Range
    r.End = r.End - 1
    r.text = ""
    r.Collapse wdCollapseStart

    ActiveDocument.Fields.Add _
        Range:=r, _
        Type:=wdFieldMacroButton, _
        text:="OpenWeb OPEN WEB", _
        PreserveFormatting:=False

    FormatButton _
        buttonTable.Cell(2, 3), _
        RGB(89, 89, 89)

' Clear log button: Create a button to clear the log table and format it with a distinct color.

    Set r = buttonTable.Cell(2, 4).Range
    r.End = r.End - 1
    r.text = ""
    r.Collapse wdCollapseStart

    ActiveDocument.Fields.Add _
        Range:=r, _
        Type:=wdFieldMacroButton, _
        text:="ClearLog CLEAR LOG", _
        PreserveFormatting:=False

    FormatButton _
        buttonTable.Cell(2, 4), _
        RGB(112, 48, 160)

' Log table: Create a log table to display server events and initialize it with headers.

    Set LogTable = AddTable(doc, 1, 2)

    LogTable.Cell(1, 1).Range.text = "TIME"
    LogTable.Cell(1, 2).Range.text = "EVENT"

    FormatHeader LogTable.Cell(1, 1)
    FormatHeader LogTable.Cell(1, 2)

    AddLog "SYSTEM", "Dashboard initialized."
    AddLog "SYSTEM", "Engine: WINWORD.EXE / WINSOCK."
    AddLog "SYSTEM", "Bind: " & SERVER_HOST
    AddLog "SYSTEM", "Port: " & CStr(SERVER_PORT)
    AddLog "SYSTEM", "Web root: " & WEB_ROOT

    SetStatus "STOPPED"

    doc.Activate

' Automatically start the server after setting up the dashboard

    StartServer

    Exit Sub

ErrorHandler:

    ServerRunning = False
    CloseServer

    MsgBox _
        "SetupDashboard gagal." & vbCrLf & vbCrLf & _
        "Error " & CStr(Err.Number) & _
        ": " & Err.Description, _
        vbCritical, _
        "SERVER MANAGEMENT"

End Sub

' Table creation function: Add a new table to the document with specified rows and columns, and format it.

Private Function AddTable( _
    ByVal doc As Document, _
    ByVal rows As Long, _
    ByVal cols As Long) As Table

    Dim r As Range
    Dim t As Table

    Set r = doc.Range( _
        doc.Content.End - 1, _
        doc.Content.End - 1)

    Set t = doc.Tables.Add( _
        Range:=r, _
        NumRows:=rows, _
        NumColumns:=cols)

    t.Borders.Enable = True
    t.AllowAutoFit = True

    Set AddTable = t

    Set r = doc.Range( _
        doc.Content.End - 1, _
        doc.Content.End - 1)

    r.InsertParagraphAfter

End Function

Private Sub SetInfo( _
    ByVal rowNumber As Long, _
    ByVal labelText As String, _
    ByVal valueText As String)

    If InfoTable Is Nothing Then Exit Sub

    If rowNumber < 1 Then Exit Sub

    If rowNumber > InfoTable.rows.Count Then Exit Sub

    InfoTable.Cell( _
        rowNumber, 1).Range.text = labelText

    InfoTable.Cell( _
        rowNumber, 2).Range.text = valueText

End Sub

Private Sub FormatInfoTable()

    Dim i As Long

    If InfoTable Is Nothing Then Exit Sub

    For i = 1 To InfoTable.rows.Count

        With InfoTable.Cell(i, 1)

            .Shading.BackgroundPatternColor = _
                RGB(242, 242, 242)

            .Range.Font.name = "Calibri"
            .Range.Font.Size = 9
            .Range.Font.Bold = True

        End With

        With InfoTable.Cell(i, 2)

            .Range.Font.name = "Calibri"
            .Range.Font.Size = 9

        End With

    Next i

End Sub

Private Sub FormatHeader(ByVal c As Cell)

    With c

        .Shading.BackgroundPatternColor = _
            RGB(55, 55, 55)

        .VerticalAlignment = _
            wdCellAlignVerticalCenter

        With .Range
            .Font.name = "Calibri"
            .Font.Size = 9
            .Font.Bold = True
            .Font.Color = wdColorWhite
            .ParagraphFormat.Alignment = _
                wdAlignParagraphCenter
        End With

    End With

End Sub

Private Sub FormatButton( _
    ByVal c As Cell, _
    ByVal backgroundColor As Long)

    With c

        .Shading.BackgroundPatternColor = _
            backgroundColor

        .VerticalAlignment = _
            wdCellAlignVerticalCenter

        With .Range
            .Font.name = "Calibri"
            .Font.Size = 10
            .Font.Bold = True
            .Font.Color = wdColorWhite
            .ParagraphFormat.Alignment = _
                wdAlignParagraphCenter
        End With

    End With

End Sub

' Start the server, initialize Winsock, create a socket, bind it to the specified port, and listen for incoming connections.

Public Sub StartServer()

    Dim wsa As WSAData
    Dim result As Long
    Dim addr As SOCKADDR_IN
    Dim nonBlocking As Long

#If VBA7 Then
    Dim newSocket As LongPtr
#Else
    Dim newSocket As Long
#End If

    On Error GoTo ErrorHandler

    If InfoTable Is Nothing Then

        MsgBox _
            "Dashboard belum dibuat." & vbCrLf & _
            "Jalankan SetupDashboard terlebih dahulu.", _
            vbExclamation, _
            "SERVER MANAGEMENT"

        Exit Sub

    End If

    If ServerRunning Then

        SetStatus "RUNNING"

        UpdateInfo _
            "LAST ACTION", _
            "Server already running."

        Exit Sub

    End If

    SetStatus "STARTING"

    UpdateInfo _
        "LAST ACTION", _
        "Starting server..."

    EnsureWebRoot

' Winsock initialization: Initialize Winsock if it hasn't been initialized yet.

    If Not WinsockReady Then

        result = WSAStartup(&H202, wsa)

        If result <> 0 Then

            SetStatus "START FAILED"

            AddLog _
                "ERROR", _
                "WSAStartup failed: " & _
                CStr(result)

            Exit Sub

        End If

        WinsockReady = True

    End If

' Socket creation: Create a new TCP socket for the server.

    newSocket = socket( _
        AF_INET, _
        SOCK_STREAM, _
        IPPROTO_TCP)

    If newSocket = INVALID_SOCKET Then

        AddLog _
            "ERROR", _
            "socket() failed. WSA=" & _
            CStr(WSAGetLastError())

        CloseServer

        SetStatus "START FAILED"

        Exit Sub

    End If

    ServerSocket = newSocket
    SocketOpen = True

' Bind the server socket to the specified port and listen for incoming connections.

    addr.sin_family = AF_INET
    addr.sin_port = htons(CInt(SERVER_PORT))
    addr.sin_addr = 0

    result = bind( _
        ServerSocket, _
        addr, _
        LenB(addr))

    If result <> 0 Then

        AddLog _
            "ERROR", _
            "Bind failed on port " & _
            CStr(SERVER_PORT) & _
            ". WSA=" & _
            CStr(WSAGetLastError())

        CloseServer

        SetStatus "PORT BUSY"

        UpdateInfo _
            "LAST ACTION", _
            "Port " & CStr(SERVER_PORT) & _
            " is unavailable."

        Exit Sub

    End If

' Listen for incoming connections on the server socket with a backlog of 20.

    result = listen( _
        ServerSocket, _
        20)

    If result <> 0 Then

        AddLog _
            "ERROR", _
            "listen() failed. WSA=" & _
            CStr(WSAGetLastError())

        CloseServer

        SetStatus "START FAILED"

        Exit Sub

    End If

' Non-blocking mode: Set the server socket to non-blocking mode so that accept() does not block the main thread.

    nonBlocking = 1

    result = ioctlsocket( _
        ServerSocket, _
        FIONBIO, _
        nonBlocking)

    If result <> 0 Then

        AddLog _
            "ERROR", _
            "Unable to set non-blocking socket."

        CloseServer

        SetStatus "START FAILED"

        Exit Sub

    End If

' Reset request count and mark server as running

    RequestCount = 0
    ServerRunning = True

    SetStatus "RUNNING"

    UpdateInfo _
        "REQUEST COUNT", _
        "0"

    UpdateInfo _
        "LAST ACTION", _
        "Server started successfully."

    AddLog _
        "SERVER", _
        "Server started successfully."

    AddLog _
        "SERVER", _
        "Listening on 0.0.0.0:" & _
        CStr(SERVER_PORT)

    AddLog _
        "SERVER", _
        "URL: " & GetServerURL()

    SchedulePoll

    Exit Sub

ErrorHandler:

    ServerRunning = False

    CloseServer

    SetStatus "START FAILED"

    AddLog _
        "ERROR", _
        "StartServer " & _
        CStr(Err.Number) & _
        ": " & Err.Description

End Sub

' Stop the server, close the socket, and clean up Winsock resources.

Public Sub StopServer()

    On Error Resume Next

    ServerRunning = False
    PollScheduled = False

    CloseServer

    SetStatus "STOPPED"

    UpdateInfo _
        "LAST ACTION", _
        "Server stopped."

    AddLog _
        "SERVER", _
        "Server stopped."

    AddLog _
        "SERVER", _
        "Socket closed."

    AddLog _
        "SERVER", _
        "Port " & _
        CStr(SERVER_PORT) & _
        " released."

End Sub

' Close the server socket and clean up Winsock resources.

Private Sub CloseServer()

    On Error Resume Next

    If SocketOpen Then

        closesocket ServerSocket

        SocketOpen = False

    End If

    If WinsockReady Then

        WSACleanup

        WinsockReady = False

    End If

End Sub

' While the server is running, schedule a poll to check for incoming connections every POLL_SECONDS seconds.

Private Sub SchedulePoll()

    On Error GoTo Failed

    If Not ServerRunning Then Exit Sub

    If Not SocketOpen Then Exit Sub

    If PollScheduled Then Exit Sub

    NextPollTime = _
        Now + TimeSerial(0, 0, POLL_SECONDS)

    PollScheduled = True

    Application.OnTime _
        When:=NextPollTime, _
        name:="PollServer"

    Exit Sub

Failed:

    PollScheduled = False

End Sub

Public Sub PollServer()

    Dim addr As SOCKADDR_IN
    Dim addrLength As Long

#If VBA7 Then
    Dim client As LongPtr
#Else
    Dim client As Long
#End If

    PollScheduled = False

    If Not ServerRunning Then Exit Sub

    If Not SocketOpen Then Exit Sub

    On Error GoTo ErrorHandler

    addrLength = LenB(addr)

    client = accept( _
        ServerSocket, _
        addr, _
        addrLength)

    If client <> INVALID_SOCKET Then

        HandleClient client

    End If

    If ServerRunning Then
        SchedulePoll
    End If

    Exit Sub

ErrorHandler:

    If ServerRunning Then
        SchedulePoll
    End If

End Sub

' Client handling function: Receive the HTTP request, parse it, and serve the requested file or send an error response.

#If VBA7 Then

Private Sub HandleClient(ByVal client As LongPtr)

#Else

Private Sub HandleClient(ByVal client As Long)

#End If

    Dim buffer(0 To 16383) As Byte
    Dim received As Long
    Dim requestText As String
    Dim requestLine As String
    Dim method As String
    Dim path As String

    On Error GoTo ErrorHandler

    received = recv( _
        client, _
        buffer(0), _
        16384, _
        0)

    If received <= 0 Then

        closesocket client

        Exit Sub

    End If

    requestText = _
        BytesToString(buffer, received)

    requestLine = _
        GetFirstLine(requestText)

    ParseRequest _
        requestLine, _
        method, _
        path

    If method = "" Then

        SendError _
            client, _
            400, _
            "Bad Request"

        closesocket client

        Exit Sub

    End If

    RequestCount = RequestCount + 1

    UpdateInfo _
        "REQUEST COUNT", _
        CStr(RequestCount)

    AddLog _
        "HTTP", _
        method & " " & path

    If method <> "GET" And _
       method <> "HEAD" Then

        SendError _
            client, _
            405, _
            "Method Not Allowed"

        closesocket client

        Exit Sub

    End If

    ServeFile _
        client, _
        path, _
        (method = "HEAD")

    closesocket client

    Exit Sub

ErrorHandler:

    On Error Resume Next

    closesocket client

End Sub

' Serve file function: Serve the requested file to the client, handling errors and sending appropriate HTTP responses.

#If VBA7 Then

Private Sub ServeFile( _
    ByVal client As LongPtr, _
    ByVal urlPath As String, _
    ByVal headOnly As Boolean)

#Else

Private Sub ServeFile( _
    ByVal client As Long, _
    ByVal urlPath As String, _
    ByVal headOnly As Boolean)

#End If

    Dim cleanPath As String
    Dim filePath As String
    Dim fileData() As Byte
    Dim header As String

    On Error GoTo ErrorHandler

    cleanPath = urlPath

    If InStr(cleanPath, "?") > 0 Then
        cleanPath = _
            Left$( _
                cleanPath, _
                InStr(cleanPath, "?") - 1)
    End If

    cleanPath = Replace( _
        cleanPath, _
        "/", _
        "\")

    Do While Left$(cleanPath, 1) = "\"
        cleanPath = Mid$(cleanPath, 2)
    Loop

    If cleanPath = "" Then
        cleanPath = "index.html"
    End If

    If InStr( _
        1, _
        cleanPath, _
        "..", _
        vbBinaryCompare) > 0 Then

        SendError _
            client, _
            403, _
            "Forbidden"

        Exit Sub

    End If

    filePath = _
        WEB_ROOT & "\" & cleanPath

    If Not FileExists(filePath) Then

        SendError _
            client, _
            404, _
            "Not Found"

        AddLog _
            "HTTP", _
            "404 " & urlPath

        Exit Sub

    End If

    If Not ReadFile( _
        filePath, _
        fileData) Then

        SendError _
            client, _
            500, _
            "Internal Server Error"

        Exit Sub

    End If

    header = _
        "HTTP/1.1 200 OK" & vbCrLf & _
        "Content-Type: " & _
        GetMimeType(filePath) & vbCrLf & _
        "Content-Length: " & _
        CStr(ByteLength(fileData)) & vbCrLf & _
        "Connection: close" & vbCrLf & _
        "Cache-Control: no-cache" & _
        vbCrLf & vbCrLf

    SendString _
        client, _
        header

    If Not headOnly Then

        SendBuffer _
            client, _
            fileData

    End If

    AddLog _
        "HTTP", _
        "200 " & urlPath

    Exit Sub

ErrorHandler:

    SendError _
        client, _
        500, _
        "Internal Server Error"

End Sub

' Send functions: Send string and send byte buffer to the client socket.

#If VBA7 Then

Private Sub SendString( _
    ByVal client As LongPtr, _
    ByVal text As String)

#Else

Private Sub SendString( _
    ByVal client As Long, _
    ByVal text As String)

#End If

    Dim data() As Byte

    If Len(text) = 0 Then Exit Sub

    data = StrConv( _
        text, _
        vbFromUnicode)

    SendBuffer _
        client, _
        data

End Sub

#If VBA7 Then

Private Sub SendBuffer( _
    ByVal client As LongPtr, _
    ByRef data() As Byte)

#Else

Private Sub SendBuffer( _
    ByVal client As Long, _
    ByRef data() As Byte)

#End If

    Dim total As Long
    Dim sent As Long
    Dim result As Long

    total = ByteLength(data)

    If total <= 0 Then Exit Sub

    Do While sent < total

        result = send( _
            client, _
            data(sent), _
            total - sent, _
            0)

        If result <= 0 Then Exit Do

        sent = sent + result

    Loop

End Sub

#If VBA7 Then

Private Sub SendError( _
    ByVal client As LongPtr, _
    ByVal statusCode As Long, _
    ByVal statusText As String)

#Else

Private Sub SendError( _
    ByVal client As Long, _
    ByVal statusCode As Long, _
    ByVal statusText As String)

#End If

    Dim body As String
    Dim data() As Byte

    body = _
        "<!DOCTYPE html>" & _
        "<html><head>" & _
        "<meta charset=""utf-8"">" & _
        "<title>" & _
        CStr(statusCode) & _
        "</title></head><body>" & _
        "<h1>" & CStr(statusCode) & _
        " - " & statusText & _
        "</h1></body></html>"

    data = StrConv( _
        body, _
        vbFromUnicode)

    SendString _
        client, _
        "HTTP/1.1 " & _
        CStr(statusCode) & " " & _
        statusText & vbCrLf & _
        "Content-Type: text/html; charset=utf-8" & vbCrLf & _
        "Content-Length: " & _
        CStr(ByteLength(data)) & vbCrLf & _
        "Connection: close" & _
        vbCrLf & vbCrLf

    SendBuffer _
        client, _
        data

End Sub

' Request parsing functions: Convert byte array to string, extract the first line of the request, and parse the HTTP method and path.

Private Function BytesToString( _
    ByRef data() As Byte, _
    ByVal length As Long) As String

    Dim i As Long
    Dim result As String

    If length <= 0 Then Exit Function

    result = String$( _
        length, _
        vbNullChar)

    For i = 0 To length - 1

        Mid$( _
            result, _
            i + 1, _
            1) = Chr$(data(i))

    Next i

    BytesToString = result

End Function

Private Function GetFirstLine( _
    ByVal requestText As String) As String

    Dim position As Long

    position = InStr( _
        1, _
        requestText, _
        vbCrLf)

    If position > 0 Then

        GetFirstLine = _
            Left$( _
                requestText, _
                position - 1)

    Else

        GetFirstLine = requestText

    End If

End Function

Private Sub ParseRequest( _
    ByVal requestLine As String, _
    ByRef method As String, _
    ByRef path As String)

    Dim parts() As String

    method = ""
    path = ""

    parts = Split( _
        Trim$(requestLine), _
        " ")

    If UBound(parts) >= 1 Then

        method = UCase$(parts(0))
        path = parts(1)

    End If

End Sub

' File handling functions: Read file content, check file existence, and determine MIME types.

Private Function ReadFile( _
    ByVal filePath As String, _
    ByRef data() As Byte) As Boolean

    Dim fileNumber As Integer
    Dim fileSize As Long

    On Error GoTo Failed

    fileNumber = FreeFile

    Open filePath For Binary Access Read As #fileNumber

    fileSize = LOF(fileNumber)

    If fileSize > 0 Then

        ReDim data(0 To fileSize - 1)

        Get #fileNumber, , data

    Else

        ReDim data(0 To 0)

    End If

    Close #fileNumber

    ReadFile = True

    Exit Function

Failed:

    On Error Resume Next
    Close #fileNumber

    ReadFile = False

End Function

Private Function FileExists( _
    ByVal filePath As String) As Boolean

    Dim fileNumber As Integer

    On Error GoTo Failed

    fileNumber = FreeFile

    Open filePath For Binary Access Read As #fileNumber

    Close #fileNumber

    FileExists = True

    Exit Function

Failed:

    On Error Resume Next
    Close #fileNumber

    FileExists = False

End Function

Private Function FolderExists( _
    ByVal folderPath As String) As Boolean

    Dim attributes As Long

    On Error GoTo Failed

    attributes = GetAttr(folderPath)

    FolderExists = _
        ((attributes And vbDirectory) <> 0)

    Exit Function

Failed:

    FolderExists = False

End Function

Private Function ByteLength( _
    ByRef data() As Byte) As Long

    On Error GoTo EmptyData

    ByteLength = UBound(data) + 1

    Exit Function

EmptyData:

    ByteLength = 0

End Function

' Mime type mapping based on file extension. Returns the appropriate MIME type for the given file path.

Private Function GetMimeType( _
    ByVal filePath As String) As String

    Dim extension As String

    extension = _
        LCase$( _
            GetExtension(filePath))

    Select Case extension

        Case "html", "htm"
            GetMimeType = _
                "text/html; charset=utf-8"

        Case "css"
            GetMimeType = _
                "text/css; charset=utf-8"

        Case "js"
            GetMimeType = _
                "application/javascript; charset=utf-8"

        Case "json"
            GetMimeType = _
                "application/json; charset=utf-8"

        Case "txt"
            GetMimeType = _
                "text/plain; charset=utf-8"

        Case "png"
            GetMimeType = "image/png"

        Case "jpg", "jpeg"
            GetMimeType = "image/jpeg"

        Case "gif"
            GetMimeType = "image/gif"

        Case "svg"
            GetMimeType = "image/svg+xml"

        Case "webp"
            GetMimeType = "image/webp"

        Case "ico"
            GetMimeType = "image/x-icon"

        Case "pdf"
            GetMimeType = "application/pdf"

        Case Else
            GetMimeType = _
                "application/octet-stream"

    End Select

End Function

Private Function GetExtension( _
    ByVal filePath As String) As String

    Dim position As Long

    position = InStrRev( _
        filePath, _
        ".")

    If position > 0 Then

        GetExtension = _
            Mid$( _
                filePath, _
                position + 1)

    End If

End Function

' Dashboard update functions: Update information and status in the dashboard.

Private Sub UpdateInfo( _
    ByVal labelText As String, _
    ByVal valueText As String)

    Dim i As Long
    Dim currentLabel As String

    On Error Resume Next

    If InfoTable Is Nothing Then Exit Sub

    For i = 1 To InfoTable.rows.Count

        currentLabel = _
            CleanText( _
                InfoTable.Cell( _
                    i, 1).Range.text)

        If StrComp( _
            currentLabel, _
            labelText, _
            vbTextCompare) = 0 Then

            InfoTable.Cell( _
                i, 2).Range.text = valueText

            Exit For

        End If

    Next i

End Sub

Private Sub SetStatus( _
    ByVal statusText As String)

    Dim i As Long
    Dim labelText As String

    If InfoTable Is Nothing Then Exit Sub

    UpdateInfo _
        "SERVER STATUS", _
        statusText

    For i = 1 To InfoTable.rows.Count

        labelText = _
            CleanText( _
                InfoTable.Cell( _
                    i, 1).Range.text)

        If StrComp( _
            labelText, _
            "SERVER STATUS", _
            vbTextCompare) = 0 Then

            With InfoTable.Cell(i, 2)

                Select Case UCase$(statusText)

                    Case "RUNNING"

                        .Shading.BackgroundPatternColor = _
                            RGB(0, 176, 80)

                        .Range.Font.Color = _
                            wdColorWhite

                    Case "STARTING"

                        .Shading.BackgroundPatternColor = _
                            RGB(255, 192, 0)

                        .Range.Font.Color = _
                            wdColorBlack

                    Case Else

                        .Shading.BackgroundPatternColor = _
                            RGB(192, 0, 0)

                        .Range.Font.Color = _
                            wdColorWhite

                End Select

                .Range.Font.Bold = True

                .Range.ParagraphFormat.Alignment = _
                    wdAlignParagraphCenter

            End With

            Exit For

        End If

    Next i

End Sub

' Logging function to add entries to the log table in the dashboard.

Private Sub AddLog( _
    ByVal source As String, _
    ByVal message As String)

    Dim newRow As Row

    On Error Resume Next

    If LogTable Is Nothing Then Exit Sub

    Set newRow = LogTable.rows.Add

    newRow.Cells(1).Range.text = _
        Format( _
            Now, _
            "yyyy-mm-dd hh:nn:ss")

    newRow.Cells(2).Range.text = _
        "[" & source & "] " & message

End Sub

Public Sub ClearLog()

    Dim i As Long

    On Error GoTo ErrorHandler

    If LogTable Is Nothing Then

        MsgBox _
            "Dashboard belum dibuat.", _
            vbExclamation, _
            "SERVER MANAGEMENT"

        Exit Sub

    End If

    Application.ScreenUpdating = False

    For i = LogTable.rows.Count To 2 Step -1

        LogTable.rows(i).Delete

    Next i

    AddLog _
        "SYSTEM", _
        "Log cleared."

    UpdateInfo _
        "LAST ACTION", _
        "Log cleared."

    Application.ScreenUpdating = True

    Exit Sub

ErrorHandler:

    Application.ScreenUpdating = True

    MsgBox _
        "Clear Log gagal." & vbCrLf & _
        CStr(Err.Number) & ": " & _
        Err.Description, _
        vbExclamation, _
        "SERVER MANAGEMENT"

End Sub

' Web browser integration: Open the default web browser to the server URL.

Public Sub OpenWeb()

    On Error GoTo ErrorHandler

    If Not ServerRunning Then

        MsgBox _
            "Server belum berjalan." & vbCrLf & _
            "Klik START SERVER terlebih dahulu.", _
            vbExclamation, _
            "SERVER MANAGEMENT"

        Exit Sub

    End If

    ActiveDocument.FollowHyperlink _
        Address:=GetServerURL()

    AddLog _
        "WEB", _
        "Opened " & GetServerURL()

    Exit Sub

ErrorHandler:

    MsgBox _
        "Browser tidak dapat dibuka." & vbCrLf & _
        CStr(Err.Number) & ": " & _
        Err.Description, _
        vbExclamation, _
        "SERVER MANAGEMENT"

End Sub

Private Function GetServerURL() As String

    GetServerURL = _
        "http://" & _
        SERVER_URL_HOST & ":" & _
        CStr(SERVER_PORT)

End Function

' Default Web Server Root Directory

Private Sub EnsureWebRoot()

    Dim fileNumber As Integer
    Dim indexFile As String

    On Error GoTo Failed

    If Not FolderExists(WEB_ROOT) Then

        MkDir WEB_ROOT

    End If

    indexFile = _
        WEB_ROOT & "\index.html"

    If FileExists(indexFile) Then Exit Sub

    fileNumber = FreeFile

    Open indexFile For Output As #fileNumber

    Print #fileNumber, "<!DOCTYPE html>"
    Print #fileNumber, "<html>"
    Print #fileNumber, "<head>"
    Print #fileNumber, "<meta charset=""UTF-8"">"
    Print #fileNumber, "<meta name=""viewport"" content=""width=device-width,initial-scale=1"">"
    Print #fileNumber, "<title>Word Web Server</title>"
    Print #fileNumber, "<style>"
    Print #fileNumber, "*{box-sizing:border-box}"
    Print #fileNumber, "body{margin:0;min-height:100vh;background:#f4f6f8;font-family:Arial,sans-serif;display:flex;align-items:center;justify-content:center}"
    Print #fileNumber, ".card{width:650px;max-width:90%;padding:50px;background:#fff;border-radius:18px;text-align:center;box-shadow:0 15px 40px rgba(0,0,0,.12)}"
    Print #fileNumber, "h1{margin:0 0 15px;color:#1f4e79}"
    Print #fileNumber, "p{color:#555}"
    Print #fileNumber, ".status{display:inline-block;margin-top:20px;padding:10px 25px;border-radius:30px;background:#e2f0d9;color:#006100;font-weight:bold}"
    Print #fileNumber, "</style>"
    Print #fileNumber, "</head>"
    Print #fileNumber, "<body>"
    Print #fileNumber, "<div class=""card"">"
    Print #fileNumber, "<h1>WORD WEB SERVER</h1>"
    Print #fileNumber, "<p>HTTP server hosted directly by Microsoft Word.</p>"
    Print #fileNumber, "<span class=""status"">SERVER ONLINE</span>"
    Print #fileNumber, "</div>"
    Print #fileNumber, "</body>"
    Print #fileNumber, "</html>"

    Close #fileNumber

    Exit Sub

Failed:

    On Error Resume Next
    Close #fileNumber

End Sub

' Utility function to clean text by removing carriage returns and bell characters, and trimming whitespace.

Private Function CleanText( _
    ByVal text As String) As String

    text = Replace( _
        text, _
        Chr$(13), _
        "")

    text = Replace( _
        text, _
        Chr$(7), _
        "")

    CleanText = Trim$(text)

End Function



