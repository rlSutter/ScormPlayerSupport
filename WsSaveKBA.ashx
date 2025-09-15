<%@ WebHandler Language="VB" Class="WsSaveKBA" %>

Imports System
Imports System.Web
Imports System.Configuration
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Web.Script.Serialization
Imports System.Xml
Imports System.Text
Imports Newtonsoft.Json.Converters
Imports log4net
Imports CachingWrapper.LocalCache

Public Class WsSaveKBA : Implements IHttpHandler
    
    Enum enumObjectType
        StrType = 0
        IntType = 1
        DblType = 2
        DteType = 3
    End Enum
    
    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        
        ' Parameter Declarations
        Dim Debug, callback As String
        
        ' Result Declarations
        Dim jdoc As String
        Dim results As String
        
        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String

        ' Caching declarations
        Dim WeightCache As New CachingWrapper.LocalCache
        Dim wttbl As DataTable = New DataTable          ' Weight table
        Dim fetbl As DataTable = New DataTable          ' Special Fees table
        
        ' Logging declarations
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("SaveKBADebugLog")
        Dim logfile, tempdebug As String
        Dim Logging As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"
        
        ' Context declarations
        Dim callip As String = context.Request.ServerVariables("HTTP_X_FORWARDED_FOR")
        If callip Is Nothing Then
            callip = context.Request.UserHostAddress
        Else
            If callip.Contains(",") Then
                callip = Left(callip, callip.IndexOf(",") - 1)
            Else
                callip = callip
            End If
        End If
        Dim PrevLink As String = Trim(context.Request.ServerVariables("HTTP_REFERER"))
        Dim BROWSER As String = Trim(context.Request.ServerVariables("HTTP_USER_AGENT"))
                
        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service
        Dim Processing As New com.certegrity.cloudsvc.processing.Service
        
        ' Variable declarations
        Dim errmsg, ErrLvl, temp As String
        Dim UserId As String
        Dim REG_ID, REG_NUM, HOME_PAGE, LANG_CD, UID, SessID, EMAIL_ADDR, COURSE, JURIS, FST_NAME, LAST_NAME, KBA_NOTICE As String
        Dim LOGGED_IN, CONTACT_ID, SUB_ID, CONTACT_OU_ID, DOMAIN As String
        Dim SCtr, CrLf, SUCCESS As String
        Dim NUM_QUESTIONS, ctr, ntemp As Integer
        Dim Q_ID(50) As String
        Dim Q_TEXT(50) As String
        Dim A_TEXT(50) As String
        Dim ROW_ID As String
        Dim SendFrom, SendTo, QUESTIONS, MergeData As String
 
        ' ============================================
        ' Variable setup
        Debug = "N"
        Logging = "Y"
        errmsg = ""
        UserId = ""
        LANG_CD = "ENU"
        SUCCESS = "True"
        NUM_QUESTIONS = 0
        CrLf = Chr(10)
        DOMAIN = "TIPS"
        HOME_PAGE = ""
        SessID = ""
        REG_ID = ""
        REG_NUM = ""
        EMAIL_ADDR = ""
        COURSE = ""
        JURIS = ""
        FST_NAME = ""
        LAST_NAME = ""
        KBA_NOTICE = "N"
        UID = ""
        LOGGED_IN = "N"
        CONTACT_ID = ""
        CONTACT_OU_ID = ""
        QUESTIONS = ""
        MergeData = ""
        callback = ""
        jdoc = ""
        ErrLvl = "Error"
        
        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
            tempdebug = System.Configuration.ConfigurationManager.AppSettings.Get("SaveKBA_debug")
            If tempdebug = "Y" And Debug <> "T" Then Debug = "Y"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            GoTo CloseOut2
        End Try
        
        ' ============================================
        ' Open log file if applicable
        If Logging = "Y" Then
            logfile = "C:\Logs\WsSaveKBA.log"
            Try
                log4net.GlobalContext.Properties("SaveKBALogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                GoTo CloseOut2
            End Try
        End If
        
        ' ============================================
        ' Get cookie values
        UID = context.Request.Cookies("ID").Value
        
        ' ============================================
        ' Get Context parameters    
        '   REG_NUM         - The user's S_CONTACT.X_REGISTRATION_NUM
        '   REG_ID          - The ROW_ID of the CX_SESS_REG record
        '   SessID          - The user's session id
        '   DOMAIN          - The root domain that invoked this service
        '   HOME_PAGE       - The root page that invoked this service
        '   LANG_CD         - The language code of the student
        '   Callback        - The name of the Javascript callback function in which to wrap the resulting JSON . Default freightCalc         
        
        If Not context.Request.QueryString("RID") Is Nothing Then
            REG_ID = context.Request.QueryString("RID")
        End If
        
        If Not context.Request.QueryString("RG") Is Nothing Then
            REG_NUM = context.Request.QueryString("RG")
        End If

        If Not context.Request.QueryString("SESS") Is Nothing Then
            SessID = context.Request.QueryString("SESS")
        End If

        If Not context.Request.QueryString("DOMAIN") Is Nothing Then
            DOMAIN = UCase(context.Request.QueryString("DOMAIN"))
        End If
        If DOMAIN = "" Then DOMAIN = "TIPS"

        If Not context.Request.QueryString("HP") Is Nothing Then
            HOME_PAGE = context.Request.QueryString("HP")
        End If

        If Not context.Request.QueryString("LANG") Is Nothing Then
            LANG_CD = UCase(context.Request.QueryString("LANG"))
        End If
        
        If Not context.Request.QueryString("callback") Is Nothing Then
            callback = context.Request.QueryString("callback")
        End If
        
        If Debug = "Y" Then
            mydebuglog.Debug("----------------------------------")
            mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
            mydebuglog.Debug("Parameters-")
            mydebuglog.Debug("  REG_ID: " & REG_ID)
            mydebuglog.Debug("  REG_NUM: " & REG_NUM)
            mydebuglog.Debug("  UID: " & UID)
            mydebuglog.Debug("  SessID: " & SessID)
            mydebuglog.Debug("  DOMAIN: " & DOMAIN)
            mydebuglog.Debug("  HOME_PAGE: " & HOME_PAGE)
            mydebuglog.Debug("  LANG_CD: " & LANG_CD)
        End If

        ' Get KBA parameters    
        '   EA              - Email address
        '   CRS             - The course 
        '   JUR             - The jurisdiction id
        '   FST             - First name
        '   LST             - Last name
        '   KN              - Whether to send a KBA notice to the student
        '   NQ              - The number of question answers to store
        '   ID_#            - The ids of the answers (.. repeated to the number of answers)
        '   A_#	            - The text of the answers (.. repeated to the number of answers)

        If Not context.Request.QueryString("EA") Is Nothing Then
            EMAIL_ADDR = context.Request.QueryString("EA")
        End If
 
        If Not context.Request.QueryString("CRS") Is Nothing Then
            COURSE = context.Request.QueryString("CRS")
        End If

        If Not context.Request.QueryString("JUR") Is Nothing Then
            JURIS = context.Request.QueryString("JUR")
        End If
        
        If Not context.Request.QueryString("FST") Is Nothing Then
            FST_NAME = context.Request.QueryString("FST")
        End If
        
        If Not context.Request.QueryString("LST") Is Nothing Then
            LAST_NAME = context.Request.QueryString("LST")
        End If
        
        If Not context.Request.QueryString("KN") Is Nothing Then
            KBA_NOTICE = context.Request.QueryString("KN")
        End If
        
        If Not context.Request.QueryString("NQ") Is Nothing Then
            temp = context.Request.QueryString("NQ")
            If IsNumeric(temp) Then NUM_QUESTIONS = Val(temp)
        End If

        If Debug = "Y" Then
            mydebuglog.Debug("  EMAIL_ADDR: " & EMAIL_ADDR)
            mydebuglog.Debug("  COURSE: " & COURSE)
            mydebuglog.Debug("  JURIS: " & JURIS)
            mydebuglog.Debug("  FST_NAME: " & FST_NAME)
            mydebuglog.Debug("  LAST_NAME: " & LAST_NAME)
            mydebuglog.Debug("  KBA_NOTICE: " & KBA_NOTICE)
            mydebuglog.Debug("  NUM_QUESTIONS: " & Str(NUM_QUESTIONS) & vbCrLf)
        End If

        ' Get answers and ids
        ReDim Q_ID(NUM_QUESTIONS)
        ReDim A_TEXT(NUM_QUESTIONS)
        ReDim Q_TEXT(NUM_QUESTIONS)
        For i = 1 To NUM_QUESTIONS
            SCtr = Trim(Str(i))
            If Not context.Request.QueryString("ID_" & SCtr) Is Nothing Then
                Q_ID(i) = context.Request.QueryString("ID_" & SCtr)
                A_TEXT(i) = context.Request.QueryString("A_" & SCtr)
                If Debug = "Y" Then mydebuglog.Debug("  " & SCtr & " : " & Q_ID(i) & " .. " & A_TEXT(i))
            End If
        Next
        
        ' ============================================
        ' Validate 
        'If REG_NUM <> UID Then GoTo AccessError        
        If NUM_QUESTIONS = 0 Then GoTo DataError
        If REG_ID = "" Then GoTo DataError
        
        ' ============================================
        ' Open database connection 
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            errmsg = errmsg & "Unable to open the database connection. " & vbCrLf
            GoTo CloseOut
        End If

        ' ============================================
        ' Prepare results
        If Not cmd Is Nothing Then
                        
            ' ================================================   
            ' GET USER PROFILE
            SqlS = "SELECT TOP 1 (SELECT CASE WHEN H.LOGOUT_DT IS NULL THEN 'Y' ELSE 'N' END) AS LOGGED_IN, SC.CON_ID, SC.SUB_ID, C.PR_DEPT_OU_ID, S.DOMAIN " & _
                "FROM siebeldb.dbo.CX_SUB_CON_HIST H " & _
                "LEFT OUTER JOIN siebeldb.dbo.CX_SUB_CON SC ON SC.ROW_ID=H.SUB_CON_ID " & _
                "LEFT OUTER JOIN siebeldb.dbo.CX_SUBSCRIPTION S ON S.ROW_ID=SC.SUB_ID " & _
                "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=SC.CON_ID " & _
                "WHERE USER_ID='" & REG_NUM & "' AND SESSION_ID='" & SessID & "'"
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Verify user: " & vbCrLf & "  " & SqlS)
            Try
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        LOGGED_IN = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                        CONTACT_ID = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                        SUB_ID = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                        CONTACT_OU_ID = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                        If DOMAIN = "" Then DOMAIN = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                    End While
                End If
            Catch ex As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Unable to locate credentials. Error: " & vbCrLf & ex.ToString & vbCrLf)
                GoTo AccessError
            End Try
            dr.Close()
            If Debug = "Y" Then
                mydebuglog.Debug("  .. LOGGED_IN: " & LOGGED_IN)
                mydebuglog.Debug("  .. CONTACT_ID: " & CONTACT_ID)
                mydebuglog.Debug("  .. CONTACT_OU_ID: " & CONTACT_OU_ID)
                mydebuglog.Debug("  .. DOMAIN: " & DOMAIN)
            End If
            If LOGGED_IN <> "Y" Then GoTo AccessError
            
            ' ======================================================
            ' DELETE ANY EXISTING ANSWERS
            SqlS = "DELETE FROM elearning.dbo.KBA_ANSR WHERE REG_ID='" & REG_ID & "'"
            temp = ExecQuery("Delete", "KBA_ANSR", cmd, SqlS, mydebuglog, Debug)
            
            ' ======================================================
            ' CREATE RECORDS FROM ANSWERS
            ctr = 0
            ntemp = 0
            For i = 1 To NUM_QUESTIONS
                ROW_ID = Trim(REG_ID & "-" & Trim(Str(i)))
InsertQuestion:
                SqlS = "INSERT elearning.dbo.KBA_ANSR (ROW_ID,REG_ID,USER_ID,QUES_ID,ANSR_TEXT) " & _
                  "VALUES ('" & ROW_ID & "','" & REG_ID & "','" & REG_NUM & "','" & Q_ID(i) & "','" & SqlString(A_TEXT(i)) & "')"
                temp = ExecQuery("Insert", "KBA_ANSR", cmd, SqlS, mydebuglog, Debug)

                ' Verify that the answers were written
                temp = ""
                SqlS = "SELECT A.ROW_ID, Q.QUES_TEXT " & _
                "FROM elearning.dbo.KBA_ANSR A " & _
                "LEFT OUTER JOIN elearning.dbo.KBA_QUES Q ON Q.ROW_ID=A.QUES_ID " & _
                "WHERE A.ROW_ID='" & ROW_ID & "'"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  VERIFY ID " & ROW_ID & " INSERT TO KBA_ANSR:  " & SqlS)
                Try
                    cmd.CommandText = SqlS
                    dr = cmd.ExecuteReader()
                    If Not dr Is Nothing Then
                        While dr.Read()
                            temp = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            Q_TEXT(i) = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                        End While
                    End If
                Catch ex As Exception
                End Try
                If Debug = "Y" Then mydebuglog.Debug("    .. Q_TEXT: " & Q_TEXT(i))
                dr.Close()
                
                If temp <> ROW_ID Then
                    ntemp = ntemp + 1
                    If ntemp < 5 Then
                        ROW_ID = "KBA" & Trim(Str(Minute(Now))) & Trim(Str(Second(Now))) & Chr(Str(Int(Rnd() * 26)) + 65) & Chr(Str(Int(Rnd() * 26)) + 65) & Chr(Str(Int(Rnd() * 26)) + 65) & Chr(Str(Int(Rnd() * 26)) + 65)
                        GoTo InsertQuestion
                    End If
                End If
                ntemp = 0
            Next
            
            ' ======================================================
            ' SEND AN UPDATE CONFIRMATION - MESSAGE 0069
            If NUM_QUESTIONS > 0 And EMAIL_ADDR <> "" And KBA_NOTICE = "Y" Then
                SendFrom = "techsupport@gettips.com"
                SendTo = EMAIL_ADDR
                For i = 1 To NUM_QUESTIONS
                    QUESTIONS = QUESTIONS & Q_TEXT(i) & " " & CrLf & " "
                Next
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  QUESTIONS: " & vbCrLf & "  " & QUESTIONS)
                
                MergeData = "<messages>" & _
                    "<message send_to=""" & XmlSafe(SendTo$) & """ send_from=""" & SendFrom$ & """ from_name=""Technical Support"" from_id=""1-6HBOY"" to_id=""" & CONTACT_ID & """>" & _
                    "<FST_NAME>" & XmlSafe(FST_NAME) & "</FST_NAME>" & _
                    "<LAST_NAME>" & XmlSafe(LAST_NAME) & "</LAST_NAME>" & _
                    "<JURIS>" & XmlSafe(JURIS) & "</JURIS>" & _
                    "<COURSE>" & XmlSafe(COURSE) & "</COURSE>" & _
                    "<QUESTIONS>" & XmlSafe(QUESTIONS) & "</QUESTIONS>" & _
                    "<DOMAIN>" & DOMAIN & "</DOMAIN>" & _
                    "</message>" & _
                    "</messages>"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  MergeData: " & vbCrLf & "  " & MergeData)
                results = Processing.EmailXsltMerge(MergeData, "352", "352", "EMAIL", LANG_CD, Debug)
                If Debug = "Y" Then mydebuglog.Debug("   .. results: " & results)
            End If
        Else
            GoTo DBError
        End If
        GoTo CloseOut
        
DBError:
        If Debug = "Y" Then mydebuglog.Debug(">>DBError")
        SUCCESS = "False"
        Select Case LANG_CD
            Case "ESN"
                errmsg = "Ha habido un error del sistema. Int&eacute;ntalo de nuevo later."
            Case Else
                errmsg = "There has been a system error.<br>Please try again later."
        End Select
        GoTo CloseOut
        
DataError:
        ErrLvl = "Warning"
        SUCCESS = "False"
        If Debug = "Y" Then mydebuglog.Debug(">>DataError")
        Select Case LANG_CD
            Case "ESN"
                errmsg = "No podemos guardar sus respuestas debido a un problema con su registro. Por favor cont&aacute;ctenos para asistencia."
            Case Else
                errmsg = "We are unable to save your answers due to a problem with your registration.  Please contact us for assistance."
        End Select
        GoTo CloseOut
        
AccessError:
        ErrLvl = "Warning"
        SUCCESS = "False"
        If Debug = "Y" Then mydebuglog.Debug(">>AccessError")
        Select Case LANG_CD
            Case "ESN"
                errmsg = "No tienes acceso a este registro de clase. Por favor cont&aacute;ctenos para asistencia."
            Case Else
                errmsg = "You do not have access to this class registration.  Please contact us for assistance."
        End Select
        ErrLvl = "Warning"
        
CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            dr = Nothing
            con.Dispose()
            con = Nothing
            cmd.Dispose()
            cmd = Nothing
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the database connection. " & vbCrLf
        End Try
        
CloseOut2:
        ' ============================================
        ' Finalize output
        jdoc = jdoc & """ErrMsg"":""" & errmsg & ""","
        jdoc = jdoc & """Success"":""" & SUCCESS & """"
        jdoc = callback & "({""ResultSet"": {" & jdoc & "} })"
        
        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("WsSaveKBA.ashx : " & ErrLvl & ": " & Trim(errmsg))
        myeventlog.Info("WsSaveKBA.ashx : Success: " & SUCCESS & " for user id: " & REG_NUM & "  and registration id:" & REG_ID)
        If Debug = "Y" Or (Logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("  JDOC: " & jdoc & vbCrLf)
                mydebuglog.Debug("Success: " & SUCCESS & " for user id: " & REG_NUM & "  and registration id:" & REG_ID)
                If Debug = "Y" Then
                    mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                    mydebuglog.Debug("----------------------------------")
                End If
            Catch ex As Exception
            End Try
        End If

        ' Log Performance Data
        If Debug <> "T" Then
            ' ============================================
            ' Send the web request
            Try
                LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, "WSSAVEKBA", LogStartTime, VersionNum, Debug)
            Catch ex As Exception
            End Try
        End If
        
        ' Send results        
        If Debug = "T" Then
            context.Response.ContentType = "text/html"
            If jdoc <> "" Then
                context.Response.Write("Success")
            Else
                context.Response.Write("Failure")
            End If
        Else
            If jdoc = "" Then jdoc = errmsg
            context.Response.ContentType = "application/json"
            context.Response.Write(jdoc)
        End If
        
    End Sub
    
    ' =================================================
    ' JSON FUNCTIONS
    Function DataSetToJSON(ByVal ds As DataSet) As String
    
        Dim json As String
        Dim dt As DataTable = ds.Tables(0)
        json = Newtonsoft.Json.JsonConvert.SerializeObject(dt)
        Return json
        
    End Function
 
    Function EscapeJSON(ByVal todo As String) As String
        If todo = "" Then
            EscapeJSON = ""
            Exit Function
        End If
        todo = Replace(todo, "\", "\\")
        todo = Replace(todo, "/", "\/")
        todo = Replace(todo, """", "\""")
        todo = Replace(todo, Chr(13), "<br>")
        todo = Replace(todo, Chr(10), "<br>")
        todo = Replace(todo, "   ", " ")
        EscapeJSON = todo
    End Function
    
    ' =================================================
    ' STRING FUNCTIONS
    Public Function XmlSafe(ByVal Instring As String) As String
        ' This function encodes the following special characters
        '	", &, ', < and >
        Instring = Replace(Instring, "&", "&amp;")
        Instring = Replace(Instring, "<", "&lt;")
        Instring = Replace(Instring, ">", "&gt;")
        Instring = Replace(Instring, "'", "&apos;")
        Instring = Replace(Instring, """", "&quot;")
        Instring = Replace(Instring, Chr(13), " ")
        Instring = Replace(Instring, Chr(10), " ")
        XmlSafe = Instring
    End Function
    
    Public Function ReverseString(ByVal InputString As String) As String
        ' Reverses a string
        Dim lLen As Long, lCtr As Long
        Dim sChar As String
        Dim sAns As String
        sAns = ""
        lLen = Len(InputString)
        For lCtr = lLen To 1 Step -1
            sChar = Mid(InputString, lCtr, 1)
            sAns = sAns & sChar
        Next
        ReverseString = sAns
    End Function

    Function EmailAddressCheck(ByVal emailAddress As String) As Boolean
        ' Validate email address

        Dim pattern As String = "^[a-zA-Z][\w\.-]*[a-zA-Z0-9]@[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]$"
        Dim emailAddressMatch As Match = Regex.Match(emailAddress, pattern)
        If emailAddressMatch.Success Then
            EmailAddressCheck = True
        Else
            EmailAddressCheck = False
        End If

    End Function

    Function FilterString(ByVal Instring As String) As String
        ' Remove any characters not within the ASCII 31-127 range
        Dim temp As String
        Dim outstring As String
        Dim i, j As Integer

        If Len(Instring) = 0 Or Instring Is Nothing Then
            FilterString = ""
            Exit Function
        End If
        temp = Instring.ToString
        outstring = ""
        For i = 1 To Len(temp$)
            j = Asc(Mid(temp, i, 1))
            If j > 30 And j < 128 Then
                outstring = outstring & Mid(temp, i, 1)
            End If
        Next
        FilterString = outstring
    End Function
    Function SqlString(ByVal Instring As String) As String
        ' Make a string safe for use in a SQL query
        Dim temp As String
        Dim outstring As String
        Dim i As Integer

        If Len(Instring) = 0 Or Instring Is Nothing Then
            SqlString = ""
            Exit Function
        End If
        temp = Instring.ToString
        outstring = ""
        For i = 1 To Len(temp$)
            If Mid(temp, i, 1) = "'" Then
                outstring = outstring & "''"
            Else
                outstring = outstring & Mid(temp, i, 1)
            End If
        Next
        SqlString = outstring
    End Function

    Function CheckNull(ByVal Instring As String) As String
        ' Check to see if a string is null
        If Instring Is Nothing Then
            CheckNull = ""
        Else
            CheckNull = Instring
        End If
    End Function

    Public Function CheckDBNull(ByVal obj As Object, _
    Optional ByVal ObjectType As enumObjectType = enumObjectType.StrType) As Object
        ' Checks an object to determine if its null, and if so sets it to a not-null empty value
        Dim objReturn As Object
        objReturn = obj
        If ObjectType = enumObjectType.StrType And IsDBNull(obj) Then
            objReturn = ""
        ElseIf ObjectType = enumObjectType.IntType And IsDBNull(obj) Then
            objReturn = 0
        ElseIf ObjectType = enumObjectType.DblType And IsDBNull(obj) Then
            objReturn = 0.0
        ElseIf ObjectType = enumObjectType.DteType And IsDBNull(obj) Then
            objReturn = Now
        End If
        Return objReturn
    End Function

    Public Function NumString(ByVal strString As String) As String
        ' Remove everything but numbers from a string
        Dim bln As Boolean
        Dim i As Integer
        Dim iv As String
        NumString = ""

        'Can array element be evaluated as a number?
        For i = 1 To Len(strString)
            iv = Mid(strString, i, 1)
            bln = IsNumeric(iv)
            If bln Then NumString = NumString & iv
        Next

    End Function

    Public Function ToBase64(ByVal data() As Byte) As String
        ' Encode a Base64 string
        If data Is Nothing Then Throw New ArgumentNullException("data")
        Return Convert.ToBase64String(data)
    End Function

    Public Function FromBase64(ByVal base64 As String) As String
        ' Decode a Base64 string
        Dim results As String
        If base64 Is Nothing Then Throw New ArgumentNullException("base64")
        results = System.Text.Encoding.ASCII.GetString(Convert.FromBase64String(base64))
        Return results
    End Function

    Function DeSqlString(ByVal Instring As String) As String
        ' Convert a string from SQL query encoded to non-encoded
        Dim temp As String
        Dim outstring As String
        Dim i As Integer

        CheckDBNull(Instring, enumObjectType.StrType)
        If Len(Instring) = 0 Then
            DeSqlString = ""
            Exit Function
        End If
        temp = Instring.ToString
        outstring = ""
        For i = 1 To Len(temp$)
            If Mid(temp, i, 2) = "''" Then
                outstring = outstring & "'"
                i = i + 1
            Else
                outstring = outstring & Mid(temp, i, 1)
            End If
        Next
        DeSqlString = outstring
    End Function

    Public Function StringToBytes(ByVal str As String) As Byte()
        ' Convert a random string to a byte array
        ' e.g. "abcdefg" to {a,b,c,d,e,f,g}
        Dim s As Char()
        Dim t As Char
        s = str.ToCharArray
        Dim b(s.Length - 1) As Byte
        Dim i As Integer
        For i = 0 To s.Length - 1
            If Asc(s(i)) < 128 And Asc(s(i)) > 0 Then
                Try
                    b(i) = Convert.ToByte(s(i))
                Catch ex As Exception
                    b(i) = Convert.ToByte(Chr(32))
                End Try
            Else
                ' Filter out extended ASCII - convert common symbols when possible
                t = Chr(32)
                Try
                    Select Case Asc(s(i))
                        Case 147
                            t = Chr(34)
                        Case 148
                            t = Chr(34)
                        Case 145
                            t = Chr(39)
                        Case 146
                            t = Chr(39)
                        Case 150
                            t = Chr(45)
                        Case 151
                            t = Chr(45)
                        Case Else
                            t = Chr(32)
                    End Select
                Catch ex As Exception
                End Try
                b(i) = Convert.ToByte(t)
            End If
        Next
        Return b
    End Function

    Public Function EncodeParamSpaces(ByVal InVal As String) As String
        ' If given a urlencoded parameter value, replace spaces with "+" signs

        Dim temp As String
        Dim i As Integer

        If InStr(InVal, " ") > 0 Then
            temp = ""
            For i = 1 To Len(InVal)
                If Mid(InVal, i, 1) = " " Then
                    temp = temp & "+"
                Else
                    temp = temp & Mid(InVal, i, 1)
                End If
            Next
            EncodeParamSpaces = temp
        Else
            EncodeParamSpaces = InVal
        End If
    End Function

    Public Function DecodeParamSpaces(ByVal InVal As String) As String
        ' If given an encoded parameter value, replace "+" signs with spaces

        Dim temp As String
        Dim i As Integer

        If InStr(InVal, "+") > 0 Then
            temp = ""
            For i = 1 To Len(InVal)
                If Mid(InVal, i, 1) = "+" Then
                    temp = temp & " "
                Else
                    temp = temp & Mid(InVal, i, 1)
                End If
            Next
            DecodeParamSpaces = temp
        Else
            DecodeParamSpaces = InVal
        End If
    End Function

    Public Function NumStringToBytes(ByVal str As String) As Byte()
        ' Convert a string containing numbers to a byte array
        ' e.g. "1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16" to 
        '  {1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16}
        Dim s As String()
        s = str.Split(" ")
        Dim b(s.Length - 1) As Byte
        Dim i As Integer
        For i = 0 To s.Length - 1
            b(i) = Convert.ToByte(s(i))
        Next
        Return b
    End Function

    Public Function BytesToString(ByVal b() As Byte) As String
        ' Convert a byte array to a string
        Dim i As Integer
        Dim s As New System.Text.StringBuilder()
        For i = 0 To b.Length - 1
            Console.WriteLine(b(i))
            If i <> b.Length - 1 Then
                s.Append(b(i) & " ")
            Else
                s.Append(b(i))
            End If
        Next
        Return s.ToString
    End Function
    
    ' =================================================
    ' DATABASE FUNCTIONS
    Public Function OpenDBConnection(ByVal ConnS As String, ByRef con As SqlConnection, ByRef cmd As SqlCommand) As String
        ' Function to open a database connection with extreme error-handling
        ' Returns an error message if unable to open the connection
        Dim SqlS As String
        SqlS = ""
        OpenDBConnection = ""

        Try
            con = New SqlConnection(ConnS)
            con.Open()
            If Not con Is Nothing Then
                Try
                    cmd = New SqlCommand(SqlS, con)
                    cmd.CommandTimeout = 300
                Catch ex2 As Exception
                    OpenDBConnection = "Error opening the command string: " & ex2.ToString
                End Try
            End If
        Catch ex As Exception
            If con.State <> Data.ConnectionState.Closed Then con.Dispose()
            ConnS = ConnS & ";Pooling=false"
            Try
                con = New SqlConnection(ConnS)
                con.Open()
                If Not con Is Nothing Then
                    Try
                        cmd = New SqlCommand(SqlS, con)
                        cmd.CommandTimeout = 300
                    Catch ex2 As Exception
                        OpenDBConnection = "Error opening the command string: " & ex2.ToString
                    End Try
                End If
            Catch ex2 As Exception
                OpenDBConnection = "Unable to open database connection for connection string: " & ConnS & vbCrLf & "Windows error: " & vbCrLf & ex2.ToString & vbCrLf
            End Try
        End Try

    End Function
    
    Public Function CloseDBConnection(ByRef con As SqlConnection, ByRef cmd As SqlCommand, ByRef dr As SqlDataReader) As String
        ' This function closes a database connection safely
        Dim ErrMsg As String
        ErrMsg = ""

        ' Handle datareader
        Try
            dr.Close()
        Catch ex As Exception
        End Try
        Try
            dr = Nothing
        Catch ex As Exception
        End Try

        ' Handle command
        Try
            cmd.Dispose()
        Catch ex As Exception
        End Try
        Try
            cmd = Nothing
        Catch ex As Exception
        End Try

        ' Handle connection
        Try
            con.Close()
        Catch ex As Exception
        End Try
        Try
            SqlConnection.ClearPool(con)
        Catch ex As Exception
        End Try
        Try
            con.Dispose()
        Catch ex As Exception
        End Try
        Try
            con = Nothing
        Catch ex As Exception
        End Try

        ' Exit
        Return ErrMsg
    End Function
    
    Public Function ExecQuery(ByVal QType As String, ByVal QRec As String, ByVal cmd As SqlCommand, ByVal SqlS As String, ByVal mydebuglog As ILog, ByVal Debug As String) As String
        Dim returnv As Integer
        Dim errmsg As String
        errmsg = ""
        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  " & QType & " " & QRec & " record: " & SqlS)
        Try
            cmd.CommandText = SqlS
            returnv = cmd.ExecuteNonQuery()
            If returnv = 0 Then
                errmsg = errmsg & "The " & QRec & " record was not " & QType & vbCrLf
            End If
        Catch ex As Exception
            errmsg = errmsg & "Error " & QType & " record. " & ex.ToString & vbCrLf & "Query: " & SqlS
        End Try
        Return errmsg
    End Function

    ' =================================================
    ' DEBUG FUNCTIONS
    Public Sub writeoutputfs(ByVal fs As FileStream, ByVal instring As String)
        ' This function writes a line to a previously opened filestream, and then flushes it
        ' promptly.  This assists in debugging services
        fs.Write(StringToBytes(instring), 0, Len(instring))
        fs.Write(StringToBytes(vbCrLf), 0, 2)
        fs.Flush()
    End Sub
        
    Public ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class