<%@ WebHandler Language="VB" Class="WsAcceptClass" %>

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

Public Class WsAcceptClass : Implements IHttpHandler
    
    Enum enumObjectType
        StrType = 0
        IntType = 1
        DblType = 2
        DteType = 3
    End Enum
    
    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        
        ' Parameter Declarations
        Dim Debug, temp, temp2 As String
        
        ' Result Declarations
        Dim jdoc As String
        
        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String

        ' Logging declarations
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("AcceptClassDebugLog")
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
        Dim cookieid As String = Trim(context.Request.Cookies.Item("ID").Value.ToString())
        Dim SessID As String = Trim(context.Request.Cookies.Item("Sess").Value.ToString())
        
        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service
        
        ' Variable declarations
        Dim errmsg, ErrLvl As String
        Dim REG_ID, HOME_PAGE, LANG_CD, UID, callback, DOMAIN As String
        Dim LOGGED_IN, CONTACT_ID, SUB_ID, CONTACT_OU_ID As String
        Dim ENTER_FLG, REG_STATUS_CD, PASSED_FLG, FST_NAME, LAST_NAME, CRSE_NAME, COMPLETE_DATE, CON_ID, SESS_PART_ID As String
        Dim EMAIL_ADDR, USERNAME, PASSWORD, SCORM, CRSE_ID, TEST_STATUS, CRSE_TSTRUN_ID, EXAM_REQD, EXAM_ENGINE As String
        Dim JURIS_ID, LANG_ID, USER_ID, CRSE_TST_ID, SHELL_TYPE, CRSE_CONTENT_URL, RETAKE_FLG, ALT_RETAKE_FLG As String
        Dim LaunchProtocol As String
        Dim DupeCall, InsertExit, MayTest, CompletedCourse, OldScorm As Boolean
        Dim NumPts, StartP, EndP As Integer
        Dim completion_status, suspend_data, current_location As String
        Dim High_Water, Progress, Results As String
        Dim User_Loc As Integer
        
        ' ============================================
        ' Variable setup
        Debug = "N"
        Logging = "Y"
        errmsg = ""
        Results = "Failure"
        LOGGED_IN = "N"
        CONTACT_ID = ""
        SUB_ID = ""
        CONTACT_OU_ID = ""
        LANG_CD = "ENU"
        callback = ""
        DOMAIN = "TIPS"
        HOME_PAGE = ""
        REG_ID = ""
        UID = ""
        SHELL_TYPE = "FLASH"
        LaunchProtocol = "http:"
        suspend_data = ""
        completion_status = ""
        current_location = ""
        CompletedCourse = False
        OldScorm = False
        Progress = ""
        User_Loc = 0
        CRSE_NAME = ""
        REG_STATUS_CD = ""
        SESS_PART_ID = ""
        REG_STATUS_CD = ""
        REG_ID = ""
        CRSE_TSTRUN_ID = ""
        CRSE_TST_ID = ""
        CRSE_CONTENT_URL = ""
        SCORM = ""
        LANG_ID = ""
        EXAM_REQD = ""
        SHELL_TYPE = ""
        RETAKE_FLG = ""
        ALT_RETAKE_FLG = ""
        CRSE_ID = ""
        EXAM_ENGINE = ""
        USER_ID = ""
        ErrLvl = "Error"
        
        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server=DB_SERVER;uid=DB_USER;pwd=DB_PASSWORD;database=siebeldb"
            tempdebug = System.Configuration.ConfigurationManager.AppSettings.Get("AcceptClass_debug")
            If tempdebug = "Y" And Debug <> "T" Then Debug = "Y"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            GoTo CloseOut2
        End Try
        
        ' ============================================
        ' Open log file if applicable
        If Logging = "Y" Then
            logfile = "C:\Logs\WsAcceptClass.log"
            Try
                log4net.GlobalContext.Properties("AcceptClassLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                GoTo CloseOut2
            End Try
        End If
        
        ' ============================================
        ' Get parameters    
        '   REG_ID          - The registration id of the course to start (CX_SESS_REG.ROW_ID)
        '   UID	            - The user's S_CONTACT.X_REGISTRATION_NUM

        If Not context.Request.QueryString("ID") Is Nothing Then
            REG_ID = context.Request.QueryString("ID")
        End If

        If Not context.Request.QueryString("USR") Is Nothing Then
            UID = context.Request.QueryString("USR")
        End If
                
        If Debug = "Y" Then
            mydebuglog.Debug("----------------------------------")
            mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
            mydebuglog.Debug("Parameters-")
            mydebuglog.Debug("  REG_ID: " & REG_ID)
            mydebuglog.Debug("  UID: " & UID)
            mydebuglog.Debug("  cookieid: " & cookieid)
        End If
           
        ' Validate parameters
        If REG_ID = "" Then GoTo DataError
        If UID = "" Then GoTo AccessError
        If cookieid <> UID Then GoTo AccessError
        
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
            ' GET THE RESULTS OF THE CLASS AS REPORTED IN CX_SESS_REG
            SqlS = "SELECT R.STATUS_CD, T.DOMAIN, T.CRSE_CONTENT_URL, SP.PASS_FLG, C.FST_NAME, C.LAST_NAME, " & _
            "CR.NAME, FORMAT(SP.CREATED,'M/d/yyyy'), C.ROW_ID, SP.ROW_ID, C.EMAIL_ADDR, C.LOGIN, C.X_PASSWORD, CR.X_SCORM_FLG, " & _
            "CR.ROW_ID, CT.STATUS_CD, CT.ROW_ID, CR.X_EXAM_REQD, E.X_ENGINE, R.JURIS_ID, " & _
            "(SELECT CASE WHEN T.LANG_ID IS NULL THEN CR.X_LANG_CD ELSE T.LANG_ID END) AS LANG_ID, " & _
            "C.X_REGISTRATION_NUM, CT.CRSE_TST_ID, CR.X_FORMAT, CR.X_CRSE_CONTENT_URL, R.RETAKE_FLG, CR.X_ALT_RETAKE_FLG " & _
            "FROM siebeldb.dbo.CX_SESS_REG R " & _
            "LEFT OUTER JOIN siebeldb.dbo.CX_TRAIN_OFFR T ON T.ROW_ID=R.TRAIN_OFFR_ID " & _
            "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=R.CONTACT_ID " & _
            "LEFT OUTER JOIN siebeldb.dbo.CX_SESS_PART_X SP ON SP.ROW_ID=R.SESS_PART_ID " & _
            "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TSTRUN CT ON CT.ROW_ID=SP.CRSE_TSTRUN_ID " & _
            "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TST E ON E.ROW_ID=CT.CRSE_TST_ID " & _
            "LEFT OUTER JOIN siebeldb.dbo.S_CRSE CR ON CR.ROW_ID=R.CRSE_ID " & _
            "WHERE R.ROW_ID='" & REG_ID & "' AND C.X_REGISTRATION_NUM='" & UID & "'"
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  RETRIEVE CLASS INFORMATION: " & vbCrLf & "  " & SqlS)
            cmd.CommandText = SqlS
            dr = cmd.ExecuteReader()
            If Not dr Is Nothing Then
                While dr.Read()
                    Try
                        REG_STATUS_CD = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                        DOMAIN = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                        CRSE_CONTENT_URL = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                        PASSED_FLG = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                        FST_NAME = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                        LAST_NAME = Trim(CheckDBNull(dr(5), enumObjectType.StrType))
                        CRSE_NAME = Trim(CheckDBNull(dr(6), enumObjectType.StrType))
                        COMPLETE_DATE = Trim(CheckDBNull(dr(7), enumObjectType.StrType))
                        CON_ID = Trim(CheckDBNull(dr(8), enumObjectType.StrType))
                        SESS_PART_ID = Trim(CheckDBNull(dr(9), enumObjectType.StrType))
                        EMAIL_ADDR = Trim(CheckDBNull(dr(10), enumObjectType.StrType))
                        USERNAME = Trim(CheckDBNull(dr(11), enumObjectType.StrType))
                        PASSWORD = Trim(CheckDBNull(dr(12), enumObjectType.StrType))
                        SCORM = Trim(CheckDBNull(dr(13), enumObjectType.StrType))
                        CRSE_ID = Trim(CheckDBNull(dr(14), enumObjectType.StrType))
                        TEST_STATUS = Trim(CheckDBNull(dr(15), enumObjectType.StrType))
                        CRSE_TSTRUN_ID = Trim(CheckDBNull(dr(16), enumObjectType.StrType))
                        EXAM_REQD = Trim(CheckDBNull(dr(17), enumObjectType.StrType))
                        EXAM_ENGINE = Trim(CheckDBNull(dr(18), enumObjectType.StrType))
                        JURIS_ID = Trim(CheckDBNull(dr(19), enumObjectType.StrType))
                        LANG_ID = Trim(CheckDBNull(dr(20), enumObjectType.StrType))
                        USER_ID = Trim(CheckDBNull(dr(21), enumObjectType.StrType))
                        CRSE_TST_ID = Trim(CheckDBNull(dr(22), enumObjectType.StrType))
                        SHELL_TYPE = Trim(CheckDBNull(dr(23), enumObjectType.StrType))
                        If InStr(CRSE_CONTENT_URL, "https:") > 0 Then
                            LaunchProtocol = "https:"
                        End If
                        RETAKE_FLG = Trim(CheckDBNull(dr(25), enumObjectType.StrType))
                        ALT_RETAKE_FLG = Trim(CheckDBNull(dr(26), enumObjectType.StrType))
                    Catch ex As Exception
                        GoTo DBError
                    End Try
                End While
            Else
                GoTo DataError
            End If
            dr.Close()
            If REG_STATUS_CD = "" Then GoTo AccessError
            If Debug = "Y" Then
                mydebuglog.Debug("  .. REG_STATUS_CD : " & REG_STATUS_CD)
                mydebuglog.Debug("  .. CRSE_ID : " & CRSE_ID)
                mydebuglog.Debug("  .. USER_ID : " & USER_ID)
                mydebuglog.Debug("  .. CRSE_TSTRUN_ID : " & CRSE_TSTRUN_ID)
                mydebuglog.Debug("  .. CRSE_TST_ID : " & CRSE_TST_ID)
                mydebuglog.Debug("  .. CRSE_CONTENT_URL : " & CRSE_CONTENT_URL)
                mydebuglog.Debug("  .. SCORM : " & SCORM)
                mydebuglog.Debug("  .. LANG_ID : " & LANG_ID)
                mydebuglog.Debug("  .. EXAM_REQD : " & EXAM_REQD)
                mydebuglog.Debug("  .. SHELL_TYPE : " & SHELL_TYPE)
                mydebuglog.Debug("  .. RETAKE_FLG : " & RETAKE_FLG)
                mydebuglog.Debug("  .. ALT_RETAKE_FLG : " & ALT_RETAKE_FLG)
                mydebuglog.Debug("  .. LaunchProtocol : " & LaunchProtocol)
                mydebuglog.Debug("  .. EXAM_ENGINE : " & EXAM_ENGINE)
                mydebuglog.Debug("  .. CRSE_TSTRUN_ID : " & CRSE_TSTRUN_ID)
            End If
            If UID <> USER_ID Then GoTo AccessError
            
            ' ================================================
            ' CHECK PROGRESS FOR COURSES WITH SCORM ASSESSMENTS
            temp = ""
            If SCORM = "Y" Then
                ' Get stored progress data
                SqlS = "SELECT CAST(suspend_data AS VARCHAR(8000)), completion_status, location " & _
                "FROM elearning.dbo.Elearning_Player_Data " & _
                "WHERE reg_id='" & REG_ID & "'"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Query for SCORM data: " & vbCrLf & "  " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            suspend_data = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            completion_status = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            current_location = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                        Catch ex As Exception
                            GoTo DBError
                        End Try
                    End While
                End If
                dr.Close()
                If Debug = "Y" Then
                    mydebuglog.Debug("  .. suspend_data : " & suspend_data)
                    mydebuglog.Debug("  .. completion_status : " & completion_status)
                    mydebuglog.Debug("  .. current_location : " & current_location)
                End If
                
                ' Verify status received versus stored in the database
                Select Case completion_status
                    Case "1"
                        completion_status = "In Progress"
                    Case "2"
                        If EXAM_REQD = "Y" Then completion_status = "Exam Reqd" Else completion_status = "Completed"
                        CompletedCourse = True
                    Case "3"
                        completion_status = "In Progress"
                    Case "4"
                        completion_status = "In Progress"
                    Case "5"
                        completion_status = "In Progress"
                    Case Else
                        completion_status = "Failure"
                End Select
                If Debug = "Y" Then
                    mydebuglog.Debug("  .. scorm completion_status : " & completion_status)
                    mydebuglog.Debug("  .. REG_STATUS_CD : " & REG_STATUS_CD)
                End If
      
                ' Procession suspend_data based on shell type
                temp2 = ""

                Select Case SHELL_TYPE
                    Case "HTML5"
                        ' Decrypt shell suspend_data if applicable
                        temp = suspend_data
                        StartP = InStr(temp, "<paging>") + 8
                        EndP = InStr(temp, "</paging>") - StartP
                        If StartP > 0 And EndP > 0 Then
                            temp2 = Mid(temp, StartP, EndP)
                            temp2 = Replace(temp2, ",", "")
                            NumPts = Len(temp2)
                        End If
                        If Debug = "Y" Then
                            mydebuglog.Debug("  .. paging data : " & temp2)
                            mydebuglog.Debug("  .. Number screens in course : " & Str(NumPts))
                        End If

                        ' Check to see if on last page if IN PROGRESS status
                        Select Case REG_STATUS_CD
                            Case "In Progress"
                                If Right(temp2, 1) = "1" And Val(current_location) < (NumPts - 1) Then MayTest = True
                            Case "Exam Reqd"
                                If Right(temp2, 1) = "1" And Val(current_location) < (NumPts - 1) Then MayTest = True
                                CompletedCourse = True
                            Case Else
                                MayTest = False
                        End Select
                    Case "FLASH"
                        ' Check to see if on last page if IN PROGRESS status
                        If REG_STATUS_CD = "In Progress" Then
                            If EXAM_ENGINE = "SCORM" Then
                                If InStr(suspend_data, "1</pageCompletion>") > 0 Then MayTest = True
                            End If
                        End If
                End Select
                
                ' Reset status set by AcceptRollup if the user is not on the last screen when they exit.  That service
                ' sets the status to "Exam Reqd", while the user might actually want to continue the course
                If MayTest Then
                    If Not DupeCall Then
                        SqlS = "UPDATE siebeldb.dbo.CX_SESS_REG " & _
                            "SET STATUS_CD='In Progress' " & _
                            "WHERE ROW_ID='" & REG_ID & "'"
                        temp = ExecQuery("Update", "CX_SESS_REG", cmd, SqlS, mydebuglog, Debug)
                    End If
                    REG_STATUS_CD = "In Progress"
                End If
            
            Else
                If EXAM_ENGINE = "SCORM" And (REG_STATUS_CD = "Accepted" Or REG_STATUS_CD = "In Progress") Then
                    ' Get number of points in course
                    SqlS = "SELECT NUM_ELEMENTS, XML_CRSE_ID FROM elearning.dbo.ELN_COURSE_MAP WHERE CRSE_ID='" & CRSE_ID & "'"
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  NUM_ELEMENTS QUERY: " & vbCrLf & "  " & SqlS)
                    cmd.CommandText = SqlS
                    dr = cmd.ExecuteReader()
                    If Not dr Is Nothing Then
                        While dr.Read()
                            Try
                                NumPts = CheckDBNull(dr(0), enumObjectType.IntType)
                                temp = Trim(CheckDBNull(dr(1), enumObjectType.StrType))                ' XML Course Id
                            Catch ex As Exception
                                GoTo DBError
                            End Try
                        End While
                    End If
                    dr.Close()
                    If Debug = "Y" Then
                        mydebuglog.Debug("  .. XML Course Id : " & temp)
                        mydebuglog.Debug("  .. Number screens in course : " & Str(NumPts))
                    End If
                    
                    ' Get high water mark
                    High_Water = "1"
                    SqlS = "SELECT TOP 1 HIGH_WATER FROM elearning.dbo.ELN_USER_PROGRESS WHERE SESS_REG_ID='" & REG_ID & "' ORDER BY START_DATE DESC, HIGH_WATER DESC"
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Get SCORM package id: " & vbCrLf & "  " & SqlS)
                    cmd.CommandText = SqlS
                    Try
                        High_Water = CheckDBNull(cmd.ExecuteScalar(), enumObjectType.StrType)
                        If High_Water = "" Or High_Water = "0" Or High_Water = "False" Then High_Water = "1"
                    Catch ex As Exception
                        GoTo AccessError
                    End Try
                    If Debug = "Y" Then mydebuglog.Debug("  .. High_Water : " & High_Water)

                    If High_Water = "1" Then
                        User_Loc = 1
                    Else
                        SqlS = "SELECT COUNT(*) " & _
                            "FROM elearning.dbo.ELN_COURSE_ELEMENT E " & _
                            "LEFT OUTER JOIN elearning.dbo.ELN_COURSE_ELEMENT_META M ON M.id=E.id AND E.id is not null " & _
                            "WHERE M.course_id is not null AND M.course_id='" & temp & "' and E.type='G' AND E.high_water<='" & High_Water & "'  AND E.high_water>=0"
                        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  COUNT ELEMENTS QUERY: " & vbCrLf & "  " & SqlS)
                        cmd.CommandText = SqlS
                        Try
                            User_Loc = CheckDBNull(cmd.ExecuteScalar(), enumObjectType.IntType)
                        Catch ex As Exception
                            GoTo AccessError
                        End Try
                    End If
                    If Debug = "Y" Then mydebuglog.Debug("  .. User_Loc : " & Str(User_Loc))
                End If

                ' If the user is at the end, set the status accordingly
                If User_Loc >= NumPts Then MayTest = True

                ' If status Accepted then change to In Progress if applicable
                If REG_STATUS_CD = "Accepted" And User_Loc > 1 Then
                    REG_STATUS_CD = "In Progress"
                    If Not DupeCall Then
                        SqlS = "UPDATE siebeldb.dbo.CX_SESS_REG " & _
                           "SET STATUS_CD='In Progress' " & _
                           "WHERE ROW_ID='" & REG_ID & "'"
                        temp = ExecQuery("Update", "CX_SESS_REG", cmd, SqlS, mydebuglog, Debug)
                    End If
                End If
                
                ' Update progress field
                If NumPts > 0 And User_Loc > 0 Then
                    If NumPts >= User_Loc Then
                        Progress = Str((User_Loc / NumPts) * 100)
                    Else
                        Progress = "100.00"
                    End If
                End If
                If Progress <> "" Then
                    If Not DupeCall Then
                        SqlS = "UPDATE siebeldb.dbo.CX_SESS_REG " & _
                           "SET PROGRESS='" & Progress & "' WHERE ROW_ID='" & REG_ID & "'"
                        temp = ExecQuery("Update", "CX_SESS_REG", cmd, SqlS, mydebuglog, Debug)
                    End If
                End If
            End If
            If Debug = "Y" Then mydebuglog.Debug("  .. MayTest : " & MayTest)
            
            ' ================================================
            ' LOG THE EXIT FROM THE COURSE IF NOT ALREADY LOGGED
            ' Should not be necessary if AcceptRollup service ran correctly
            ENTER_FLG = ""
            InsertExit = False
            SqlS = "SELECT TOP 1 ENTER_FLG FROM siebeldb.dbo.CX_TRAIN_OFFR_ACCESS WHERE REG_ID='" & REG_ID & "' ORDER BY CREATED DESC"
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  CHECK ACCESS SEMAPHORE: " & vbCrLf & "  " & SqlS)
            cmd.CommandText = SqlS
            Try
                ENTER_FLG = CheckDBNull(cmd.ExecuteScalar(), enumObjectType.StrType)
            Catch ex As Exception
                InsertExit = True
            End Try
            If ENTER_FLG = "Y" Or ENTER_FLG = "" Then InsertExit = True
            If Debug = "Y" Then mydebuglog.Debug("  .. InsertExit: " & InsertExit)
            If InsertExit Then
                SqlS = "INSERT INTO siebeldb.dbo.CX_TRAIN_OFFR_ACCESS (ROW_ID, CREATED, CREATED_BY, LAST_UPD, LAST_UPD_BY, " & _
                "MODIFICATION_NUM, CONFLICT_ID, REG_ID, ENTER_FLG, EXIT_FLG, MOBILE) " & _
                "SELECT '" & REG_ID & "-'+LTRIM(CAST(COUNT(*)+1 AS VARCHAR)), GETDATE(), '0-1', GETDATE(), '0-1', " & _
                "0, 0, '" & REG_ID & "','N','Y','Y' " & _
                "FROM siebeldb.dbo.CX_TRAIN_OFFR_ACCESS " & _
                "WHERE REG_ID = '" & REG_ID & "'"
                temp = ExecQuery("Insert", "CX_TRAIN_OFFR_ACCESS", cmd, SqlS, mydebuglog, Debug)
            End If
                        
            ' Log to CM activity log
            SqlS = "INSERT INTO reports.dbo.CM_LOG(REG_ID, SESSION_ID, RECORD_ID, REMOTE_ADDR, ACTION, BROWSER, RECORD_DATA) " & _
                "VALUES('" & UID & "','" & SessID & "','" & REG_ID & "','" & callip & "','EXITED COURSE','" & Left(BROWSER, 200) & "','" & Replace(suspend_data, "'", "''") & "')"
            temp = ExecQuery("Insert", "CM_LOG", cmd, SqlS, mydebuglog, Debug)
            Results = "Success"
        Else
            GoTo DBError
        End If
        GoTo CloseOut

DBError:
        If Debug = "Y" Then mydebuglog.Debug(">>DBError")
        Select Case LANG_CD
            Case "ESN"
                errmsg = "Ha habido un error del sistema. Int&eacute;ntalo de nuevo later."
            Case Else
                errmsg = "There has been a system error.<br>Please try again later."
        End Select
        GoTo CloseOut
        
DataError:
        ErrLvl = "Warning"
        If Debug = "Y" Then mydebuglog.Debug(">>DataError")
        Select Case LANG_CD
            Case "ESN"
                errmsg = "No podemos abrir su clase debido a un problema con su registro. Por favor cont&aacute;ctenos para asistencia."
            Case Else
                errmsg = "We are unable to open your class due to a problem with your registration.  Please contact us for assistance."
        End Select
        GoTo CloseOut
        
AccessError:
        ErrLvl = "Warning"
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
        jdoc = """ErrMsg"":""" & errmsg & ""","
        jdoc = jdoc & """Results"":""" & Results & ""","
        jdoc = callback & "({""ResultSet"": {" & jdoc & "} })"
        
        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("WsAcceptClass.ashx : " & ErrLvl & ": " & Trim(errmsg))
        myeventlog.Info("WsAcceptClass.ashx : UID: " & UID & ", SessID: " & SessID & ", and RegID:" & REG_ID)
        If Debug = "Y" Or (Logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("  JDOC: " & jdoc & vbCrLf)
                mydebuglog.Debug("Results:  UID: " & UID & ", SessID: " & SessID & ", and RegID:" & REG_ID)
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
                LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, "WSACCEPTCLASS", LogStartTime, VersionNum, Debug)
            Catch ex As Exception
            End Try
        End If
        
        ' Send results        
        If jdoc = "" Then jdoc = errmsg
        context.Response.ContentType = "application/json"
        context.Response.Write(jdoc)
        
    End Sub
    
    ' =================================================d
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