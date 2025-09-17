<%@ WebHandler Language="VB" Class="WsLeaveClass" %>

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

Public Class WsLeaveClass : Implements IHttpHandler
    
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
        mydebuglog = log4net.LogManager.GetLogger("LeaveClassDebugLog")
        Dim logfile, tempdebug As String
        Dim Logging As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "101"

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
        
        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service
        
        ' Variable declarations
        Dim errmsg, ErrLvl As String
        Dim REG_ID, REG_NUM, HOME_PAGE, LANG_CD, UID, SessID, CURRENT_PAGE, callback, DOMAIN As String
        Dim LOGGED_IN, CONTACT_ID, SUB_ID, CONTACT_OU_ID As String
        Dim ENTER_FLG, REG_STATUS_CD, PASSED_FLG, FST_NAME, LAST_NAME, CRSE_NAME, COMPLETE_DATE, CON_ID, SESS_PART_ID As String
        Dim EMAIL_ADDR, USERNAME, PASSWORD, SCORM, CRSE_ID, TEST_STATUS, CRSE_TSTRUN_ID, EXAM_REQD, EXAM_ENGINE As String
        Dim JURIS_ID, LANG_ID, USER_ID, CRSE_TST_ID, SHELL_TYPE, CRSE_CONTENT_URL, RETAKE_FLG, ALT_RETAKE_FLG As String
        Dim MSG1, redirect, NextLink, ExamLink, ClassLink, Refresh, LaunchProtocol As String
        Dim DupeCall, InsertExit, MayTest, CompletedCourse As Boolean
        Dim NumPts, StartP, EndP As Integer
        Dim completion_status, suspend_data, current_location As String
        Dim High_Water, Progress As String
        Dim User_Loc As Integer
        
        ' ============================================
        ' Variable setup
        Debug = "N"
        Logging = "Y"
        errmsg = ""
        LOGGED_IN = "N"
        CONTACT_ID = ""
        SUB_ID = ""
        CONTACT_OU_ID = ""
        LANG_CD = "ENU"
        callback = ""
        DOMAIN = "TIPS"
        HOME_PAGE = ""
        REG_ID = ""
        REG_NUM = ""
        UID = ""
        SessID = ""
        CURRENT_PAGE = ""
        ErrLvl = "Error"
        DupeCall = False
        MSG1 = ""
        redirect = ""
        NextLink = ""
        ExamLink = ""
        ClassLink = ""
        Refresh = ""
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
        LaunchProtocol = "http:"
        CRSE_ID = ""
        EXAM_ENGINE = ""
        JURIS_ID = ""
        MayTest = False
        suspend_data = ""
        completion_status = ""
        current_location = ""
        CompletedCourse = False
        Progress = ""
        User_Loc = 0
        CRSE_NAME = ""
        
        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server=DB_SERVER;uid=DB_USER;pwd=DB_PASSWORD;database=siebeldb"
            tempdebug = System.Configuration.ConfigurationManager.AppSettings.Get("LeaveClass_debug")
            If tempdebug = "Y" And Debug <> "T" Then Debug = "Y"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            GoTo CloseOut2
        End Try
        
        ' ============================================
        ' Open log file if applicable
        If Logging = "Y" Then
            logfile = "C:\Logs\WsLeaveClass.log"
            Try
                log4net.GlobalContext.Properties("LeaveClassLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                GoTo CloseOut2
            End Try
        End If
        
        ' ============================================
        ' Get parameters    
        '   REG_ID          - The registration id of the course to start (CX_SESS_REG.ROW_ID)
        '   REG_NUM         - The user's S_CONTACT.X_REGISTRATION_NUM
        '   SessID          - The user's session id
        '   CURRENT_PAGE    - Optional. The page from which this service was called
        '   HOME_PAGE       - Optional. The root domain that invoked this service
        '   Callback        - The name of the Javascript callback function in which to wrap the resulting JSON . Default freightCalc
        '   DOMAIN          - The student's domain      
        '   LANG_CD         - The language code

        If Not context.Request.QueryString("RID") Is Nothing Then
            REG_ID = context.Request.QueryString("RID")
        End If

        If Not context.Request.QueryString("UID") Is Nothing Then
            REG_NUM = context.Request.QueryString("UID")
        End If
        
        If Not context.Request.QueryString("SES") Is Nothing Then
            SessID = context.Request.QueryString("SES")
        End If

        If Not context.Request.QueryString("CUR") Is Nothing Then
            CURRENT_PAGE = context.Request.QueryString("CUR")
        End If

        If Not context.Request.QueryString("HP") Is Nothing Then
            HOME_PAGE = context.Request.QueryString("HP")
        End If
        
        If Not context.Request.QueryString("PP") Is Nothing Then
            DOMAIN = UCase(context.Request.QueryString("PP"))
        End If
        
        If Not context.Request.QueryString("LANG") Is Nothing Then
            LANG_CD = UCase(context.Request.QueryString("LANG"))
        End If
        
        If Not context.Request.QueryString("callback") Is Nothing Then
            callback = context.Request.QueryString("callback")
        End If
        
        ' Validate parameters
        If Left(CURRENT_PAGE, 2) = "//" Then CURRENT_PAGE = "https:" & CURRENT_PAGE
        If LANG_CD <> "ESN" And LANG_CD <> "ENU" Then LANG_CD = "ENU"
        If InStr(1, PrevLink, "WsLeaveClass") > 0 Then
            If Debug = "Y" Then mydebuglog.Debug(" Duplicate call to this agent ")
            DupeCall = True
        End If        
        
        If Debug = "Y" Then
            mydebuglog.Debug("----------------------------------")
            mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
            mydebuglog.Debug("Parameters-")
            mydebuglog.Debug("  REG_ID: " & REG_ID)
            mydebuglog.Debug("  REG_NUM: " & REG_NUM)
            mydebuglog.Debug("  cookieid: " & cookieid)
            mydebuglog.Debug("  SessID: " & SessID)
            mydebuglog.Debug("  CURRENT_PAGE: " & CURRENT_PAGE)
            mydebuglog.Debug("  HOME_PAGE: " & HOME_PAGE)
            mydebuglog.Debug("  DOMAIN: " & DOMAIN)
            mydebuglog.Debug("  LANG_CD: " & LANG_CD)
            mydebuglog.Debug("  PrevLink: " & PrevLink)
            mydebuglog.Debug("  BROWSER: " & BROWSER)
            mydebuglog.Debug("  DupeCall: " & DupeCall)
            mydebuglog.Debug("  callback: " & callback & vbCrLf)
        End If
           
        ' Validate parameters
        If REG_ID = "" Then GoTo DataError
        If REG_NUM = "" Or SessID = "" Then GoTo AccessError
        If cookieid <> REG_NUM Then
            GoTo AccessError
        End If
        
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
                mydebuglog.Debug("  .. DOMAIN: " & DOMAIN)
            End If
            If LOGGED_IN <> "Y" Then GoTo AccessError        
            
            ' ================================================
            ' LOG THE EXIT FROM THE COURSE IF NOT ALREADY LOGGED
            ' Should not be necessary if AcceptRollup service ran 
            SqlS = "IF (SELECT TOP 1 EXIT_FLG FROM siebeldb.dbo.CX_TRAIN_OFFR_ACCESS WHERE REG_ID='" & REG_ID & "' ORDER BY CREATED DESC)='N' OR " &
                "(SELECT TOP 1 EXIT_FLG FROM siebeldb.dbo.CX_TRAIN_OFFR_ACCESS WHERE REG_ID='" & REG_ID & "' ORDER BY CREATED DESC) IS NULL BEGIN; " &
                "INSERT INTO siebeldb.dbo.CX_TRAIN_OFFR_ACCESS(ROW_ID, CREATED, CREATED_BY, LAST_UPD, LAST_UPD_BY, " &
                "MODIFICATION_NUM, CONFLICT_ID, REG_ID, ENTER_FLG, EXIT_FLG, MOBILE) " &
                "SELECT '" & REG_ID & "-'+LTRIM(CAST(COUNT(*)+1 AS VARCHAR)),GETDATE(),'0-1',GETDATE(),'0-1'," &
                "0,0,'" & REG_ID & "','N','Y','Y' " &
                "FROM siebeldb.dbo.CX_TRAIN_OFFR_ACCESS " &
                "WHERE REG_ID='" & REG_ID & "'; END;"
            temp = ExecQuery("Insert", "CX_TRAIN_OFFR_ACCESS", cmd, SqlS, mydebuglog, Debug)
            
            'ENTER_FLG = ""
            'InsertExit = False
            'SqlS = "SELECT TOP 1 ENTER_FLG FROM siebeldb.dbo.CX_TRAIN_OFFR_ACCESS WHERE REG_ID='" & REG_ID & "' ORDER BY CREATED DESC"
            'If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  CHECK ACCESS SEMAPHORE: " & vbCrLf & "  " & SqlS)
            'cmd.CommandText = SqlS
            'Try
            'ENTER_FLG = CheckDBNull(cmd.ExecuteScalar(), enumObjectType.StrType)
            'Catch ex As Exception
            'InsertExit = True
            'End Try
            'If ENTER_FLG = "Y" Or ENTER_FLG = "" Then InsertExit = True
            'If Debug = "Y" Then mydebuglog.Debug("  .. InsertExit: " & InsertExit)
            'If InsertExit Then
            'SqlS = "INSERT INTO siebeldb.dbo.CX_TRAIN_OFFR_ACCESS (ROW_ID, CREATED, CREATED_BY, LAST_UPD, LAST_UPD_BY, " & _
            '"MODIFICATION_NUM, CONFLICT_ID, REG_ID, ENTER_FLG, EXIT_FLG, MOBILE) " & _
            '"SELECT '" & REG_ID & "-'+LTRIM(CAST(COUNT(*)+1 AS VARCHAR)), GETDATE(), '0-1', GETDATE(), '0-1', " & _
            '"0, 0, '" & REG_ID & "','N','Y','Y' " & _
            '"FROM siebeldb.dbo.CX_TRAIN_OFFR_ACCESS " & _
            '"WHERE REG_ID = '" & REG_ID & "'"
            'SqlS = "IF (SELECT TOP 1 EXIT_FLG FROM siebeldb.dbo.CX_TRAIN_OFFR_ACCESS WHERE REG_ID='" & REG_ID & "' ORDER BY CREATED DESC)='N' OR " &
            '    "(SELECT TOP 1 EXIT_FLG FROM siebeldb.dbo.CX_TRAIN_OFFR_ACCESS WHERE REG_ID='" & REG_ID & "' ORDER BY CREATED DESC) IS NULL BEGIN; " &
            '    "INSERT INTO siebeldb.dbo.CX_TRAIN_OFFR_ACCESS(ROW_ID, CREATED, CREATED_BY, LAST_UPD, LAST_UPD_BY, " &
            '    "MODIFICATION_NUM, CONFLICT_ID, REG_ID, ENTER_FLG, EXIT_FLG, MOBILE) " &
            '    "SELECT '" & REG_ID & "-'+LTRIM(CAST(COUNT(*)+1 AS VARCHAR)),GETDATE(),'0-1',GETDATE(),'0-1'," &
            '    "0,0,'" & REG_ID & "','N','Y','Y' " &
            '    "FROM siebeldb.dbo.CX_TRAIN_OFFR_ACCESS " &
            '    "WHERE REG_ID='" & REG_ID & "'; END;"
            'temp = ExecQuery("Insert", "CX_TRAIN_OFFR_ACCESS", cmd, SqlS, mydebuglog, Debug)
            'End If
            
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
            "WHERE R.ROW_ID='" & REG_ID & "' AND C.X_REGISTRATION_NUM='" & REG_NUM & "'"
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  REG QUERY: " & vbCrLf & "  " & SqlS)
            cmd.CommandText = SqlS
            dr = cmd.ExecuteReader()
            If Not dr Is Nothing Then
                While dr.Read()
                    Try
                        REG_STATUS_CD = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                        If DOMAIN = "" Then DOMAIN = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
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
                        CRSE_CONTENT_URL = Trim(CheckDBNull(dr(24), enumObjectType.StrType))
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
                mydebuglog.Debug("  .. REG_ID : " & REG_ID)
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
            
            ' ================================================
            ' IF THE TEST ISN'T FOUND BECAUSE A S_CRSE_TSTRUN RECORD DOESN'T EXIST, LOCATE IT
            If CRSE_TST_ID = "" Then
                If LANG_ID = "" Then LANG_ID = "ENU"
                If ALT_RETAKE_FLG = "Y" And RETAKE_FLG = "Y" Then
                    If JURIS_ID = "" Then
                        SqlS = "SELECT TOP 1 ROW_ID, X_ENGINE " & _
                        "FROM siebeldb.dbo.S_CRSE_TST " & _
                        "WHERE CRSE_ID='" & CRSE_ID & "' AND (X_JURIS_ID IS NULL OR X_JURIS_ID='') AND X_LANG_ID='" & LANG_ID & "' AND STATUS_CD='Retake' " & _
                        "ORDER BY X_VERSION DESC"
                    Else
                        SqlS = "SELECT TOP 1 ROW_ID, X_ENGINE " & _
                        "FROM siebeldb.dbo.S_CRSE_TST " & _
                        "WHERE CRSE_ID='" & CRSE_ID & "' AND X_JURIS_ID='" & JURIS_ID & "' AND X_LANG_ID='" & LANG_ID & "' AND STATUS_CD='Retake' " & _
                        "ORDER BY X_VERSION DESC"
                    End If
                Else
                    If JURIS_ID = "" Then
                        SqlS = "SELECT TOP 1 ROW_ID, X_ENGINE " & _
                        "FROM siebeldb.dbo.S_CRSE_TST " & _
                        "WHERE CRSE_ID='" & CRSE_ID & "' AND (X_JURIS_ID IS NULL OR X_JURIS_ID='') AND X_LANG_ID='" & LANG_ID & "' AND STATUS_CD='Active' " & _
                        "ORDER BY X_VERSION DESC"
                    Else
                        SqlS = "SELECT TOP 1 ROW_ID, X_ENGINE " & _
                        "FROM siebeldb.dbo.S_CRSE_TST " & _
                        "WHERE CRSE_ID='" & CRSE_ID & "' AND X_JURIS_ID='" & JURIS_ID & "' AND X_LANG_ID='" & LANG_ID & "' AND STATUS_CD='Active' " & _
                        "ORDER BY X_VERSION DESC"
                    End If
                End If
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  RETRIEVE TEST INFORMATION 1: " & vbCrLf & "  " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            CRSE_TST_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            EXAM_ENGINE = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                        Catch ex As Exception
                            GoTo DBError
                        End Try
                    End While
                End If
                dr.Close()
                
                ' Try again if a jurisdiction specific course was not found
                If CRSE_TST_ID = "" And JURIS_ID <> "" Then
                    If ALT_RETAKE_FLG = "Y" And RETAKE_FLG = "Y" Then
                        SqlS = "SELECT TOP 1 ROW_ID, X_ENGINE " & _
                        "FROM siebeldb.dbo.S_CRSE_TST " & _
                        "WHERE CRSE_ID='" & CRSE_ID & "' AND (X_JURIS_ID IS NULL OR X_JURIS_ID='') AND X_LANG_ID='" & LANG_ID & "' AND STATUS_CD='Retake' " & _
                        "ORDER BY X_VERSION DESC"
                    Else
                        SqlS = "SELECT TOP 1 ROW_ID, X_ENGINE " & _
                        "FROM siebeldb.dbo.S_CRSE_TST " & _
                        "WHERE CRSE_ID='" & CRSE_ID & "' AND (X_JURIS_ID IS NULL OR X_JURIS_ID='') AND X_LANG_ID='" & LANG_ID & "' AND STATUS_CD='Active' " & _
                        "ORDER BY X_VERSION DESC"
                    End If
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  RETRIEVE TEST INFORMATION 2: " & vbCrLf & "  " & SqlS)
                    cmd.CommandText = SqlS
                    dr = cmd.ExecuteReader()
                    If Not dr Is Nothing Then
                        While dr.Read()
                            Try
                                CRSE_TST_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                                EXAM_ENGINE = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            Catch ex As Exception
                                GoTo DBError
                            End Try
                        End While
                    End If
                    dr.Close()
                End If
                
                ' Try again if a language is not specified
                If CRSE_TST_ID = "" Then
                    SqlS = "SELECT TOP 1 ROW_ID, X_ENGINE " & _
                        "FROM siebeldb.dbo.S_CRSE_TST " & _
                        "WHERE CRSE_ID='" & CRSE_ID & "' AND (X_JURIS_ID IS NULL OR X_JURIS_ID='') AND (X_LANG_ID='' OR X_LANG_ID IS NULL) AND STATUS_CD='Active' " & _
                        "ORDER BY X_VERSION DESC"
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  RETRIEVE TEST INFORMATION 3: " & vbCrLf & "  " & SqlS)
                    cmd.CommandText = SqlS
                    dr = cmd.ExecuteReader()
                    If Not dr Is Nothing Then
                        While dr.Read()
                            Try
                                CRSE_TST_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                                EXAM_ENGINE = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            Catch ex As Exception
                                GoTo DBError
                            End Try
                        End While
                    End If
                    dr.Close()
                End If
                If Debug = "Y" Then
                    mydebuglog.Debug("  .. CRSE_TST_ID : " & CRSE_TST_ID)
                    mydebuglog.Debug("  .. EXAM_ENGINE : " & EXAM_ENGINE)
                End If
                
            End If
            
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
            ' UPDATE USER REGISTRATION COUNT FOR PORTAL
            If Not DupeCall Then
                SqlS = "UPDATE siebeldb.dbo.CX_SUB_CON " & _
                    "SET NEW_CLS=U.NUM_REG " & _
                    "FROM (SELECT COUNT(*) AS NUM_REG " & _
                    "FROM siebeldb.dbo.CX_SESS_REG " & _
                    "WHERE STATUS_CD in ('Accepted','In Progress','Retake','Exam Reqd','Tentative','Incomplete') AND CONTACT_ID='" & CONTACT_ID & "') U " & _
                    "WHERE siebeldb.dbo.CX_SUB_CON.CON_ID='" & CONTACT_ID & "'"
                temp = ExecQuery("Update", "CX_SUB_CON", cmd, SqlS, mydebuglog, Debug)
            End If
            
            ' ================================================
            ' Log to CM activity log
            SqlS = "INSERT INTO reports.dbo.CM_LOG(REG_ID, SESSION_ID, RECORD_ID, REMOTE_ADDR, ACTION, BROWSER, RECORD_DATA) " & _
                "VALUES('" & REG_NUM & "','" & SessID & "','" & REG_ID & "','" & callip & "','EXITED COURSE','" & Left(BROWSER, 200) & "','" & Replace(suspend_data, "'", "''") & "')"
            temp = ExecQuery("Insert", "CM_LOG", cmd, SqlS, mydebuglog, Debug)
            
            ' ================================================
            ' DETERMINE NEXT STEP BASED ON REGISTRATION STATUS
            ' Compute course access URL
Finished:
            If LANG_CD <> "ENU" Then
                ClassLink = LaunchProtocol & "//hciscorm.certegrity.com/ls/" & LANG_CD & "/OpenClass.html?RID=" & REG_ID & "&UID=" & REG_NUM & "&SES=" & SessID & "&PP=" & DOMAIN & "&HP=" & HOME_PAGE & "&LANG=" & LANG_CD & "&CUR=" & CURRENT_PAGE
            Else
                ClassLink = LaunchProtocol & "//hciscorm.certegrity.com/ls/OpenClass.html?RID=" & REG_ID & "&UID=" & REG_NUM & "&SES=" & SessID & "&PP=" & DOMAIN & "&HP=" & HOME_PAGE & "&CUR=" & CURRENT_PAGE
            End If
            If Debug = "Y" Then mydebuglog.Debug("  .. ClassLink : " & ClassLink)
            
            ' Determine next step from status
            If Debug = "Y" Then mydebuglog.Debug("  .. REG_STATUS_CD : " & REG_STATUS_CD)
            Select Case REG_STATUS_CD
                Case "On-Hold"            ' Go to #reg screen
                    Refresh = "10"
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = "<FONT COLOR=""Red""><b>Su clase se ha puesto en espera. <br> Póngase en contacto con nosotros para resolver este problema.</b></font><br><br>"
                            MSG1 = MSG1 & "Un momento por favor...</font>"
                        Case Else
                            MSG1 = "<FONT COLOR=""Red""><b>Your class has been put on hold.<br>Please contact us to resolve this issue.</b></font><br><br>"
                            MSG1 = MSG1 & "One moment please...</font>"
                    End Select
                    If CURRENT_PAGE <> "" Then
                        NextLink = CURRENT_PAGE
                        If InStr(NextLink, "UID=") = 0 Then NextLink = NextLink & "?UID=" & REG_NUM
                        If InStr(NextLink, "SES=") = 0 Then NextLink = NextLink & "&SES=" & SessID
                        If InStr(NextLink, "RG=") = 0 Then NextLink = NextLink & "&RG=" & REG_ID
                        If InStr(NextLink, "PP=") = 0 Then NextLink = NextLink & "&PP=" & DOMAIN
                        If InStr(NextLink, "#reg") = 0 Then NextLink = NextLink & "#reg"
                    Else
                        If LANG_CD <> "ENU" Then
                            NextLink = "https://" & HOME_PAGE & "/mobile/" & LANG_CD & "/index.html" & "?UID=" & REG_NUM & "&SES=" & SessID & "&PP=" & DOMAIN & "#reg"
                        Else
                            NextLink = "https://" & HOME_PAGE & "/mobile/index.html" & "?UID=" & REG_NUM & "&SES=" & SessID & "&PP=" & DOMAIN & "#reg"
                        End If
                    End If
                Case "Completed"         ' Go to OpenCertificate agent (completion certificate or consolation message)
                    Refresh = "40"
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = "Usted no ha terminado este curso.<br><br>"
                            MSG1 = MSG1 & "<center>Puedes <a href=https://hciscorm.certegrity.com/ls/" & LANG_CD & "/OpenCertificate.html?RID=" & REG_ID & "&UID=" & REG_NUM & "&SES=" & SessID & "&HP=" & HOME_PAGE & "&LANG=" & LANG_CD & "&CUR=" & CURRENT_PAGE & " data-role=""button"" rel=""external"" data-theme=""b"">haga clic para revisar sus resultados </a> en cualquier momento.</center>"
                        Case Else
                            MSG1 = "You have completed your course.<br><br>"
                            MSG1 = MSG1 & "<center>You may <a href=https://hciscorm.certegrity.com/ls/OpenCertificate.html?RID=" & REG_ID & "&UID=" & REG_NUM & "&SES=" & SessID & "&LANG=" & LANG_CD & "&HP=" & HOME_PAGE & "&CUR=" & CURRENT_PAGE & " data-role=""button"" rel=""external"" data-theme=""b"">click to review your results</a> at any time.</center>"
                    End Select
                    If CURRENT_PAGE <> "" Then
                        NextLink = CURRENT_PAGE
                        If InStr(NextLink, "UID=") = 0 Then NextLink = NextLink & "?UID=" & REG_NUM
                        If InStr(NextLink, "SES=") = 0 Then NextLink = NextLink & "&SES=" & SessID
                        If InStr(NextLink, "RG=") = 0 Then NextLink = NextLink & "&RG=" & REG_ID
                        If InStr(NextLink, "PP=") = 0 Then NextLink = NextLink & "&PP=" & DOMAIN
                        If InStr(NextLink, "#reg") = 0 Then NextLink = NextLink & "#reg"
                    Else
                        If LANG_CD <> "ENU" Then
                            NextLink = "https://" & HOME_PAGE & "/mobile/" & LANG_CD & "/index.html" & "?UID=" & REG_NUM & "&SES=" & SessID & "&PP=" & DOMAIN & "&RG=" & REG_ID & "#reg"
                        Else
                            NextLink = "https://" & HOME_PAGE & "/mobile/index.html" & "?UID=" & REG_NUM & "&SES=" & SessID & "&PP=" & DOMAIN & "&RG=" & REG_ID & "#reg"
                        End If
                    End If
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = MSG1 & "<br><br><a href=" & NextLink & " data-role=""button"" rel=""external"" data-theme=""b"">Haga clic para volver al portal</a>"
                        Case Else
                            MSG1 = MSG1 & "<br><br><a href=" & NextLink & " data-role=""button"" rel=""external"" data-theme=""b"">Click to return to the portal</a>"
                    End Select
                Case "Exam Reqd"         ' Go to WsGetClass agent
                    ' Assume if they have this status then they directly want to go to the exam
                    Refresh = "0"
                    If SCORM = "Y" Then
                        If LANG_CD <> "ENU" Then
                            NextLink = LaunchProtocol & "//hciscorm.certegrity.com/ls/" & LANG_CD & "/OpenAssessment.html?RID=" & REG_ID & "&TID=" & CRSE_TST_ID & "&UID=" & REG_NUM & "&SES=" & SessID & "&LANG=" & LANG_CD & "&HP=" & HOME_PAGE & "&CUR=" & CURRENT_PAGE
                        Else
                            NextLink = LaunchProtocol & "//hciscorm.certegrity.com/ls/OpenAssessment.html?RID=" & REG_ID & "&TID=" & CRSE_TST_ID & "&UID=" & REG_NUM & "&SES=" & SessID & "&LANG=" & LANG_CD & "&HP=" & HOME_PAGE & "&CUR=" & CURRENT_PAGE
                        End If
                    End If
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = "Has completado tu curso y ahora debes tomar el examen de certificaci&oacute;n.<br><br>"
                            MSG1 = "Un momento por favor...</font>"
                        Case Else
                            MSG1 = "You have completed your course, and must now take the certification exam.<br><br>"
                            MSG1 = "One moment please...</font>"
                    End Select
                Case "Retake"            ' Go to OpenClass
                    Refresh = "40"
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = "Puedes retomar el curso en cualquier momento.<br><br>"
                            MSG1 = MSG1 & "<center><a href=" & ClassLink & " data-role=""button"" rel=""external"" data-theme=""b"">Haz click para tomar la clase</A></center>"
                        Case Else
                            MSG1 = "You may retake the course at any time.<br><br>"
                            MSG1 = MSG1 & "<center><a href=" & ClassLink & " data-role=""button"" rel=""external"" data-theme=""b"">Click to take the class</A></center>"
                    End Select
                    If CURRENT_PAGE <> "" Then
                        NextLink = CURRENT_PAGE
                        If InStr(NextLink, "UID=") = 0 Then NextLink = NextLink & "?UID=" & REG_NUM
                        If InStr(NextLink, "SES=") = 0 Then NextLink = NextLink & "&SES=" & SessID
                        If InStr(NextLink, "RG=") = 0 Then NextLink = NextLink & "&RG=" & REG_ID
                        If InStr(NextLink, "PP=") = 0 Then NextLink = NextLink & "&PP=" & DOMAIN
                        If InStr(NextLink, "#reg") = 0 Then NextLink = NextLink & "#reg"
                    Else
                        If LANG_CD <> "ENU" Then
                            NextLink = "https://" & HOME_PAGE & "/mobile/" & LANG_CD & "/index.html" & "?UID=" & REG_NUM & "&SES=" & SessID & "&PP=" & DOMAIN & "#reg"
                        Else
                            NextLink = "https://" & HOME_PAGE & "/mobile/index.html" & "?UID=" & REG_NUM & "&SES=" & SessID & "&PP=" & DOMAIN & "#reg"
                        End If
                    End If
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = MSG1 & "<br><br><a href=" & NextLink & " data-role=""button"" rel=""external"" data-theme=""b"">Haga clic para volver a su portal</a>"
                        Case Else
                            MSG1 = MSG1 & "<br><br><a href=" & NextLink & " data-role=""button"" rel=""external"" data-theme=""b"">Click to return to the portal</a>"
                    End Select
                Case "Accepted"            ' Set status to "In Progress" - handle as "In Progress"
                    Refresh = "40"
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = "Usted no ha terminado este curso. Usted puede continuar su curso en cualquier momento.<br><br>"
                            MSG1 = MSG1 & "Hazlo ahora <a href=" & ClassLink & "  data-role=""button"" rel=""external"" data-theme=""b"">Al hacer clic en este enlace</A>"
                        Case Else
                            MSG1 = "You have not completed this course.  You may begin or continue your course at any time.<br><br>"
                            MSG1 = MSG1 & "Do so now by <a href=" & ClassLink & "  data-role=""button"" rel=""external"" data-theme=""b"">Clicking on this link</A>"
                    End Select
                    If CURRENT_PAGE <> "" Then
                        NextLink = CURRENT_PAGE
                        If InStr(NextLink, "UID=") = 0 Then NextLink = NextLink & "?UID=" & REG_NUM
                        If InStr(NextLink, "SES=") = 0 Then NextLink = NextLink & "&SES=" & SessID
                        If InStr(NextLink, "RG=") = 0 Then NextLink = NextLink & "&RG=" & REG_ID
                        If InStr(NextLink, "PP=") = 0 Then NextLink = NextLink & "&PP=" & DOMAIN
                        If InStr(NextLink, "#reg") = 0 Then NextLink = NextLink & "#reg"
                    Else
                        If LANG_CD <> "ENU" Then
                            NextLink = "https://" & HOME_PAGE & "/mobile/" & LANG_CD & "/index.html" & "?UID=" & REG_NUM & "&SES=" & SessID & "&PP=" & DOMAIN & "#reg"
                        Else
                            NextLink = "https://" & HOME_PAGE & "/mobile/index.html" & "?UID=" & REG_NUM & "&SES=" & SessID & "&PP=" & DOMAIN & "#reg"
                        End If
                    End If
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = MSG1 & "<br><br><a href=" & NextLink & " data-role=""button"" rel=""external"" data-theme=""b"">Haga clic para volver a su portal</a>"
                        Case Else
                            MSG1 = MSG1 & "<br><br><a href=" & NextLink & " data-role=""button"" rel=""external"" data-theme=""b"">Click to return to the portal</a>"
                    End Select
                Case "In Progress"         ' Provide ability to re-enter or close course      
                    Refresh = "40"
                    MSG1 = "<center><h3 style=""color: red;"">"
                    If EXAM_ENGINE = "SCORM" And SCORM <> "Y" And MayTest Then
                        Refresh = "0"
                        If LANG_CD <> "ENU" Then
                            ExamLink = LaunchProtocol & "//hciscorm.certegrity.com/ls/" & LANG_CD & "/OpenAssessment.html?RID=" & REG_ID & "&TID=" & CRSE_TST_ID & "&UID=" & REG_NUM & "&SES=" & SessID & "&LANG=" & LANG_CD & "&HP=" & HOME_PAGE & "&FNC=Y&CUR=" & CURRENT_PAGE
                        Else
                            ExamLink = LaunchProtocol & "//hciscorm.certegrity.com/ls/OpenAssessment.html?RID=" & REG_ID & "&TID=" & CRSE_TST_ID & "&UID=" & REG_NUM & "&SES=" & SessID & "&HP=" & HOME_PAGE & "&FNC=Y&CUR=" & CURRENT_PAGE
                        End If
                        Select Case LANG_CD
                            Case "ESN"
                                MSG1 = MSG1 & "Usted est&aacute; al final del curso y puede <a href=" & ExamLink & "  data-role=""button"" data-theme=""b"">" & _
                                "realizar el examen de certificaci&oacute;n</a> ahora si elige.<br>Tambi&eacute;n puedes salir y hacer el examen m&aacute;s tarde."
                            Case Else
                                MSG1 = MSG1 & "You are at the end of the course and may <a href=" & ExamLink & "  data-role=""button"" data-theme=""b"">" & _
                                "take the certification exam</a> now if you choose.<br>You may also exit and take the exam later."
                        End Select
                    Else
                        If Not CompletedCourse Then
                            Select Case LANG_CD
                                Case "ESN"
                                    MSG1 = MSG1 & "No ha completado este curso. <br />Puede continuar su curso en cualquier momento.<br />"
                                    MSG1 = MSG1 & "<br></h3><a href=" & ClassLink & " data-role=""button"" rel=""external"" data-theme=""b"">Haga clic para volver a su curso</a><br>"
                                Case Else
                                    MSG1 = MSG1 & "You have not completed this course.<br> You may continue your course at any time.<br>"
                                    MSG1 = MSG1 & "<br></h3><a href=" & ClassLink & " data-role=""button"" rel=""external"" data-theme=""b"">Click to return to your course</a><br>"
                            End Select
                        End If
                        If MayTest Then
                            If LANG_CD <> "ENU" Then
                                ExamLink = LaunchProtocol & "//hciscorm.certegrity.com/ls/" & LANG_CD & "/OpenAssessment.html?RID=" & REG_ID & "&TID=" & CRSE_TST_ID & "&UID=" & REG_NUM & "&SES=" & SessID & "&LANG=" & LANG_CD & "&HP=" & HOME_PAGE & "&FNC=Y&CUR=" & CURRENT_PAGE
                            Else
                                ExamLink = LaunchProtocol & "//hciscorm.certegrity.com/ls/OpenAssessment.html?RID=" & REG_ID & "&TID=" & CRSE_TST_ID & "&UID=" & REG_NUM & "&SES=" & SessID & "&HP=" & HOME_PAGE & "&FNC=Y&CUR=" & CURRENT_PAGE
                            End If
                            Select Case LANG_CD
                                Case "ESN"
                                    MSG1 = MSG1 & "<br>Est&aacute; en la &uacute;ltima pantalla del curso y puede tomar el examen de certificaci&oacute;n <br> si lo desea. " & _
                                    "Si comienzas tu examen no podr&aacute;s volver al curso<br><br></h3><a href=" & ExamLink & "  data-role=""button"" rel=""external"" data-theme=""b"">" & _
                                    "Tomar el examen de certificaci&oacute;n</a><br> ."
                                Case Else
                                    MSG1 = MSG1 & "<br>You are on the last screen of the course, and can take the certification exam<br>if you choose. " & _
                                    "If you begin your exam you may not return to the course<br><br></h3><a href=" & ExamLink & "  data-role=""button"" rel=""external"" data-theme=""b"">" & _
                                    "Take the Certification Exam</a><br> ."
                            End Select
                        End If
                    End If
                    If CURRENT_PAGE <> "" Then
                        NextLink = CURRENT_PAGE
                        If InStr(NextLink, "UID=") = 0 Then NextLink = NextLink & "?UID=" & REG_NUM
                        If InStr(NextLink, "SES=") = 0 Then NextLink = NextLink & "&SES=" & SessID
                        If InStr(NextLink, "RG=") = 0 Then NextLink = NextLink & "&RG=" & REG_ID
                        If InStr(NextLink, "PP=") = 0 Then NextLink = NextLink & "&PP=" & DOMAIN
                        If InStr(NextLink, "#reg") = 0 Then NextLink = NextLink & "#reg"
                    Else
                        If LANG_CD <> "ENU" Then
                            NextLink = "https://" & HOME_PAGE & "/mobile/" & LANG_CD & "/index.html" & "?UID=" & REG_NUM & "&PP=" & DOMAIN & "&SES=" & SessID & "#reg"
                        Else
                            NextLink = "https://" & HOME_PAGE & "/mobile/index.html" & "?UID=" & REG_NUM & "&PP=" & DOMAIN & "&SES=" & SessID & "#reg"
                        End If
                    End If
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = MSG1 & "<br><a href=" & NextLink & " data-role=""button"" rel=""external"" data-theme=""b"">Haga clic para volver a su portal</a>"
                        Case Else
                            MSG1 = MSG1 & "<br><a href=" & NextLink & " data-role=""button"" rel=""external"" data-theme=""b"">Click to return to the portal</a>"
                    End Select
                    MSG1 = MSG1 & "</center>"
                Case Else                  ' Error out
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = " No podemos localizar su registro. Puede haber un problema con nuestro sistema. <br> <br> P&oacute;ngase en contacto con nuestro Soporte T&eacute;cnico Departamento de asistencia."
                        Case Else
                            MSG1 = " We are unable to locate your registration.  There may be an issue with our system. <br><br>Please contact our Technical Support department for assistance."
                    End Select
            End Select
            If Debug = "Y" Then
                mydebuglog.Debug("  .. MSG1 : " & MSG1)
                mydebuglog.Debug("  .. ExamLink : " & ExamLink)
                mydebuglog.Debug("  .. NextLink : " & NextLink)
            End If
            
        Else
            GoTo DBError
        End If
        GoTo CloseOut

Incomplete:
        If Debug = "Y" Then mydebuglog.Debug(">>Incomplete")
        Select Case LANG_CD
            Case "ESN"
                errmsg = "No hemos podido procesar su salida de clase. Por favor cont&aacute;ctenos para asistencia."
            Case Else
                errmsg = "We were unable to process your class exit.  Please contact us for assistance."
        End Select
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
        ' ============================================
        ' Finalize output        
        jdoc = """Refresh"":""" & Refresh & ""","
        jdoc = jdoc & """REG_STATUS_CD"":""" & REG_STATUS_CD & ""","
        jdoc = jdoc & """MSG1"":""" & EscapeJSON(MSG1) & ""","
        jdoc = jdoc & """CRSE_NAME"":""" & EscapeJSON(CRSE_NAME) & ""","
        jdoc = jdoc & """redirect"":""" & EscapeJSON(redirect) & ""","
        jdoc = jdoc & """NextLink"":""" & EscapeJSON(NextLink) & ""","
        jdoc = jdoc & """ExamLink"":""" & EscapeJSON(ExamLink) & ""","
        jdoc = jdoc & """ClassLink"":""" & EscapeJSON(ClassLink) & ""","
        jdoc = jdoc & """ErrMsg"":""" & errmsg & ""","
        jdoc = callback & "({""ResultSet"": {" & jdoc & "} })"
        
        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("WsLeaveClass.ashx : " & ErrLvl & ": " & Trim(errmsg))
        myeventlog.Info("WsLeaveClass.ashx : UID: " & REG_NUM & ", SessID: " & SessID & ", and RegID:" & REG_ID & ", NextLink: " & NextLink)
        If Debug = "Y" Or (Logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("  JDOC: " & jdoc & vbCrLf)
                mydebuglog.Debug("Results:  UID: " & REG_NUM & ", SessID: " & SessID & ", and RegID:" & REG_ID & ", NextLink: " & NextLink)
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
                LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, "WSLEAVECLASS", LogStartTime, VersionNum, Debug)
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