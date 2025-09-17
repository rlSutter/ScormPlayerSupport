<%@ WebHandler Language="VB" Class="WsGetAssessment" %>

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

Public Class WsGetAssessment : Implements IHttpHandler
    
    Enum enumObjectType
        StrType = 0
        IntType = 1
        DblType = 2
        DteType = 3
    End Enum
    
    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        
        ' Parameter Declarations
        Dim REG_ID, UID, SessID, CURRENT_PAGE, HOME_PAGE, LANG_CD, callback, myprotocol As String
        Dim Debug As String
        
        ' Result Declarations
        Dim outdata As String
        
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
        mydebuglog = log4net.LogManager.GetLogger("GetAssessmentDebugLog")
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
        
        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service
        Dim Processing As New com.certegrity.cloudsvc.processing.Service
        Dim Scormsvc As New com.certegrity.hciscormsvc.scorm.Service
        
        ' Variable declarations
        Dim errmsg, EOL, ErrLvl, MSG1, LAST_INST As String
        Dim LOGGED_IN, CONTACT_ID, SUB_ID, CONTACT_OU_ID, DOMAIN, LogoutUser As String
        Dim REG_STATUS_CD, CRSE_CONTENT_URL, SCORM, CRSE_ID, REGISTRANT_ID, CRSE_TST_ID, EXAM_ENGINE As String
        Dim COURSE, JURIS_ID, KBA_REQD, temp, TEST_FLG, USER_NAME, TEST_STATUS_CD As String
        Dim EMAIL_ADDR, X_FORMAT, FST_NAME, LAST_NAME, PAUSE As String
        Dim ClassLink, NextLink, RefreshLink, ExamLink, LaunchProtocol As String
        Dim mailpath As String
        Dim KBA_QUES_NUM, ENGINE, EXAM, CALL_SCREEN, CALL_ID, EXAM_ID As String
        Dim attachlink, AccessToken As String
        Dim TYPE_CD, SESS_PART_ID, sCALL_SCREEN, REMOTE_ADDR, COURSE_TYPE_CD As String
        Dim TST_ID, FINISH_CLASS, UserKey, ANON_REF_ID, CON_ID As String
        Dim X_SURVEY_FLG, TRX_ID, CONTINUABLE_FLG, LANG_ID, LANG, ALT_JURIS_ID, CRSE_JURIS_ID, SKILL_LEVEL_CD, CATEGORY As String
        Dim RETAKE_FLG, ALT_RETAKE_FLG, TST_RETAKE_FLG, RESOLUTION, ENTER_FLG, EXIT_FLG As String
        Dim UnLoadLink, AutoStart, delay, ALT_ASSESS, EXAM_STATUS_CD, RES_X, RES_Y, PERSON_ID As String
        Dim HCIPlayer As Boolean
        Dim ACTIVITY_COUNT As Integer
        Dim SCORM_ACTIVITY, btext As String
        Dim ALT_EXAMS, ALT_TEST, ALT_TEST_ID, ALT_LANG_CD, ALT_LANG, CRSE_NAME, RefreshID As String
        Dim ctr, lockoutcount As Integer

        ' ============================================
        ' Variable setup
        Debug = "N"
        lockoutcount = 0
        errmsg = ""
        LOGGED_IN = "N"
        TEST_STATUS_CD = ""
        CONTACT_ID = ""
        SUB_ID = ""
        RESOLUTION = ""
        CONTACT_OU_ID = ""
        DOMAIN = "TIPS"
        LogoutUser = "N"
        LANG_CD = ""
        HOME_PAGE = ""
        REG_ID = ""
        myprotocol = ""
        callback = ""
        CURRENT_PAGE = ""
        Logging = "Y"
        REGISTRANT_ID = ""
        CRSE_TST_ID = ""
        EXAM_ENGINE = ""
        CRSE_CONTENT_URL = ""
        TEST_FLG = ""
        KBA_REQD = ""
        LANG_CD = ""
        COURSE = ""
        FST_NAME = ""
        LAST_NAME = ""
        EMAIL_ADDR = ""
        X_FORMAT = ""
        REG_STATUS_CD = ""
        COURSE_TYPE_CD = ""
        attachlink = ""
        AccessToken = ""
        EOL = Chr(13)
        ErrLvl = "Error"
        PAUSE = "0"
        LAST_INST = ""
        TYPE_CD = ""
        UID = ""
        SessID = ""
        REG_ID = ""
        mailpath = ""
        CRSE_ID = ""
        SCORM = ""
        outdata = ""
        USER_NAME = ""
        ClassLink = ""
        NextLink = ""
        RefreshLink = ""
        ExamLink = ""
        KBA_QUES_NUM = ""
        ENGINE = ""
        EXAM = ""
        CALL_SCREEN = ""
        CALL_ID = ""
        MSG1 = ""
        EXAM_ID = ""
        SESS_PART_ID = ""
        sCALL_SCREEN = ""
        REMOTE_ADDR = ""
        TST_ID = ""
        FINISH_CLASS = ""
        UserKey = ""
        LaunchProtocol = "http:"
        ANON_REF_ID = ""
        CON_ID = ""
        RETAKE_FLG = ""
        ALT_RETAKE_FLG = ""
        TST_RETAKE_FLG = ""
        SKILL_LEVEL_CD = ""
        LANG_ID = ""
        LANG = ""
        X_SURVEY_FLG = ""
        CONTINUABLE_FLG = ""
        JURIS_ID = ""
        ALT_JURIS_ID = ""
        CRSE_JURIS_ID = ""
        PERSON_ID = ""
        HCIPlayer = False
        delay = "0"
        EXIT_FLG = ""
        ENTER_FLG = ""
        SCORM_ACTIVITY = ""
        btext = ""
        ACTIVITY_COUNT = 0
        ALT_EXAMS = ""
        ALT_ASSESS = "N"
        EXAM_STATUS_CD = ""
        AutoStart = False
        RES_X = ""
        RES_Y = ""
        UnLoadLink = ""
        RefreshID = ""
        CATEGORY = ""
        
        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server=DB_SERVER;uid=DB_USER;pwd=DB_PASSWORD;database=siebeldb"
            tempdebug = System.Configuration.ConfigurationManager.AppSettings.Get("GetAssessment_debug")
            If tempdebug = "Y" And Debug <> "T" Then Debug = "Y"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("Lockoutcount")
            If temp <> "" And IsNumeric(temp) Then lockoutcount = Val(temp)
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            GoTo CloseOut2
        End Try
        
        ' ============================================
        ' Open log file if applicable
        If Logging = "Y" Then
            logfile = "C:\Logs\WsGetAssessment.log"
            Try
                log4net.GlobalContext.Properties("GetAssessmentLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                GoTo CloseOut2
            End Try
        End If
                
        ' ============================================
        ' Get parameters    
        If Not context.Request.QueryString("RID") Is Nothing Then
            REG_ID = context.Request.QueryString("RID")
        End If
        
        If Not context.Request.QueryString("TID") Is Nothing Then
            TST_ID = context.Request.QueryString("TID")
        End If

        If Not context.Request.QueryString("EID") Is Nothing Then
            EXAM_ID = context.Request.QueryString("EID")
        End If
        
        If Not context.Request.QueryString("FNC") Is Nothing Then
            FINISH_CLASS = context.Request.QueryString("FNC")
        End If

        If Not context.Request.QueryString("HP") Is Nothing Then
            HOME_PAGE = context.Request.QueryString("HP")
        End If

        If Not context.Request.QueryString("UID") Is Nothing Then
            UID = context.Request.QueryString("UID")
        End If

        If Not context.Request.QueryString("SES") Is Nothing Then
            SessID = context.Request.QueryString("SES")
        End If

        If Not context.Request.QueryString("PROT") Is Nothing Then
            myprotocol = LCase(context.Request.QueryString("PROT"))
        End If

        If Not context.Request.QueryString("CUR") Is Nothing Then
            CURRENT_PAGE = context.Request.QueryString("CUR")
        End If
        
        If Not context.Request.QueryString("RFR") Is Nothing Then
            RefreshID = context.Request.QueryString("RFR")
        End If        
 
        If Not context.Request.QueryString("TYP") Is Nothing Then
            TYPE_CD = LCase(context.Request.QueryString("TYP"))
        End If

        If Not context.Request.QueryString("LANG") Is Nothing Then
            LANG_CD = UCase(context.Request.QueryString("LANG"))
        End If
 
        If Not context.Request.QueryString("PP") Is Nothing Then
            DOMAIN = UCase(context.Request.QueryString("PP"))
        End If
        
        If Not context.Request.QueryString("callback") Is Nothing Then
            callback = context.Request.QueryString("callback")
        End If
        
        ' Validate parameters
        If myprotocol = "" Then myprotocol = "http:"
        If TYPE_CD = "" Then TYPE_CD = "exam"
        If LANG_CD <> "ESN" And LANG_CD <> "ENU" Then LANG_CD = "ENU"
        If callback = "" Then callback = "?"
        If InStr(1, PrevLink, "?UID") = 0 Then PrevLink = PrevLink & "?UID=" & UID & "&SES=" & SessID
        PrevLink = Replace(PrevLink, "#reg", "")
 
        If Debug = "Y" Then
            mydebuglog.Debug("----------------------------------")
            mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
            mydebuglog.Debug("Parameters-")
            mydebuglog.Debug("  lockoutcount: " & Str(lockoutcount))
            mydebuglog.Debug("  UID: " & UID)
            mydebuglog.Debug("  cookieid: " & cookieid)
            mydebuglog.Debug("  SessID: " & SessID)
            mydebuglog.Debug("  REG_ID: " & REG_ID)
            mydebuglog.Debug("  TST_ID: " & TST_ID)
            mydebuglog.Debug("  EXAM_ID: " & EXAM_ID)
            mydebuglog.Debug("  HOME_PAGE: " & HOME_PAGE)
            mydebuglog.Debug("  myprotocol: " & myprotocol)
            mydebuglog.Debug("  CURRENT_PAGE : " & CURRENT_PAGE)
            mydebuglog.Debug("  RefreshID : " & RefreshID)            
            mydebuglog.Debug("  DOMAIN: " & DOMAIN)
            mydebuglog.Debug("  LANG_CD: " & LANG_CD)
            mydebuglog.Debug("  PrevLink: " & PrevLink)
            mydebuglog.Debug("  BROWSER: " & BROWSER)
            mydebuglog.Debug("  callback: " & callback)
        End If        
        If Left(HOME_PAGE, 4) <> "web." And Left(HOME_PAGE, 4) <> "www." And HOME_PAGE <> "certegrity.com" Then
            If InStr(1, PrevLink, "web.") > 0 Then HOME_PAGE = "web." & HOME_PAGE Else HOME_PAGE = "www." & HOME_PAGE
        End If
        If Debug = "Y" Then mydebuglog.Debug("  HOME_PAGE: " & HOME_PAGE)
        If REG_ID = "" And EXAM_ID = "" And TST_ID = "" Then GoTo AccessError
        
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
            If UID <> "" And SessID <> "" Then
                SqlS = "SELECT TOP 1 (SELECT CASE WHEN H.LOGOUT_DT IS NULL THEN 'Y' ELSE 'N' END) AS LOGGED_IN, SC.CON_ID, SC.SUB_ID, C.PR_DEPT_OU_ID, S.DOMAIN " & _
                    "FROM siebeldb.dbo.CX_SUB_CON_HIST H " & _
                    "LEFT OUTER JOIN siebeldb.dbo.CX_SUB_CON SC ON SC.ROW_ID=H.SUB_CON_ID " & _
                    "LEFT OUTER JOIN siebeldb.dbo.CX_SUBSCRIPTION S ON S.ROW_ID=SC.SUB_ID " & _
                    "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=SC.CON_ID " & _
                    "WHERE USER_ID='" & UID & "' AND SESSION_ID='" & SessID & "'"
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
                    GoTo AccessError
                End Try
                dr.Close()
                If Debug = "Y" Then
                    mydebuglog.Debug("  .. LOGGED_IN: " & LOGGED_IN)
                    mydebuglog.Debug("  .. CONTACT_ID: " & CONTACT_ID)
                    mydebuglog.Debug("  .. CONTACT_OU_ID: " & CONTACT_OU_ID)
                    mydebuglog.Debug("  .. DOMAIN: " & DOMAIN)
                End If
                If LOGGED_IN = "N" Then GoTo AutoLoggedOut
                If LOGGED_IN = "" Then GoTo AccessError

                ' Generate user key
                Dim data As Byte()
                data = System.Text.ASCIIEncoding.ASCII.GetBytes(UID)
                UserKey = ReverseString(System.Convert.ToBase64String(data))      ' Generate user key
                If Debug = "Y" Then mydebuglog.Debug("  .. UserKey: " & UserKey)
            
            End If
            
            ' ================================================
            ' ASSESSED FAILED ATTEMPT
            If EXAM_ID = "Failure" Then
                If REG_ID = "" And TST_ID = "" Then
                    If UID <> "" Then
                        ' Try to locate existing registration id
                        SqlS = "SELECT R.ROW_ID, SP.ROW_ID " & _
                        "FROM siebeldb.dbo.CX_SESS_PART_X SP " & _
                        "LEFT OUTER JOIN siebeldb.dbo.CX_SESS_REG R ON R.SESS_PART_ID=SP.ROW_ID " & _
                        "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=R.CONTACT_ID " & _
                        "WHERE C.X_REGISTRATION_NUM='" & UID & "' AND SP.CRSE_TSTRUN_ID='Failure'"
                        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Locating Failed Registration: " & vbCrLf & "  " & SqlS)
                        cmd.CommandText = SqlS
                        dr = cmd.ExecuteReader()
                        If Not dr Is Nothing Then
                            While dr.Read()
                                Try
                                    REG_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                                    SESS_PART_ID = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                                Catch ex As Exception
                                    MSG1 = ex.Message
                                    GoTo DBError
                                End Try
                            End While
                        Else
                            GoTo DataError
                        End If
                        dr.Close()
                        If Debug = "Y" Then
                            mydebuglog.Debug("  .. REG_ID : " & REG_ID)
                            mydebuglog.Debug("  .. SESS_PART_ID  : " & SESS_PART_ID)
                        End If
                        If REG_ID = "" Then GoTo AccessError ' No registration, leave
                        If SESS_PART_ID <> "" Then
                            SqlS = "UPDATE siebeldb.dbo.CX_SESS_PART_X SET CRSE_TSTRUN_ID=NULL WHERE ROW_ID='" & SESS_PART_ID & "'"
                            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Resetting test-run id: " & vbCrLf & "  " & SqlS)
                            temp = ExecQuery("Update", "S_CRSE_TSTRUN", cmd, SqlS, mydebuglog, "N")
                        End If
                    Else
                        GoTo AccessError
                    End If
                Else
                    EXAM_ID = ""
                End If
            End If
            
            ' ================================================
            ' EVALUATE PARAMETERS AND QUERY FOR OTHERS
            ' Determine if there is already an exam if a registration was supplied
            If REG_ID <> "" Then
                SqlS = "SELECT SP.ROW_ID, SP.CRSE_TSTRUN_ID, R.JURIS_ID, R.CRSE_ID, TR.CRSE_TST_ID, " & _
                      "R.STATUS_CD, T.X_SURVEY_FLG, T.X_RESOLUTION, R.TEST_FLG, TRX.ROW_ID, T.X_CONTINUABLE_FLG, " & _
                      "(SELECT CASE WHEN TC.LANG_ID IS NULL THEN CR.X_LANG_CD ELSE TC.LANG_ID END) AS LANG_ID, L.NAME, " & _
                      "T.X_KBA_QUES_NUM, (SELECT CASE WHEN A.X_JURIS_ID IS NULL THEN PA.X_JURIS_ID ELSE A.X_JURIS_ID END) AS ALT_JURIS_ID, " & _
                      "CR.X_JURIS_ID, T.X_FORMAT, C.FST_NAME+' '+C.LAST_NAME, C.EMAIL_ADDR, T.SKILL_LEVEL_CD, CR.X_CRSE_CONTENT_URL, R.RETAKE_FLG, " & _
                      "CR.X_ALT_RETAKE_FLG, T.X_RETAKE_FLG, CR.TYPE_CD, TR.STATUS_CD " & _
                      "FROM siebeldb.dbo.CX_SESS_REG R " & _
                      "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=R.CONTACT_ID " & _
                      "LEFT OUTER JOIN siebeldb.dbo.S_ADDR_ORG A ON A.ROW_ID=R.ADDR_ID " & _
                      "LEFT OUTER JOIN siebeldb.dbo.S_ADDR_PER PA ON PA.ROW_ID=R.PER_ADDR_ID " & _
                      "LEFT OUTER JOIN siebeldb.dbo.CX_TRAIN_OFFR TC ON TC.ROW_ID=R.TRAIN_OFFR_ID " & _
                      "LEFT OUTER JOIN siebeldb.dbo.S_CRSE CR ON CR.ROW_ID=R.CRSE_ID " & _
                      "LEFT OUTER JOIN siebeldb.dbo.CX_SESS_PART_X SP ON SP.ROW_ID=R.SESS_PART_ID " & _
                      "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TSTRUN TR ON TR.ROW_ID=SP.CRSE_TSTRUN_ID " & _
                      "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TSTRUN_X TRX ON TRX.PAR_ROW_ID=TR.ROW_ID " & _
                      "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TST T ON T.ROW_ID=TR.CRSE_TST_ID " & _
                      "LEFT OUTER JOIN siebeldb.dbo.S_LANG L ON L.LANG_CD=T.X_LANG_ID " & _
                      "WHERE R.ROW_ID='" & REG_ID & "'"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  CHECK FOR EXAM W/REG_ID: " & vbCrLf & "  " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            temp = ""
                            SESS_PART_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            EXAM_ID = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            JURIS_ID = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                            CRSE_ID = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                            temp = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                            If temp <> "" Then TST_ID = temp
                            REG_STATUS_CD = Trim(CheckDBNull(dr(5), enumObjectType.StrType))
                            X_SURVEY_FLG = Trim(CheckDBNull(dr(6), enumObjectType.StrType))
                            RESOLUTION = Trim(CheckDBNull(dr(7), enumObjectType.StrType))
                            TEST_FLG = Trim(CheckDBNull(dr(8), enumObjectType.StrType))
                            TRX_ID = Trim(CheckDBNull(dr(9), enumObjectType.StrType))
                            CONTINUABLE_FLG = Trim(CheckDBNull(dr(10), enumObjectType.StrType))
                            LANG_ID = Trim(CheckDBNull(dr(11), enumObjectType.StrType))
                            LANG = Trim(CheckDBNull(dr(12), enumObjectType.StrType))
                            KBA_QUES_NUM = Trim(Str(CheckDBNull(dr(13), enumObjectType.IntType)))
                            ALT_JURIS_ID = Trim(CheckDBNull(dr(14), enumObjectType.StrType))
                            CRSE_JURIS_ID = Trim(CheckDBNull(dr(15), enumObjectType.StrType))
                            X_FORMAT = Trim(CheckDBNull(dr(16), enumObjectType.StrType))
                            USER_NAME = Trim(CheckDBNull(dr(17), enumObjectType.StrType))
                            EMAIL_ADDR = Trim(CheckDBNull(dr(18), enumObjectType.StrType))
                            SKILL_LEVEL_CD = Trim(CheckDBNull(dr(19), enumObjectType.StrType))
                            CRSE_CONTENT_URL = Trim(CheckDBNull(dr(20), enumObjectType.StrType))
                            RETAKE_FLG = Trim(CheckDBNull(dr(21), enumObjectType.StrType))
                            ALT_RETAKE_FLG = Trim(CheckDBNull(dr(22), enumObjectType.StrType))
                            TST_RETAKE_FLG = Trim(CheckDBNull(dr(23), enumObjectType.StrType))
                            COURSE_TYPE_CD = Trim(CheckDBNull(dr(24), enumObjectType.StrType))
                            TEST_STATUS_CD = Trim(CheckDBNull(dr(25), enumObjectType.StrType))
                        Catch ex As Exception
                            MSG1 = ex.Message
                            GoTo DBError
                        End Try
                    End While
                Else
                    GoTo DataError
                End If
                dr.Close()
                If Debug = "Y" Then
                    mydebuglog.Debug("  .. REG_ID  : " & REG_ID)
                    mydebuglog.Debug("  .. JURIS_ID: " & JURIS_ID)
                    mydebuglog.Debug("  .. ALT_JURIS_ID: " & ALT_JURIS_ID)
                    mydebuglog.Debug("  .. CRSE_JURIS_ID: " & CRSE_JURIS_ID)
                    mydebuglog.Debug("  .. COURSE_TYPE_CD: " & COURSE_TYPE_CD)
                    mydebuglog.Debug("  .. TEST_STATUS_CD: " & TEST_STATUS_CD)
                    mydebuglog.Debug("  .. REG_STATUS_CD: " & REG_STATUS_CD)
                    mydebuglog.Debug("  .. EXAM_ID: " & EXAM_ID & vbCrLf)
                End If
                
                ' Check for jurisdiction course/registration mismatch
                If CRSE_JURIS_ID <> "" Then
                    
                    ' The registration jurisdiction matches the course jurisdiction
                    If JURIS_ID = CRSE_JURIS_ID Then
                        If Debug = "Y" Then mydebuglog.Debug(">>REGISTRATION AND COURSE JURISDICTION AGREE")
                        GoTo ByPassJuris
                    End If
                    
                    ' The alternative jurisdiction is correct
                    If JURIS_ID <> CRSE_JURIS_ID And ALT_JURIS_ID = CRSE_JURIS_ID Then
                        If Debug = "Y" Then mydebuglog.Debug(">>RESETTING JURISDICTION ID TO: " & ALT_JURIS_ID)
                        JURIS_ID = ALT_JURIS_ID
                        SqlS = "UPDATE siebeldb.dbo.CX_SESS_REG " & _
                        "SET JURIS_ID='" & JURIS_ID & "' " & _
                        "WHERE ROW_ID='" & REG_ID & "'"
                        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  UPDATE JURISDICTION ID: " & vbCrLf & "  " & SqlS)
                        temp = ExecQuery("Update", "CX_SESS_REG", cmd, SqlS, mydebuglog, "N")
                        GoTo ByPassJuris
                    End If
                    
                    ' The registration jurisdiction is empty
                    If (JURIS_ID = "" Or JURIS_ID = "ZZ") And ALT_JURIS_ID = "" Then
                        If Debug = "Y" Then mydebuglog.Debug(">>RESETTING JURISDICTION ID TO: " & CRSE_JURIS_ID)
                        JURIS_ID = CRSE_JURIS_ID
                        SqlS = "UPDATE siebeldb.dbo.CX_SESS_REG " & _
                        "SET JURIS_ID='" & JURIS_ID & "' " & _
                        "WHERE ROW_ID='" & REG_ID & "'"
                        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  UPDATE JURISDICTION ID: " & vbCrLf & "  " & SqlS)
                        temp = ExecQuery("Update", "CX_SESS_REG", cmd, SqlS, mydebuglog, "N")
                        GoTo ByPassJuris
                    End If
                    
                    ' The registration and alternative jurisdictions are wrong
                    If JURIS_ID <> CRSE_JURIS_ID And ALT_JURIS_ID <> CRSE_JURIS_ID Then
                        If Debug = "Y" Then mydebuglog.Debug(">>RESETTING JURISDICTION ID TO: " & CRSE_JURIS_ID)
                        JURIS_ID = CRSE_JURIS_ID
                        SqlS = "UPDATE siebeldb.dbo.CX_SESS_REG " & _
                        "SET JURIS_ID='" & JURIS_ID & "' " & _
                        "WHERE ROW_ID='" & REG_ID & "'"
                        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  UPDATE JURISDICTION ID: " & vbCrLf & "  " & SqlS)
                        temp = ExecQuery("Update", "CX_SESS_REG", cmd, SqlS, mydebuglog, "N")
                    End If
                End If
                
ByPassJuris:
                ' Verify jurisdiction id and correct
                If JURIS_ID = "" Or JURIS_ID = "ZZ" Then
                    If ALT_JURIS_ID <> "" Then
                        If Debug = "Y" Then mydebuglog.Debug(">>RESETTING JURISDICTION ID TO: " & ALT_JURIS_ID)
                        JURIS_ID = ALT_JURIS_ID
                        SqlS = "UPDATE siebeldb.dbo.CX_SESS_REG " & _
                        "SET JURIS_ID='" & JURIS_ID & "' " & _
                        "WHERE ROW_ID='" & REG_ID & "'"
                        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  UPDATE JURISDICTION ID: " & vbCrLf & "  " & SqlS)
                        temp = ExecQuery("Update", "CX_SESS_REG", cmd, SqlS, mydebuglog, "N")
                    Else
                        GoTo AccessError
                    End If
                End If
                
                ' Check from the CX_SESS_PART_X record in case there is a disconnect
                If SESS_PART_ID = "" Or EXAM_ID = "" Or JURIS_ID = "" Or CRSE_ID = "" Or TST_ID = "" Then
                    SqlS = "SELECT SP.ROW_ID, SP.CRSE_TSTRUN_ID, SP.JURIS_ID, SP.CRSE_TST_ID, TR.CRSE_TST_ID " & _
                             "FROM siebeldb.dbo.CX_SESS_PART_X SP " & _
                             "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TSTRUN TR ON TR.ROW_ID=SP.CRSE_TSTRUN_ID " & _
                             "WHERE SP.REG_ID='" & REG_ID & "'"
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  LOCATE SESS_PART_ID: " & vbCrLf & "  " & SqlS)
                    cmd.CommandText = SqlS
                    dr = cmd.ExecuteReader()
                    If Not dr Is Nothing Then
                        While dr.Read()
                            Try
                                If SESS_PART_ID = "" Then SESS_PART_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                                If EXAM_ID = "" Then EXAM_ID = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                                If JURIS_ID = "" Then JURIS_ID = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                                If CRSE_ID = "" Then CRSE_ID = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                                If TST_ID = "" Then TST_ID = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                            Catch ex As Exception
                                MSG1 = ex.Message
                                GoTo DBError
                            End Try
                        End While
                    End If
                    dr.Close()
                End If
                If Debug = "Y" Then
                    mydebuglog.Debug("  .. SESS_PART_ID: " & SESS_PART_ID)
                    mydebuglog.Debug("  .. EXAM_ID: " & EXAM_ID)
                    mydebuglog.Debug("  .. JURIS_ID: " & JURIS_ID)
                    mydebuglog.Debug("  .. ALT_JURIS_ID: " & ALT_JURIS_ID)
                    mydebuglog.Debug("  .. CRSE_JURIS_ID: " & CRSE_JURIS_ID)
                    mydebuglog.Debug("  .. CRSE_ID: " & CRSE_ID)
                    mydebuglog.Debug("  .. TST_ID: " & TST_ID)
                    mydebuglog.Debug("  .. RETAKE_FLG: " & RETAKE_FLG)
                    mydebuglog.Debug("  .. ALT_RETAKE_FLG: " & ALT_RETAKE_FLG)
                    mydebuglog.Debug("  .. TST_RETAKE_FLG: " & TST_RETAKE_FLG)
                    mydebuglog.Debug("  .. SKILL_LEVEL_CD: " & SKILL_LEVEL_CD)
                    mydebuglog.Debug("  .. LANG_ID: " & LANG_ID)
                    mydebuglog.Debug("  .. LANG: " & LANG)
                    mydebuglog.Debug("  .. REG_STATUS_CD: " & REG_STATUS_CD)
                    mydebuglog.Debug("  .. RESOLUTION: " & RESOLUTION)
                    mydebuglog.Debug("  .. X_SURVEY_FLG: " & X_SURVEY_FLG)
                    mydebuglog.Debug("  .. CONTINUABLE_FLG: " & CONTINUABLE_FLG)
                    mydebuglog.Debug("  .. KBA_QUES_NUM: " & KBA_QUES_NUM)
                    mydebuglog.Debug("  .. USER_NAME: " & USER_NAME)
                    mydebuglog.Debug("  .. CRSE_CONTENT_URL: " & CRSE_CONTENT_URL)
                    mydebuglog.Debug("  .. EMAIL_ADDR: " & EMAIL_ADDR)
                End If
                
                ' Double check to see if we already have a test out there but it's not linked from the participation record
                ' but is linked from the testrun record
                If EXAM_ID = "" And SESS_PART_ID <> "" Then
                    SqlS = "SELECT ROW_ID, CRSE_TST_ID FROM siebeldb.dbo.S_CRSE_TSTRUN WHERE X_PART_ID='" & SESS_PART_ID & "'"
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  CHECK FOR MISLINKED EXAM_ID: " & vbCrLf & "  " & SqlS)
                    cmd.CommandText = SqlS
                    dr = cmd.ExecuteReader()
                    If Not dr Is Nothing Then
                        While dr.Read()
                            Try
                                EXAM_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                                TST_ID = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            Catch ex As Exception
                            End Try
                        End While
                    End If
                    dr.Close()
                    
                    ' Update CX_SESS_PART_X link to S_CRSE_TSTRUN   
                    If EXAM_ID <> "" Then
                        If SESS_PART_ID <> "" Then
                            SqlS = "UPDATE siebeldb.dbo.CX_SESS_PART_X SET CRSE_TSTRUN_ID='" & EXAM_ID & "' WHERE ROW_ID='" & SESS_PART_ID & "'"
                            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Update CX_SESS_PART_X link: " & vbCrLf & "  " & SqlS)
                            temp = ExecQuery("Update", "CX_SESS_PART_X", cmd, SqlS, mydebuglog, "N")
                        End If
                    End If
                End If
                
            Else
                ' See if we have a registration for this exam anyway
                If EXAM_ID <> "" Then
                    SqlS = "SELECT R.ROW_ID, TR.ROW_ID, R.JURIS_ID, T.CRSE_ID, TR.CRSE_TST_ID, R.STATUS_CD, " & _
                             "T.X_SURVEY_FLG, T.X_RESOLUTION, T.X_ANON_REF_ID, TR.PERSON_ID, R.TEST_FLG, TRX.ROW_ID, T.X_CONTINUABLE_FLG, " & _
                             "(SELECT CASE WHEN TC.LANG_ID IS NULL OR TC.LANG_ID='' THEN CR.X_LANG_CD ELSE TC.LANG_ID END) AS LANG_ID, L.NAME, " & _
                             "T.X_KBA_QUES_NUM, T.X_FORMAT, C.FST_NAME+' '+C.LAST_NAME, C.EMAIL_ADDR, T.SKILL_LEVEL_CD, CR.X_CRSE_CONTENT_URL, " & _
                             "R.RETAKE_FLG, CR.X_ALT_RETAKE_FLG, T.X_RETAKE_FLG, SP.ROW_ID, T.X_CATEGORY, CR.TYPE_CD, TR.STATUS_CD " & _
                             "FROM siebeldb.dbo.S_CRSE_TSTRUN TR " & _
                             "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=TR.PERSON_ID " & _
                             "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TSTRUN_X TRX ON TRX.PAR_ROW_ID=TR.ROW_ID " & _
                             "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TST T ON T.ROW_ID=TR.CRSE_TST_ID " & _
                             "LEFT OUTER JOIN siebeldb.dbo.S_LANG L ON L.LANG_CD=T.X_LANG_ID " & _
                             "LEFT OUTER JOIN siebeldb.dbo.CX_SESS_PART_X SP ON SP.CRSE_TSTRUN_ID=TR.ROW_ID " & _
                             "LEFT OUTER JOIN siebeldb.dbo.CX_SESS_REG R ON R.ROW_ID=SP.REG_ID " & _
                             "LEFT OUTER JOIN siebeldb.dbo.CX_TRAIN_OFFR TC ON TC.ROW_ID=R.TRAIN_OFFR_ID " & _
                             "LEFT OUTER JOIN siebeldb.dbo.S_CRSE CR ON CR.ROW_ID=T.CRSE_ID " & _
                             "WHERE TR.ROW_ID='" & EXAM_ID & "'"
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  CHECK FOR EXAM W/EXAM_ID: " & vbCrLf & "  " & SqlS)
                    cmd.CommandText = SqlS
                    dr = cmd.ExecuteReader()
                    If Not dr Is Nothing Then
                        While dr.Read()
                            Try                                
                                temp = ""
                                REG_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                                EXAM_ID = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                                JURIS_ID = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                                CRSE_ID = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                                temp = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                                If temp <> "" Then TST_ID = temp
                                REG_STATUS_CD = Trim(CheckDBNull(dr(5), enumObjectType.StrType))
                                X_SURVEY_FLG = Trim(CheckDBNull(dr(6), enumObjectType.StrType))
                                RESOLUTION = Trim(CheckDBNull(dr(7), enumObjectType.StrType))
                                ANON_REF_ID = Trim(CheckDBNull(dr(8), enumObjectType.StrType))
                                CON_ID = Trim(CheckDBNull(dr(9), enumObjectType.StrType))
                                TEST_FLG = Trim(CheckDBNull(dr(10), enumObjectType.StrType))
                                TRX_ID = Trim(CheckDBNull(dr(11), enumObjectType.StrType))
                                CONTINUABLE_FLG = Trim(CheckDBNull(dr(12), enumObjectType.StrType))
                                LANG_ID = Trim(CheckDBNull(dr(13), enumObjectType.IntType))
                                LANG = Trim(CheckDBNull(dr(14), enumObjectType.StrType))
                                KBA_QUES_NUM = Trim(Str(CheckDBNull(dr(15), enumObjectType.StrType)))
                                X_FORMAT = Trim(CheckDBNull(dr(16), enumObjectType.StrType))
                                USER_NAME = Trim(CheckDBNull(dr(17), enumObjectType.StrType))
                                EMAIL_ADDR = Trim(CheckDBNull(dr(18), enumObjectType.StrType))
                                SKILL_LEVEL_CD = Trim(CheckDBNull(dr(19), enumObjectType.StrType))
                                CRSE_CONTENT_URL = Trim(CheckDBNull(dr(20), enumObjectType.StrType))
                                RETAKE_FLG = Trim(CheckDBNull(dr(21), enumObjectType.StrType))
                                ALT_RETAKE_FLG = Trim(CheckDBNull(dr(22), enumObjectType.StrType))
                                TST_RETAKE_FLG = Trim(CheckDBNull(dr(23), enumObjectType.StrType))
                                SESS_PART_ID = Trim(CheckDBNull(dr(24), enumObjectType.StrType))
                                CATEGORY = Trim(CheckDBNull(dr(25), enumObjectType.StrType))
                                If CATEGORY = "0" Then CATEGORY = ""
                                COURSE_TYPE_CD = Trim(CheckDBNull(dr(26), enumObjectType.StrType))
                                TEST_STATUS_CD = Trim(CheckDBNull(dr(27), enumObjectType.StrType))                                
                            Catch ex As Exception
                                MSG1 = ex.Message
                                GoTo DBError
                            End Try
                        End While
                    Else
                        GoTo AccessError
                    End If
                    dr.Close()
                    
                    If Debug = "Y" Then
                        mydebuglog.Debug("  .. REG_ID: " & REG_ID)
                        mydebuglog.Debug("  .. SESS_PART_ID: " & SESS_PART_ID)
                        mydebuglog.Debug("  .. EXAM_ID: " & EXAM_ID)
                        mydebuglog.Debug("  .. JURIS_ID: " & JURIS_ID)
                        mydebuglog.Debug("  .. CRSE_ID: " & CRSE_ID)
                        mydebuglog.Debug("  .. TST_ID: " & TST_ID)
                        mydebuglog.Debug("  .. RETAKE_FLG: " & RETAKE_FLG)
                        mydebuglog.Debug("  .. ALT_RETAKE_FLG: " & ALT_RETAKE_FLG)
                        mydebuglog.Debug("  .. TST_RETAKE_FLG: " & TST_RETAKE_FLG)
                        mydebuglog.Debug("  .. CON_ID: " & CON_ID)
                        mydebuglog.Debug("  .. LANG_ID: " & LANG_ID)
                        mydebuglog.Debug("  .. LANG: " & LANG)
                        mydebuglog.Debug("  .. REG_STATUS_CD: " & REG_STATUS_CD)
                        mydebuglog.Debug("  .. X_SURVEY_FLG: " & X_SURVEY_FLG)
                        mydebuglog.Debug("  .. ANON_REF_ID: " & ANON_REF_ID)
                        mydebuglog.Debug("  .. SKILL_LEVEL_CD: " & SKILL_LEVEL_CD)
                        mydebuglog.Debug("  .. CATEGORY: " & CATEGORY)
                        mydebuglog.Debug("  .. CONTINUABLE_FLG: " & CONTINUABLE_FLG)
                        mydebuglog.Debug("  .. TEST_FLG: " & TEST_FLG)
                        mydebuglog.Debug("  .. X_FORMAT: " & X_FORMAT)
                        mydebuglog.Debug("  .. KBA_QUES_NUM: " & KBA_QUES_NUM)
                        mydebuglog.Debug("  .. USER_NAME: " & USER_NAME)
                        mydebuglog.Debug("  .. CRSE_CONTENT_URL: " & CRSE_CONTENT_URL)
                        mydebuglog.Debug("  .. EMAIL_ADDR: " & EMAIL_ADDR)
                        mydebuglog.Debug("  .. COURSE_TYPE_CD: " & COURSE_TYPE_CD)
                        mydebuglog.Debug("  .. TEST_STATUS_CD: " & TEST_STATUS_CD & vbCrLf)
                    End If
                End If
            End If
                       
            ' Set SSLFlag based on course string
            If InStr(CRSE_CONTENT_URL, "https:") > 0 Then
                LaunchProtocol = "https:"
            End If
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  LaunchProtocol URL: " & LaunchProtocol)
            
            ' -----
            ' In registration and no participation record, create it
            If REG_ID <> "" And X_SURVEY_FLG <> "Y" And SESS_PART_ID = "" Then
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Generating participation record for REG_ID: " & REG_ID)
                SESS_PART_ID = Processing.GenerateParticipation(REG_ID, Debug)
                If Debug = "Y" Then mydebuglog.Debug("  .. SESS_PART_ID: " & SESS_PART_ID)
                If SESS_PART_ID = "" Then GoTo DataError
            End If
            
            ' -----
            ' Locate TST_ID if it is blank - we need this always - by calling GetCourseExam
            If TST_ID = "" And CRSE_ID <> "" Then
                If ALT_RETAKE_FLG = "Y" And RETAKE_FLG = "Y" And TST_RETAKE_FLG = "N" Then
                    If LANG_ID = "" Then LANG_ID = "ENU"
                    SqlS = "SELECT TOP 1 ROW_ID " & _
                             "FROM siebeldb.dbo.S_CRSE_TST " & _
                             "WHERE CRSE_ID='" & CRSE_ID & "' AND X_JURIS_ID='" & JURIS_ID & "' AND X_LANG_ID='" & LANG_ID & "' AND STATUS_CD='Retake' " & _
                             "ORDER BY X_VERSION DESC"
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  RETRIEVE ALT_RETAKE_FLG TEST: " & vbCrLf & "  " & SqlS)
                    cmd.CommandText = SqlS
                    Try
                        TST_ID = CheckDBNull(cmd.ExecuteScalar(), enumObjectType.StrType)
                    Catch ex As Exception
                        MSG1 = ex.Message
                        GoTo DBError
                    End Try
                    If Debug = "Y" Then mydebuglog.Debug("  .. TST_ID: " & TST_ID)
                End If
                
                If TST_ID = "" Then
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Calling GetCourseExam: CRSE_ID: " & CRSE_ID & " - JURIS_ID: " & JURIS_ID & " - LANG_ID: " & LANG_ID)
                    TST_ID = Scormsvc.GetCourseExam(CRSE_ID, JURIS_ID, LANG_ID, Debug)
                    ' If not found - try again with an english exam
                    If TST_ID = "" And LANG_ID <> "ENU" Then
                        LANG_ID = "ENU"
                        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Calling GetCourseExam: CRSE_ID: " & CRSE_ID & " - JURIS_ID: " & JURIS_ID & " - LANG_ID: " & LANG_ID)
                        TST_ID = Scormsvc.GetCourseExam(CRSE_ID, JURIS_ID, LANG_ID, Debug)
                    End If
                    If Debug = "Y" Then mydebuglog.Debug("  .. TST_ID: " & TST_ID)
                    If TST_ID = "" Or TST_ID = "Error" Then GoTo ExamError
                End If
                
                ' Get test information
                SqlS = "SELECT T.CRSE_ID, T.X_SURVEY_FLG, T.X_RESOLUTION, T.X_ANON_REF_ID, T.X_CONTINUABLE_FLG, " & _
                "(SELECT CASE WHEN T.X_LANG_ID IS NULL OR T.X_LANG_ID='' THEN CR.X_LANG_CD ELSE T.X_LANG_ID END) AS LANG_ID, L.NAME, " & _
                "T.X_KBA_QUES_NUM, T.X_FORMAT " & _
                "FROM siebeldb.dbo.S_CRSE_TST T " & _
                "LEFT OUTER JOIN siebeldb.dbo.S_LANG L ON L.LANG_CD=T.X_LANG_ID " & _
                "LEFT OUTER JOIN siebeldb.dbo.S_CRSE CR ON CR.ROW_ID=T.CRSE_ID " & _
                "WHERE T.ROW_ID='" & TST_ID & "'"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  CHECK FOR EXAM W/TST_ID: " & vbCrLf & "  " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            CRSE_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            X_SURVEY_FLG = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            RESOLUTION = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                            ANON_REF_ID = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                            CONTINUABLE_FLG = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                            LANG_ID = Trim(CheckDBNull(dr(5), enumObjectType.StrType))
                            LANG = Trim(CheckDBNull(dr(6), enumObjectType.StrType))
                            KBA_QUES_NUM = Trim(Str(CheckDBNull(dr(7), enumObjectType.StrType)))
                            X_FORMAT = Trim(CheckDBNull(dr(8), enumObjectType.StrType))
                        Catch ex As Exception
                            MSG1 = ex.Message
                            GoTo DBError
                        End Try
                    End While
                Else
                    GoTo AccessError
                End If
                dr.Close()
                    
                If Debug = "Y" Then
                    mydebuglog.Debug("  .. CRSE_ID: " & CRSE_ID)
                    mydebuglog.Debug("  .. X_SURVEY_FLG: " & X_SURVEY_FLG)
                    mydebuglog.Debug("  .. RESOLUTION: " & RESOLUTION)
                    mydebuglog.Debug("  .. ANON_REF_ID: " & ANON_REF_ID)
                    mydebuglog.Debug("  .. CONTINUABLE_FLG: " & CONTINUABLE_FLG)
                    mydebuglog.Debug("  .. LANG_ID: " & LANG_ID)
                    mydebuglog.Debug("  .. LANG: " & LANG)
                    mydebuglog.Debug("  .. KBA_QUES_NUM: " & KBA_QUES_NUM)
                    mydebuglog.Debug("  .. X_FORMAT: " & X_FORMAT)
                End If
            End If
            
            ' -----
            ' Prepare refresh/next link
            If Trim(EXAM_ID) <> "" Then
                If LANG_CD <> "ENU" Then
                    NextLink = LaunchProtocol & "//hciscorm.certegrity.com/ls/" & LANG_CD & "/FinishAssessment.html?ID=" & EXAM_ID & "&UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "&LANG=" & LANG_CD & "&HP=" & HOME_PAGE & "&RFR=" & RefreshID & "&CUR=" & CURRENT_PAGE
                Else
                    NextLink = LaunchProtocol & "//hciscorm.certegrity.com/ls/FinishAssessment.html?ID=" & EXAM_ID & "&UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "&HP=" & HOME_PAGE & "&RFR=" & RefreshID & "&CUR=" & CURRENT_PAGE
                End If
                RefreshLink = "window.opener.document.location='" & NextLink & "';"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  NextLink: " & NextLink)
            End If
            
            ' ================================================
            ' TRY TO LOCATE PENDING EXAM ID FOR THE SAME EXAM TO AVOID DUPING IT
            If TST_ID <> "" And EXAM_ID = "" And CON_ID <> "" Then
                SqlS = "SELECT ROW_ID " & _
                      "FROM siebeldb.dbo.S_CRSE_TSTRUN " & _
                      "WHERE PERSON_ID='" & CON_ID & "' AND STATUS_CD='Pending' AND CRSE_TST_ID='" & TST_ID & "'"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  CHECKING FOR EXISTING EXAM: " & vbCrLf & "  " & SqlS)
                cmd.CommandText = SqlS
                Try
                    EXAM_ID = CheckDBNull(cmd.ExecuteScalar(), enumObjectType.StrType)
                Catch ex As Exception
                End Try
                If Debug = "Y" Then mydebuglog.Debug("  .. EXAM_ID : " & EXAM_ID)
            End If
        
            ' ================================================
            ' VERIFY SPECIFIED EXAM_ID
            ' Get information about a specified exam
            If EXAM_ID <> "" Then
                SqlS = "SELECT PERSON_ID, STATUS_CD FROM siebeldb.dbo.S_CRSE_TSTRUN " & _
                "WHERE ROW_ID='" & EXAM_ID & "'"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  CHECK FOR MISLINKED EXAM_ID: " & vbCrLf & "  " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            PERSON_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            EXAM_STATUS_CD = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                        Catch ex As Exception
                        End Try
                    End While
                End If
                dr.Close()
                If Debug = "Y" Then
                    mydebuglog.Debug("  .. PERSON_ID: " & PERSON_ID)
                    mydebuglog.Debug("  .. EXAM_STATUS_CD: " & EXAM_STATUS_CD)
                End If
            End If
        
            ' If did not find an exam, create it
            If EXAM_ID = "" Then
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Calling GenerateNewExam: TST_ID: " & TST_ID & " - SESS_PART_ID: " & SESS_PART_ID & " - CONTACT_ID: " & CONTACT_ID)
                EXAM_ID = Processing.GenerateNewExam(TST_ID, EXAM_ID, CONTACT_ID, "N", SESS_PART_ID, Debug)
                If Debug = "Y" Then mydebuglog.Debug("  .. EXAM_ID : " & EXAM_ID)
                EXAM_STATUS_CD = "Pending"
                If EXAM_ID <> "" Then
                    ' Update CX_SESS_PART_X link to S_CRSE_TSTRUN   
                    If SESS_PART_ID <> "" Then
                        SqlS = "UPDATE siebeldb.dbo.CX_SESS_PART_X SET CRSE_TSTRUN_ID='" & EXAM_ID & "' WHERE ROW_ID='" & SESS_PART_ID & "'"
                        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Update CX_SESS_PART_X link: " & vbCrLf & "  " & SqlS)
                        temp = ExecQuery("Update", "CX_SESS_PART_X", cmd, SqlS, mydebuglog, "N")
                        If REG_ID <> "" Then
                            SqlS = "UPDATE siebeldb.dbo.CX_SESS_REG SET SESS_PART_ID='" & SESS_PART_ID & "' WHERE ROW_ID='" & REG_ID & "'"
                            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Update CX_SESS_REG link: " & vbCrLf & "  " & SqlS)
                            temp = ExecQuery("Update", "CX_SESS_REG", cmd, SqlS, mydebuglog, "N")
                        End If
                    End If
                
                    ' Re-retrieve exam parameters for KBA
                    If REG_ID <> "" Then
                        SqlS = "SELECT SP.ROW_ID, SP.CRSE_TSTRUN_ID, R.JURIS_ID, R.CRSE_ID, TR.CRSE_TST_ID, " & _
                                    "R.STATUS_CD, T.X_SURVEY_FLG, T.X_RESOLUTION, R.TEST_FLG, TRX.ROW_ID, T.X_CONTINUABLE_FLG, " & _
                                    "(SELECT CASE WHEN TC.LANG_ID IS NULL THEN CR.X_LANG_CD ELSE TC.LANG_ID END) AS LANG_ID, L.NAME, " & _
                                    "T.X_KBA_QUES_NUM, (SELECT CASE WHEN A.X_JURIS_ID IS NULL THEN PA.X_JURIS_ID ELSE A.X_JURIS_ID END) AS ALT_JURIS_ID, " & _
                                    "CR.X_JURIS_ID, R.CONTACT_ID, C.FST_NAME+' '+C.LAST_NAME, C.EMAIL_ADDR, CR.X_FORMAT, T.SKILL_LEVEL_CD " & _
                                    "FROM siebeldb.dbo.CX_SESS_REG R " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=R.CONTACT_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_ADDR_ORG A ON A.ROW_ID=R.ADDR_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_ADDR_PER PA ON PA.ROW_ID=R.PER_ADDR_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.CX_TRAIN_OFFR TC ON TC.ROW_ID=R.TRAIN_OFFR_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_CRSE CR ON CR.ROW_ID=R.CRSE_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.CX_SESS_PART_X SP ON SP.ROW_ID=R.SESS_PART_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TSTRUN TR ON TR.ROW_ID=SP.CRSE_TSTRUN_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TSTRUN_X TRX ON TRX.PAR_ROW_ID=TR.ROW_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TST T ON T.ROW_ID=TR.CRSE_TST_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_LANG L ON L.LANG_CD=T.X_LANG_ID " & _
                                    "WHERE R.ROW_ID='" & REG_ID & "'"
                        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  RE-RETRIEVE EXAM W/REG_ID: " & vbCrLf & "  " & SqlS)
                        cmd.CommandText = SqlS
                        dr = cmd.ExecuteReader()
                        If Not dr Is Nothing Then
                            While dr.Read()
                                Try
                                    temp = ""
                                    SESS_PART_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                                    EXAM_ID = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                                    JURIS_ID = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                                    CRSE_ID = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                                    TST_ID = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                                    REG_STATUS_CD = Trim(CheckDBNull(dr(5), enumObjectType.StrType))
                                    X_SURVEY_FLG = Trim(CheckDBNull(dr(6), enumObjectType.StrType))
                                    RESOLUTION = Trim(CheckDBNull(dr(7), enumObjectType.StrType))
                                    TEST_FLG = Trim(CheckDBNull(dr(8), enumObjectType.StrType))
                                    TRX_ID = Trim(CheckDBNull(dr(9), enumObjectType.StrType))
                                    CONTINUABLE_FLG = Trim(CheckDBNull(dr(10), enumObjectType.StrType))
                                    LANG_ID = Trim(CheckDBNull(dr(11), enumObjectType.StrType))
                                    LANG = Trim(CheckDBNull(dr(12), enumObjectType.StrType))
                                    KBA_QUES_NUM = Trim(Str(CheckDBNull(dr(13), enumObjectType.IntType)))
                                    ALT_JURIS_ID = Trim(CheckDBNull(dr(14), enumObjectType.StrType))
                                    CRSE_JURIS_ID = Trim(CheckDBNull(dr(15), enumObjectType.StrType))
                                    PERSON_ID = Trim(CheckDBNull(dr(16), enumObjectType.StrType))
                                    USER_NAME = Trim(CheckDBNull(dr(17), enumObjectType.StrType))
                                    EMAIL_ADDR = Trim(CheckDBNull(dr(18), enumObjectType.StrType))
                                    X_FORMAT = Trim(CheckDBNull(dr(19), enumObjectType.StrType))
                                    SKILL_LEVEL_CD = Trim(CheckDBNull(dr(20), enumObjectType.StrType))
                                Catch ex As Exception
                                    MSG1 = ex.Message
                                    GoTo DBError
                                End Try
                            End While
                        Else
                            GoTo ExamError
                        End If
                        dr.Close()
                        If Debug = "Y" Then
                            mydebuglog.Debug("  .. SESS_PART_ID: " & SESS_PART_ID)
                            mydebuglog.Debug("  .. EXAM_ID: " & EXAM_ID)
                            mydebuglog.Debug("  .. JURIS_ID: " & JURIS_ID)
                            mydebuglog.Debug("  .. CRSE_ID: " & CRSE_ID)
                            mydebuglog.Debug("  .. TST_ID: " & TST_ID)
                            mydebuglog.Debug("  .. LANG_ID: " & LANG_ID)
                            mydebuglog.Debug("  .. LANG: " & LANG)
                            mydebuglog.Debug("  .. REG_STATUS_CD: " & REG_STATUS_CD)
                            mydebuglog.Debug("  .. RESOLUTION: " & RESOLUTION)
                            mydebuglog.Debug("  .. X_SURVEY_FLG: " & X_SURVEY_FLG)
                            mydebuglog.Debug("  .. CONTINUABLE_FLG: " & CONTINUABLE_FLG)
                            mydebuglog.Debug("  .. X_FORMAT: " & X_FORMAT)
                            mydebuglog.Debug("  .. SKILL_LEVEL_CD: " & SKILL_LEVEL_CD)
                            mydebuglog.Debug("  .. KBA_QUES_NUM: " & KBA_QUES_NUM)
                            mydebuglog.Debug("  .. USER_NAME: " & USER_NAME)
                            mydebuglog.Debug("  .. EMAIL_ADDR: " & EMAIL_ADDR)
                        End If
                    End If
                Else
                    GoTo ExamError
                End If
            End If
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Person Id: " & PERSON_ID & ", Status: " & EXAM_STATUS_CD)
        
            ' If exam not found or exam not for user, then reset EXAM_ID to force a new one to be created
            If PERSON_ID = "" Or EXAM_ID = "" Then GoTo ExamError
        
            ' If the exam is already completed, exit this function
            If EXAM_STATUS_CD = "Graded" And EXAM_ID <> "" Then
                If X_SURVEY_FLG = "Y" Then
                    If LANG_CD = "ESN" Then
                        MSG1 = "<span class=""BigHeader""><font color=""Gray"">Ya has completado la encuesta. Gracias.<br><br>"
                    Else
                        MSG1 = "<span class=""BigHeader""><font color=""Gray"">You have completed the survey already.  Thank you.<br><br>"
                    End If
                Else
                    If LANG_CD = "ESN" Then
                        MSG1 = "<span class=""BigHeader""><font color=""Gray"">Ya has completado tu examen de certificaci&oacute;n.<br><br>"
                        btext = "Haga clic para revisar sus resultados </a> en cualquier momento."
                    Else
                        MSG1 = "<span class=""BigHeader""><font color=""Gray"">You have completed your certification exam already.<br><br>"
                        btext = "click to review your results</a> at any time."
                    End If
                    If REG_ID <> "" Then
                        ClassLink = "https://hciscorm.certegrity.com/ls/OpenCertificate.html?RID=" & REG_ID & "&UID=" & UID & "&SES=" & SessID & "&HP=" & HOME_PAGE & "&PP=" & DOMAIN
                        MSG1 = MSG1 & "<center>You may <a href=""https://hciscorm.certegrity.com/ls/OpenCertificate.html?RID=" & REG_ID & "&UID=" & UID & "&SES=" & SessID & "&HP=" & HOME_PAGE & "&CUR=" & CURRENT_PAGE & """ class=""button"">" & btext & "</center>"
                    Else
                        ClassLink = "https://hciscorm.certegrity.com/ls/OpenCertificate.html?TID=" & EXAM_ID & "&UID=" & UID & "&SES=" & SessID & "&HP=" & HOME_PAGE & "&PP=" & DOMAIN
                        MSG1 = MSG1 & "<center>You may <a href=""https://hciscorm.certegrity.com/ls/OpenCertificate.html?TID=" & EXAM_ID & "&UID=" & UID & "&SES=" & SessID & "&HP=" & HOME_PAGE & "&CUR=" & CURRENT_PAGE & """ class=""button"">" & btext & "</center>"
                    End If
                End If
                delay = "0"
                GoTo ReturnControl
            End If

            ' Get access semaphor count
            Dim accessfound As Integer
            SqlS = "SELECT COUNT(*) FROM siebeldb.dbo.S_CRSE_TSTRUN_ACCESS WHERE CRSE_TSTRUN_ID='" & EXAM_ID & "'"
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Get access semaphor: " & vbCrLf & "  " & SqlS)
            cmd.CommandText = SqlS
            dr = cmd.ExecuteReader()
            If Not dr Is Nothing Then
                While dr.Read()
                    Try
                        accessfound = CheckDBNull(dr(0), enumObjectType.IntType)
                        If Debug = "Y" Then mydebuglog.Debug("  .. accessfound: " & Str(accessfound))
                    Catch ex As Exception
                    End Try
                End While
            Else
                GoTo DataError
            End If
            dr.Close()
                
            ' If out-of-control access, then lockout
            If lockoutcount > 0 And accessfound >= lockoutcount Then
                If REG_ID <> "" Then
                    SqlS = "UPDATE siebeldb.dbo.CX_SESS_REG SET STATUS_CD='On-Hold' WHERE ROW_ID='" & REG_ID & "'"
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Locking registration out because access count exceeded: " & vbCrLf & "  " & SqlS)
                    temp = ExecQuery("Update", "CX_SESS_REG", cmd, SqlS, mydebuglog, "N")
                Else
                    SqlS = "UPDATE siebeldb.dbo.S_CRSE_TSTRUN SET STATUS_CD='On-Hold' WHERE ROW_ID='" & EXAM_ID & "'"
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Locking assessment out because access count exceeded: " & vbCrLf & "  " & SqlS)
                    temp = ExecQuery("Update", "S_CRSE_TSTRUN", cmd, SqlS, mydebuglog, "N")
                End If
                GoTo AccessCountError
            End If
            
            ' Check to see if on hold
            If TEST_STATUS_CD = "On-Hold" Then
                GoTo OnHoldError
            End If

            ' Is this a rerun?  
            ' If exam not completed, verify access
            SqlS = "SELECT TOP 1 ENTER_FLG, EXIT_FLG FROM siebeldb.dbo.S_CRSE_TSTRUN_ACCESS " & _
            "WHERE CRSE_TSTRUN_ID='" & EXAM_ID & "' ORDER BY CREATED DESC "
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  ACCESS QUERY: " & vbCrLf & "  " & SqlS)
            cmd.CommandText = SqlS
            dr = cmd.ExecuteReader()
            If Not dr Is Nothing Then
                While dr.Read()
                    Try
                        ENTER_FLG = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                        EXIT_FLG = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                    Catch ex As Exception
                    End Try
                End While
            End If
            dr.Close()
            If Debug = "Y" Then
                mydebuglog.Debug("  .. ENTER_FLG : " & ENTER_FLG)
                mydebuglog.Debug("  .. EXIT_FLG : " & EXIT_FLG)
            End If

            ' If the enter state is already set, then do not restart the exam - it is open in another browser window
            If ENTER_FLG = "Y" Then
                SqlS = "SELECT COUNT(*) " & _
                "FROM elearning.dbo.Elearning_Player_Data " & _
                "WHERE reg_id='" & REG_ID & "' AND crse_id='" & TST_ID & "'"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  SCORM DB QUERY: " & vbCrLf & "  " & SqlS)
                cmd.CommandText = SqlS
                Try
                    ACTIVITY_COUNT = Trim(CheckDBNull(cmd.ExecuteScalar(), enumObjectType.IntType))
                Catch ex As Exception
                End Try
                If Debug = "Y" Then mydebuglog.Debug("  .. ACTIVITY_COUNT  : " & Str(ACTIVITY_COUNT))
                
                If ACTIVITY_COUNT > 0 Then
                    SqlS = "SELECT RTRIM(CAST(active AS CHAR))+RTRIM(CAST(suspended AS CHAR)) " & _
                             "FROM elearning.dbo.Elearning_Player_Data " & _
                             "WHERE reg_id='" & REG_ID & "'"
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  SCORM CHECK FLAGS QUERY: " & vbCrLf & "  " & SqlS)
                    cmd.CommandText = SqlS
                    Try
                        SCORM_ACTIVITY = Trim(CheckDBNull(cmd.ExecuteScalar(), enumObjectType.StrType))
                    Catch ex As Exception
                    End Try
                    If Debug = "Y" Then
                        mydebuglog.Debug("  .. SCORM_ACTIVITY : " & Str(SCORM_ACTIVITY))
                    End If
                    If SCORM_ACTIVITY = "10" Then
                        If LANG_CD = "ESN" Then
                            MSG1 = "<table width=""100%"" height=""100%""><tr valign=""Middle""><td class=""BigHeader"">" & _
                                       "<font color=""Gray"">Ya tienes este examen abierto en otra ventana del navegador. <br><br>" & _
                                       "</td></tr></table>"
                        Else
                            MSG1 = "<table width=""100%"" height=""100%""><tr valign=""Middle""><td class=""BigHeader"">" & _
                                       "<font color=""Gray"">You have this exam already open in another browser window. <br><br>" & _
                                       "</td></tr></table>"
                        End If
                        MSG1 = "<table width=""100%"" height=""100%""><tr valign=""Middle""><td class=""BigHeader"">" & _
                                   "<font color=""Gray"">You have this exam already open in another browser window. <br><br>" & _
                                   "</td></tr></table>"
                        GoTo ReturnControl
                    End If
                End If
            End If
                  
            ' ================================================
            ' Prepare refresh/next link
            If Trim(EXAM_ID) <> "" Then
                If LANG_CD <> "ENU" Then
                    NextLink = LaunchProtocol & "//hciscorm.certegrity.com/ls/" & LANG_CD & "/FinishAssessment.html?ID=" & EXAM_ID & "&UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "&LANG=" & LANG_CD & "&HP=" & HOME_PAGE & "&RFR=" & RefreshID & "&CUR=" & CURRENT_PAGE
                Else
                    NextLink = LaunchProtocol & "//hciscorm.certegrity.com/ls/FinishAssessment.html?ID=" & EXAM_ID & "&UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "&HP=" & HOME_PAGE & "&RFR=" & RefreshID & "&CUR=" & CURRENT_PAGE
                End If
                RefreshLink = "window.opener.document.location='" & NextLink & "';"
            End If
            
            ' ================================================
            ' VERIFY SUB_CON_ID - NEEDED FOR REDIRECT BACK
            If CONTACT_ID <> "" Then
                If NextLink = "" Then
                    If LANG_CD <> "ENU" Then
                        MSG1 = "<table width=""100%"" height=""100%""><tr valign=""Middle""><td class=""BigHeader"">" & _
                        "<font color=""Gray"">No podemos identificar su examen. <br><br>" & _
                        "</td></tr></table>"
                    Else
                        MSG1 = "<table width=""100%"" height=""100%""><tr valign=""Middle""><td class=""BigHeader"">" & _
                        "<font color=""Gray"">We cannot identify your exam. <br><br>" & _
                        "</td></tr></table>"
                    End If
                    GoTo ReturnControl
                End If
                LAST_INST = NextLink
                SqlS = "UPDATE siebeldb.dbo.CX_SUB_CON SET LAST_INST='" & NextLink & "' " & _
                "WHERE CON_ID='" & CONTACT_ID & "'"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  UPDATING LAST_INST: " & vbCrLf & "  " & SqlS)
                temp = ExecQuery("Update", "CX_SUB_CON", cmd, SqlS, mydebuglog, "N")
            End If
            
            ' ================================================
            ' CHECK FOR FINISHED CLASS BASED ON PARAMETER
            If FINISH_CLASS = "Y" And REG_STATUS_CD = "In Progress" Then
                SqlS = "UPDATE siebeldb.dbo.CX_SESS_REG " & _
                      "SET STATUS_CD='Exam Reqd', PROGRESS='100' " & _
                      "WHERE ROW_ID='" & REG_ID & "'"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  UPDATING REG STATUS: " & vbCrLf & "  " & SqlS)
                temp = ExecQuery("Update", "CX_SESS_REG", cmd, SqlS, mydebuglog, "N")
                REG_STATUS_CD = "Exam Reqd"
            End If
            
            ' ================================================
            ' PROCESS PARAMETER CASES
            ' CASE: Online class (REG_ID<>"")
            If REG_ID <> "" Then
                ' If no exam, create one and an associated participation record if necessary
                If EXAM_ID = "" Then
                    ' Create a session participation record if it does not exist 
                    If SESS_PART_ID = "" Then
                        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Calling GenerateParticipation: REG_ID: " & REG_ID)
                        SESS_PART_ID = Processing.GenerateParticipation(REG_ID, Debug)
                        If Debug = "Y" Then mydebuglog.Debug("  .. SESS_PART_ID : " & SESS_PART_ID)
                        If SESS_PART_ID = "" Then GoTo DataError
                    End If
         
                    ' Create an instance of an exam for the test - assume a participant
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Calling GenerateNewExam: TST_ID: " & TST_ID & " - SESS_PART_ID: " & SESS_PART_ID & " - CONTACT_ID: " & CONTACT_ID)
                    EXAM_ID = REG_ID
                    EXAM_ID = Processing.GenerateNewExam(TST_ID, EXAM_ID, CONTACT_ID, "N", SESS_PART_ID, Debug)
                    If Debug = "Y" Then mydebuglog.Debug("  .. EXAM_ID : " & EXAM_ID)
                End If
      
                ' Open the exam for the student
                GoTo Process
            End If
            
            ' CASE: survey or recertification exam - only exam known
            If REG_ID = "" And TST_ID <> "" And EXAM_ID = "" Then
                ' Create an instance of an exam for the test - assume a non-participant
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Calling GenerateNewExam: TST_ID: " & TST_ID & " - SESS_PART_ID: " & SESS_PART_ID & " - CONTACT_ID: " & CONTACT_ID)
                EXAM_ID = REG_ID
                EXAM_ID = Processing.GenerateNewExam(TST_ID, EXAM_ID, CONTACT_ID, "N", "", Debug)
                If Debug = "Y" Then mydebuglog.Debug("  .. EXAM_ID : " & EXAM_ID)
      
                ' Open the exam for the student
                GoTo Process
            End If
            
            ' CASE: continuation of survey or exam
            If REG_ID = "" And EXAM_ID <> "" Then
                ' If it does not exist then error out      
                GoTo Process
            End If
            
            ' CASE: Any other permutation is a parameter error
            GoTo AccessError
            
Process:
            ' ================================================
            ' CHECK FOR ALTERNATIVE LANGUAGE EXAMS FOR THE COURSE IF THEY HAVEN'T ALREADY STARTED AN EXAM
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  EXAM_ID / TST_ID / CRSE_ID: " & EXAM_ID & " / " & TST_ID & " / " & CRSE_ID)
            ctr = 0
            If (CRSE_ID <> "" And TST_ID <> "") And EXAM_STATUS_CD = "Pending" And X_SURVEY_FLG <> "Y" And COURSE_TYPE_CD <> "TIPS Participant" Then
                SqlS = "SELECT T.ROW_ID, T.X_LANG_ID, L.NAME, C.NAME, T.NAME " & _
                "FROM siebeldb.dbo.S_CRSE_TST T " & _
                "INNER JOIN siebeldb.dbo.S_CRSE C ON C.ROW_ID=T.CRSE_ID " & _
                "LEFT OUTER JOIN siebeldb.dbo.S_LANG L ON L.LANG_CD=T.X_LANG_ID " & _
                "WHERE T.CRSE_ID='" & CRSE_ID & "' AND T.STATUS_CD='Active' AND " & _
                "T.X_SURVEY_FLG='N' AND T.ROW_ID<>'" & TST_ID & "' AND T.X_LANG_ID IS NOT NULL AND T.X_LANG_ID<>'" & LANG_CD & "'"
                If SKILL_LEVEL_CD <> "" Then SqlS = SqlS & " AND T.SKILL_LEVEL_CD='" & SKILL_LEVEL_CD & "'"
                If CATEGORY <> "" And SKILL_LEVEL_CD = "Trainer" Then SqlS = SqlS & " AND T.X_CATEGORY='" & CATEGORY & "'"
                If JURIS_ID <> "" Then
                    SqlS = SqlS & " AND T.X_JURIS_ID='" & JURIS_ID & "'"
                Else
                    SqlS = SqlS & " AND (T.X_JURIS_ID IS NULL OR T.X_JURIS_ID='')"
                End If
                SqlS = SqlS & " ORDER BY T.X_VERSION DESC"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  ALTERNATIVE LANGUAGE QUERY: " & vbCrLf & "  " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            ctr = ctr + 1
                            ALT_TEST_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            ALT_LANG_CD = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            ALT_LANG = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                            If ctr = 1 Then
                                CRSE_NAME = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                                ALT_EXAMS = "The exam for the course '" & CRSE_NAME & "' is available in other language(s):"
                            End If
                            ALT_TEST = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                            ALT_EXAMS = ALT_EXAMS & "<h1><a href=""#"" onclick=""GetUpdate('" & REG_ID & "','" & EXAM_ID & "','" & SESS_PART_ID & "','" & ALT_TEST_ID & "')"">Click to take the <u>" & ALT_LANG & "</u> exam</a></h1>"
                        Catch ex As Exception
                        End Try
                    End While
                    If ctr > 0 Then
                        ALT_EXAMS = "<table border=""0"" width=""100%"" height=""" & Trim(Str((ctr * 20) + 80)) & """ class=""list1""><tr valign=""middle""><td align=""center"" class=""BigHeader"">" & ALT_EXAMS & "<br/><i>Select "
                        If ctr > 1 Then
                            ALT_EXAMS = ALT_EXAMS & "one of the above to take an alternative language " & TYPE_CD
                        Else
                            ALT_EXAMS = ALT_EXAMS & "the above to take the alternative language " & TYPE_CD
                        End If
                        ALT_EXAMS = ALT_EXAMS & ", or select<br>the option below to take the <u>" & LANG & "</u> " & TYPE_CD & ".</i><br/></td></tr></table>"
                        ALT_ASSESS = "Y"
                        If Debug = "Y" Then
                            mydebuglog.Debug("  .. ALT_ASSESS : " & ALT_ASSESS)
                            mydebuglog.Debug("  .. ALT_EXAMS : " & ALT_EXAMS)
                        End If
                    End If
                End If
                dr.Close()
            End If
            
            ' ================================================
            ' REDIRECT TO EXAM SYSTEM
            If EXAM_ID <> "" Then
                ' CHECK THE TEST 
                ' Get default theme
                SqlS = "SELECT T.X_ENGINE, T.NAME, L.NAME " & _
                "FROM siebeldb.dbo.S_CRSE_TST T " & _
                "LEFT OUTER JOIN siebeldb.dbo.S_LANG L ON L.LANG_CD=T.X_LANG_ID " & _
                "WHERE T.ROW_ID='" & TST_ID & "'"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  CHECK PLAYER: " & vbCrLf & "  " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            ENGINE = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            EXAM = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            LANG = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                        Catch ex As Exception
                        End Try
                    End While
                End If
                dr.Close()
                If Debug = "Y" Then
                    mydebuglog.Debug("  .. ENGINE : " & ENGINE)
                    mydebuglog.Debug("  .. EXAM : " & EXAM)
                    mydebuglog.Debug("  .. LANG : " & LANG)
                End If
                If ENGINE = "" Then ENGINE = "SCORM"
                If ENGINE = "HCIPLAYER" Then HCIPlayer = True
                
                ' -----
                ' Double check player
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  CRSE_ID / REG_ID: " & CRSE_ID & " / " & REG_ID)
                temp = ""
                If CRSE_ID <> "" And REG_ID <> "" And (ENGINE = "SCORM" Or ENGINE = "HCIPLAYER") Then
                    SqlS = "SELECT from_db FROM elearning.dbo.Elearning_Player_Data WHERE reg_id='" & REG_ID & "' AND crse_id='" & CRSE_ID & "'"
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  DOUBLE CHECK PLAYER: " & vbCrLf & "  " & SqlS)
                    cmd.CommandText = SqlS
                    Try
                        temp = CheckDBNull(cmd.ExecuteScalar(), enumObjectType.StrType)
                    Catch ex As Exception
                    End Try
                    If temp <> "" Then
                        If temp = "hciscorm" Then
                            ENGINE = "SCORM"
                            HCIPlayer = False
                        Else
                            ENGINE = "HCIPLAYER"
                            HCIPlayer = True
                        End If
                    Else
                        If temp = "elearning" Then
                            ENGINE = "HCIPLAYER"
                            HCIPlayer = True
                        Else
                            ENGINE = "SCORM"
                            HCIPlayer = False
                        End If
                    End If
                End If
            
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  EXAM_ID / ENGINE: " & EXAM_ID & " / " & ENGINE)
                If EXAM_ID <> "" And (ENGINE = "SCORM" Or ENGINE = "HCIPLAYER") Then
                    temp = ""
                    SqlS = "SELECT from_db FROM elearning.dbo.Elearning_Player_Data WHERE reg_id='" & EXAM_ID & "' AND crse_type='A'"
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  TRIPLE CHECK PLAYER: " & vbCrLf & "  " & SqlS)
                    cmd.CommandText = SqlS
                    Try
                        temp = CheckDBNull(cmd.ExecuteScalar(), enumObjectType.StrType)
                    Catch ex As Exception
                    End Try
                    If temp <> "" Then
                        If temp = "elearning" Then
                            ENGINE = "HCIPLAYER"
                            HCIPlayer = True
                        Else
                            ENGINE = "SCORM"
                            HCIPlayer = False
                        End If
                    Else
                        
                    End If
                End If
                If Debug = "Y" Then
                    mydebuglog.Debug("  .. from_db: " & temp)
                    mydebuglog.Debug("  .. ENGINE: " & ENGINE)
                    mydebuglog.Debug("  .. HCIPlayer: " & Str(HCIPlayer))
                End If

                ' -----
                ' Prepare exam link      
                If Debug = "Y" Then mydebuglog.Debug("  PREPARE LINKS ")
                If LANG_CD = "ESN" Then
                    MSG1 = "<span class=""BigHeader""><font color=""Gray"">Examen de localizaci&oacute;n. Un momento por favor..</font></span>"
                Else
                    MSG1 = "<span class=""BigHeader""><font color=""Gray"">Locating exam.  One moment please..</font></span>"
                End If
                Select Case ENGINE
                    Case "SCORM"
                        If TEST_FLG = "Y" Then
                            ExamLink = LaunchProtocol & "//hciscorm.certegrity.com/scorm/defaultui/launch.aspx?registration=CrseId|" & TST_ID & "!CrseType|A!RegId|" & EXAM_ID & "!UserId|" & CONTACT_ID & "!InstanceId|0&configuration=Popup|false!DiagnosticsLog|true!DiagnosticsDetailedLog|true&forceFrameset=true&player=modern"
                        Else
                            ExamLink = LaunchProtocol & "//hciscorm.certegrity.com/scorm/defaultui/launch.aspx?registration=CrseId|" & TST_ID & "!CrseType|A!RegId|" & EXAM_ID & "!UserId|" & CONTACT_ID & "!InstanceId|0&configuration=Popup|false!DiagnosticsLog|false!DiagnosticsDetailedLog|false&forceFrameset=true&player=modern"
                        End If
                    Case "HCIPLAYER"
                        ExamLink = LaunchProtocol & "//hciscorm.certegrity.com/HCIPlayer/HCIlaunch.aspx?"
                        ExamLink = ExamLink & "CrseId=" & TST_ID & "&CrseType=A&VersionId=0&RegId=" & EXAM_ID & "&UserId=" & CONTACT_ID & "&InstanceId=0&Debug=Y"
                    Case Else
                        ExamLink = "http://elearning.certegrity.com/test-render/examClient.x?testRunId=" & EXAM_ID
                End Select
                UnLoadLink = "AcceptExam('" & EXAM_ID & "','" & REG_ID & "');"
                
                ' -----
                ' Compute resolution
                If RESOLUTION = "" Then
                    Select Case ENGINE
                        Case "SCORM"
                            RESOLUTION = "1047x791"
                        Case "HCIPLAYER"
                            RESOLUTION = "1047x791"
                        Case Else
                            RESOLUTION = "825x625"
                    End Select
                End If
                RES_X = Left(RESOLUTION, InStr(1, UCase(RESOLUTION), "X") - 1)
                RES_X = Trim(Str(Val(RES_X) + 10))
                RES_Y = Right(RESOLUTION, Len(RESOLUTION) - InStr(1, UCase(RESOLUTION), "X"))
                RES_Y = Trim(Str(Val(RES_Y) + 10))
                
                ' ================================================
                ' PREPARE CONTINUABLE WARNING IF APPLICATION
                If CONTINUABLE_FLG = "N" Then
                    AutoStart = False
                    MSG1 = "<table cellpadding=""1"" border=""0""  bgcolor=""#FFFFFF"" cellspacing=""0"" width=""100%""> " & EOL & _
                    "<tr><td valign=""middle"" height=""40%"" class=""BigHeader"" align=""center""> " & EOL & _
                    "<span class=""BigHeader""><font color=""Red"">" & EOL
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = MSG1 & "El examen <i>" & EXAM & "</i> est&aacute; en <i>Espa&ntilde;ol</i> y debe completarse en una sesi&oacute;n.<br>Si sale del examen por cualquier motivo, se anotar&aacute; en ese momento." & EOL
                            If ALT_EXAMS <> "" Then MSG1 = MSG1 & "<br><br>" & ALT_EXAMS & "</td></tr>" & EOL
                            MSG1 = MSG1 & "<tr style=""margin-top: 50px""><td align=""Center""> " & EOL & _
                            "<a href=""JavaScript:SetMsg()"" data-role=""button"" id=""AssessmentBtn"" data-theme=""b"" style=""width: 50%;""><u>Haga clic aqu&iacute; para tomar su examen de Espa&ntilde;ol ahora</u></a></font><br> " & EOL
                            MSG1 = MSG1 & "<a href=""#"" onclick=""ReturnSource()"" id=""ProceedBtn"" data-role=""button"" rel=""external"" data-theme=""c"" style=""width: 50%;"">Haga clic para tomar el examen m&aacute;s tarde</a>';"
                        Case Else
                            MSG1 = MSG1 & "The exam <i>" & EXAM & "</i> is in <i>" & LANG & "</i> and must be completed in one sitting.<br>If you exit the exam for any reason, it will be scored at that point." & EOL
                            If ALT_EXAMS <> "" Then MSG1 = MSG1 & "<br><br>" & ALT_EXAMS & "</td></tr>" & EOL
                            MSG1 = MSG1 & "<tr style=""margin-top: 50px""><td align=""Center""> " & EOL & _
                            "<a href=""JavaScript:SetMsg()"" data-role=""button"" id=""AssessmentBtn"" data-theme=""b"" style=""width: 50%;""><u>Click here to take your " & LANG & " exam now</u></a></font><br> " & EOL
                            MSG1 = MSG1 & "<a href=""#"" onclick=""ReturnSource()"" id=""ProceedBtn"" data-role=""button"" rel=""external"" data-theme=""c"" style=""width: 50%;"">Click to Take Exam Later</a>';"
                    End Select
                    MSG1 = MSG1 & "</td></tr> " & EOL & _
                    "</table>" & EOL
                Else
                    If ALT_ASSESS = "Y" And ALT_EXAMS <> "" Then
                        MSG1 = "<table cellpadding=""1"" border=""0""  bgcolor=""#FFFFFF"" cellspacing=""0"" width=""100%""> " & EOL & _
                        "<tr><td valign=""middle"" height=""40%"" class=""BigHeader"" align=""center""> " & EOL
                        MSG1 = MSG1 & "<br><br>" & ALT_EXAMS & "" & EOL
                        MSG1 = MSG1 & "</td></tr> " & EOL
                        Select Case LANG_CD
                            Case "ESN"
                                MSG1 = MSG1 & "<tr style=""margin-top: 50px""><td align=""Center""> " & EOL & _
                                "<a href=""JavaScript:SetMsg()"" data-role=""button"" id=""AssessmentBtn"" data-theme=""b"" style=""width: 50%;""><u>Haga clic aqu&iacute; para tomar su examen de Espa&ntilde;ol ahora</u></a></font><br> " & EOL
                                MSG1 = MSG1 & "<a href=""#"" onclick=""ReturnSource()"" id=""ProceedBtn"" data-role=""button"" rel=""external"" data-theme=""c"" style=""width: 50%;"">Haga clic para tomar el examen m&aacute;s tarde</a>';"
                            Case Else
                                MSG1 = MSG1 & "<tr style=""margin-top: 50px""><td align=""Center""> " & EOL & _
                                "<a href=""JavaScript:SetMsg()"" data-role=""button"" id=""AssessmentBtn"" data-theme=""b"" style=""width: 50%;""><u>Click here to take your " & LANG & " exam now</u></a></font><br> " & EOL
                                MSG1 = MSG1 & "<a href=""#"" onclick=""ReturnSource()"" id=""ProceedBtn"" data-role=""button"" rel=""external"" data-theme=""c"" style=""width: 50%;"">Click to Take Exam Later</a>';"
                        End Select
                        MSG1 = MSG1 & "</td></tr> " & EOL & _
                        "</table>" & EOL
                    Else
                        AutoStart = True
                        MSG1 = "<span class=""BigHeader"">Locating your exam.  One moment please...</span>"
                    End If
                End If

                ' ================================================
                ' LOG THE ENTRANCE TO THE EXAM
                SqlS = "IF (SELECT TOP 1 ENTER_FLG FROM siebeldb.dbo.S_CRSE_TSTRUN_ACCESS WHERE CRSE_TSTRUN_ID = '" & EXAM_ID & "' ORDER BY CREATED DESC)='N' OR " & _
                "(SELECT TOP 1 ENTER_FLG FROM siebeldb.dbo.S_CRSE_TSTRUN_ACCESS WHERE CRSE_TSTRUN_ID = '" & EXAM_ID & "' ORDER BY CREATED DESC) IS NULL BEGIN;  " & _
                "INSERT INTO siebeldb.dbo.S_CRSE_TSTRUN_ACCESS(ROW_ID, CREATED, CREATED_BY, LAST_UPD, LAST_UPD_BY, " & _
                "MODIFICATION_NUM, CONFLICT_ID, CRSE_TSTRUN_ID, ENTER_FLG, EXIT_FLG, CALL_ID, CALL_SCREEN) " & _
                "SELECT '" & EXAM_ID & "-'+LTRIM(CAST(COUNT(*)+1 AS VARCHAR)), GETDATE(), '0-1', GETDATE(), '0-1', " & _
                "0, 0, '" & EXAM_ID & "','Y' ,'N' ,'" & CALL_ID & "', '" & sCALL_SCREEN & "' " & _
                "FROM siebeldb.dbo.S_CRSE_TSTRUN_ACCESS " & _
                "WHERE CRSE_TSTRUN_ID = '" & EXAM_ID & "'; END;"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  LOG ENTER QUERY: " & vbCrLf & "  " & SqlS)
                temp = ExecQuery("Insert", "S_CRSE_TSTRUN_ACCESS", cmd, SqlS, mydebuglog, "N")
            
                ' Log to CM activity log
                SqlS = "INSERT INTO reports.dbo.CM_LOG(REG_ID, SESSION_ID, RECORD_ID, TRANSACTION_ID, REMOTE_ADDR, ACTION) " & _
                "VALUES('" & UID & "','" & SessID & "','" & EXAM_ID & "','" & CALL_ID & "','" & REMOTE_ADDR & "','ENTERED EXAM')"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  LOG TO CM LOG: " & vbCrLf & "  " & SqlS)
                temp = ExecQuery("Insert", "CM_LOG", cmd, SqlS, mydebuglog, "N")
            
                ' ================================================
                ' UPDATE TEST WITH RETURN URL
                LAST_INST = LaunchProtocol & "//hciscorm.certegrity.com/ls/FinishAssessment.html?ID=" & EXAM_ID & "&UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "&HP=" & HOME_PAGE & "&CUR=" & CURRENT_PAGE
                SqlS = "UPDATE siebeldb.dbo.S_CRSE_TSTRUN " & _
                "SET X_REDIRECT_URL='" & LAST_INST & "'"
                If SESS_PART_ID <> "" Then SqlS = SqlS & ", X_PART_ID='" & SESS_PART_ID & "'"
                SqlS = SqlS & " WHERE ROW_ID='" & EXAM_ID & "'"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  CREATE REDIRECTION URL: " & vbCrLf & "  " & SqlS)
                temp = ExecQuery("Update", "S_CRSE_TSTRUN", cmd, SqlS, mydebuglog, "N")
            Else
                GoTo AccessError
            End If
        Else
            GoTo DBError
        End If
        
        ' ================================================
        ' RETURN TO USER
        ' This creates a frame which loads the content.  If the content is an exam, it also loads a Javascript library
        ' that handles the unload event
ReturnControl:
        If Debug = "Y" Then
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & ">>Final Values")
            mydebuglog.Debug("  .. ClassLink: " & ClassLink)
            mydebuglog.Debug("  .. delay: " & delay)
            mydebuglog.Debug("  .. MSG1: " & MSG1)
            mydebuglog.Debug("  .. RefreshLink: " & RefreshLink)
            mydebuglog.Debug("  .. UnLoadLink: " & UnLoadLink)
            mydebuglog.Debug("  .. ExamLink: " & ExamLink)
            mydebuglog.Debug("  .. REG_STATUS_CD: " & REG_STATUS_CD)
            mydebuglog.Debug("  .. ALT_EXAMS: " & ALT_EXAMS)
            mydebuglog.Debug("  .. ALT_ASSESS: " & ALT_ASSESS)
            mydebuglog.Debug("  .. RESOLUTION: " & RES_X & " x " & RES_Y)
        End If
        GoTo PrepareExit
        
OnHoldError:
        If Debug = "Y" Then mydebuglog.Debug(">>OnHoldError")
        Select Case LANG_CD
            Case "ESN"
                errmsg = "Su examen est&aacute; en espera y no puede realizarlo en espera de una revisi&oacute;n."
            Case Else
                errmsg = "Your exam is On Hold and you cannot take it pending a review."
        End Select
        GoTo PrepareExit        

AccessCountError:
        If Debug = "Y" Then mydebuglog.Debug(">>AccessCountError")
        Select Case LANG_CD
            Case "ESN"
                errmsg = "Ha superado el acceso normal para su evaluaci&oacute;n y se ha puesto en espera en espera de una revisi&oacute;n."
            Case Else
                errmsg = "You exceeded normal access for your assessment and it has been placed On Hold pending a review."
        End Select
        GoTo PrepareExit

DBError:
        If Debug = "Y" Then mydebuglog.Debug(">>DBError: " & MSG1)
        MSG1 = ""
        Select Case LANG_CD
            Case "ESN"
                errmsg = "Se ha producido un error al acceder a los datos de evaluaci&oacute;n."
            Case Else
                errmsg = "There was an error accessing " & TYPE_CD & " data."
        End Select
        GoTo PrepareExit
        
ExamError:
        ErrLvl = "Warning"
        If Debug = "Y" Then mydebuglog.Debug(">>ExamError")
        Select Case LANG_CD
            Case "ESN"
                errmsg = "No podemos encontrar una evaluaci&oacute;n que coincida con su curso y jurisdicci&oacute;n. <br> P&oacute;ngase en contacto con nosotros para obtener ayuda."
            Case Else
                errmsg = "We are unable to find a(n) " & TYPE_CD & " that matches your course and jurisdiction.<br>Please contact us for assistance."
        End Select
        GoTo PrepareExit
        
DataError:
        ErrLvl = "Warning"
        If Debug = "Y" Then mydebuglog.Debug(">>DataError")
        Select Case LANG_CD
            Case "ESN"
                MSG1 = "Se ha producido un error al acceder a los datos de evaluaci&oacute;n."
            Case Else
                MSG1 = "There was an error accessing " & TYPE_CD & " data."
        End Select
        GoTo PrepareExit
        
AccessError:
        ErrLvl = "Warning"
        If Debug = "Y" Then mydebuglog.Debug(">>AccessError")
        Select Case LANG_CD
            Case "ESN"
                errmsg = "Se ha producido un error al intentar abrir esta evaluaci&oacute;n."
            Case Else
                errmsg = "There was an error attempting to open this " & TYPE_CD & "."
        End Select
        ErrLvl = "Warning"
        GoTo PrepareExit

AutoLoggedOut:
        ErrLvl = "Warning"
        If Debug = "Y" Then mydebuglog.Debug(">>AutoLoggedOut")
        Select Case LANG_CD
            Case "ESN"
                errmsg = "Nuestro sistema cerr&oacute; su sesi&oacute;n; lo m&aacute;s probable es que haya abierto nuestro portal en otra ventana del navegador. <br> Regrese al portal y vuelva a iniciar sesi&oacute;n para realizar su evaluaci&oacute;n."
            Case Else
                errmsg = "You were logged out by our system - most likely you had our portal open in another browser window.<br>Please return to the portal and log back in again to take your " & TYPE_CD & "."
        End Select
        GoTo PrepareExit
        
PrepareExit:
        If NextLink = "" Then
            If CURRENT_PAGE <> "" Then
                NextLink = CURRENT_PAGE
                If InStr(NextLink, "UID=") = 0 Then NextLink = NextLink & "?UID=" & UID
                If InStr(NextLink, "SES=") = 0 Then NextLink = NextLink & "&SES=" & SessID
                If InStr(NextLink, "PP=") = 0 Then NextLink = NextLink & "&PP=" & DOMAIN
                If TYPE_CD = "survey" Then
                    If InStr(NextLink, "#cert") = 0 Then NextLink = NextLink & "#cert"
                Else
                    If InStr(NextLink, "#reg") = 0 Then NextLink = NextLink & "#reg"
                End If
            Else
                NextLink = "https://" & HOME_PAGE & "/mobile/index.html" & "?UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN
                If TYPE_CD = "Survey" Then NextLink = NextLink & "#cert" Else NextLink = NextLink & "#reg"
            End If
            If Debug = "Y" Then mydebuglog.Debug("NextLink: " & NextLink)
            MSG1 = MSG1 & "<br><br><a href=" & NextLink & " data-role=""button"" rel=""external"" data-theme=""b"">Click to return to the portal</a>"
        End If
        
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
        outdata = outdata & """TST_ID"":""" & TST_ID & ""","
        outdata = outdata & """REG_ID"":""" & REG_ID & ""","
        outdata = outdata & """UserKey"":""" & UserKey & ""","
        outdata = outdata & """TYPE_CD"":""" & TYPE_CD & ""","
        outdata = outdata & """NextLink"":""" & EscapeJSON(NextLink) & ""","
        outdata = outdata & """RefreshLink"":""" & EscapeJSON(RefreshLink) & ""","
        outdata = outdata & """ExamLink"":""" & EscapeJSON(ExamLink) & ""","
        outdata = outdata & """UnLoadLink"":""" & EscapeJSON(UnLoadLink) & ""","
        outdata = outdata & """LAST_INST"":""" & EscapeJSON(LAST_INST) & ""","
        outdata = outdata & """AutoStart"":""" & AutoStart & ""","
        outdata = outdata & """delay"":""" & delay & ""","
        outdata = outdata & """MSG1"":""" & EscapeJSON(MSG1) & ""","
        outdata = outdata & """ALT_ASSESS"":""" & ALT_ASSESS & ""","
        outdata = outdata & """EXAM_STATUS_CD"":""" & EXAM_STATUS_CD & ""","
        outdata = outdata & """X_FORMAT"":""" & X_FORMAT & ""","
        outdata = outdata & """CONTINUABLE_FLG"":""" & CONTINUABLE_FLG & ""","
        outdata = outdata & """KBA_QUES_NUM"":""" & KBA_QUES_NUM & ""","
        outdata = outdata & """ENGINE"":""" & ENGINE & ""","
        outdata = outdata & """EXAM"":""" & EXAM & ""","
        outdata = outdata & """DOMAIN"":""" & DOMAIN & ""","
        outdata = outdata & """RESOLUTION"":""" & RESOLUTION & ""","
        outdata = outdata & """RES_X"":""" & RES_X & ""","
        outdata = outdata & """RES_Y"":""" & RES_Y & ""","
        outdata = outdata & """HCIPlayer"":""" & HCIPlayer & ""","
        outdata = outdata & """USER_NAME"":""" & USER_NAME & ""","
        outdata = outdata & """EMAIL_ADDR"":""" & EMAIL_ADDR & ""","
        outdata = outdata & """CALL_SCREEN"":""" & CALL_SCREEN & ""","
        outdata = outdata & """CALL_ID"":""" & CALL_ID & ""","
        outdata = outdata & """ErrMsg"":""" & errmsg & """ "
        outdata = callback & "({""ResultSet"": {" & outdata & "} })"
        
        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("WsGetAssessment.ashx : " & ErrLvl & ": " & Trim(errmsg))
        myeventlog.Info("WsGetAssessment.ashx : Contact Id: " & CONTACT_ID & ", Reg Id: " & REG_ID & ", Sess Id: " & SessID & ", Exam Id: " & EXAM_ID & ", Tst Id: " & TST_ID & " - NextLink: " & NextLink)
        If Debug = "Y" Or (Logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("  outdata: " & outdata & vbCrLf)
                mydebuglog.Debug("Results:  Contact Id: " & CONTACT_ID & ", Reg Id: " & REG_ID & ", Sess Id: " & SessID & ", Exam Id: " & EXAM_ID & ", Tst Id: " & TST_ID & " - NextLink: " & NextLink)
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
                LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, "WSGETASSESSMENT", LogStartTime, VersionNum, Debug)
            Catch ex As Exception
            End Try
        End If
        
        ' Send results        
        If Debug = "T" Then
            context.Response.ContentType = "text/html"
            If outdata <> "" Then
                context.Response.Write("Success")
            Else
                context.Response.Write("Failure")
            End If
        Else
            If outdata = "" Then outdata = errmsg
            context.Response.ContentType = "application/json"
            context.Response.Write(outdata)
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