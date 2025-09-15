<%@ WebHandler Language="VB" Class="WsLeaveAssessment" %>

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

Public Class WsLeaveAssessment : Implements IHttpHandler
    
    Enum enumObjectType
        StrType = 0
        IntType = 1
        DblType = 2
        DteType = 3
    End Enum
    
    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        
        '   Name:      WsLeaveAssessment service
        '   
        '   Goal:      To process the exit from an exam or survey
        '   
        '   Preconditions:   
        '      1.   Student records exist (S_CONTACT)
        '      2.   Test record exists (S_CRSE_TSTRUN)
        '      3.    Exam answer records may exist (elearning.dbo.ELN_TEST_ANSWER)
        '      4. Trainer Certification (S_CURRCLM_PER) or Student Participation (CX_SESS_PART_X) may exist
        '   
        '   Postconditions:
        '      IF AN EXAM RECORD..
        '      1.    If the exam is not finished, this agent simply notes that fact to the user, and allows the user to close the window
        '      2.    If the exam is finished, it should notify the user appropriately, and do the following to objects in the database:
        '         .. Score exam
        '         .. Update S_CRSE_TSTRUN record
        '         .. Create S_CRSE_TSTRUN_Q/A records
        '         .. Evaluate S_CRSE_TST.SKILL_LEVEL_CD:
        '            .. If "Trainer", and S_CURRCLM_PER record exists then update it
        '            .. If "Trainer", and S_CURRCLM_PER record does not exist then create it
        '            .. If "Participant", and X_PART_ID is empty, then create CX_SESS_PART_X record
        '            .. If "Participant", and X_PART_ID is not empty, update CX_SESS_PART_X record
        '      3.    If the student passes the exam..
        '            .. take the student to success page, email success message
        '      4.   If the student failed the exam email results and..
        '            .. if retakes allowed:
        '               a. create a new exam
        '               b. email link to new exam to student
        '               c. take the student to the new start page
        '            .. if retakes not allowed:
        '               a. take the student to a results page
        '      IF A SURVEY RECORD:
        '      1. If the survey was not finished, this agent notes that fact to the user and allows him/her to close the window
        '      2. If the survey was finished, it should create a "Thank you" message for the user, and create objects in the database
        '
        '   Actors:    Student
        '   
        '   Trigger:   Agent run from the course in the same window by completion of exam
        '   
        '   Flow of Events: See code below
        '   
        '   Extensions:    None
        '   
        '   Special Requirements:
        '      S_CRSE_TSTRUN.STATUS_CD field values:
        '          Pending    (ready to be taken)
        '          Unsubmitted    (exam or survey in progress)
        '          Submitted (exam taker completed exam.. sent to scoring agent)
        '          Ungraded (scoring system started processing the exam/survey)
        '          Graded (scoring system completes processing of exam/survey)
        '          Cancelled (cancelled/aborted by the user)
        '   
        '   Outstanding Issues:  none
        
        ' Parameter Declarations
        Dim Debug As String
        Dim REG_ID, UID, SessID, CURRENT_PAGE, HOME_PAGE, LANG_CD, callback, myprotocol, RefreshID As String
        
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
        mydebuglog = log4net.LogManager.GetLogger("LeaveAssessmentDebugLog")
        Dim logfile, tempdebug As String
        Dim Logging As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "100"
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
        Dim Cmp As New com.certegrity.hciscormsvc.cmp.CMProfiles
        
        ' Variable declarations
        Dim errmsg, Results As String                    ' Error message (if any)
        Dim DupeCall As Boolean
        Dim LOGGED_IN, CONTACT_ID, SUB_ID, CONTACT_OU_ID, DOMAIN As String
        Dim CRSE_CONTENT_URL As String
        Dim EMAIL_ADDR, EXM_PART_ID, CRSE_NAME, CRSE_TSTRUN_ID As String
        Dim MergeData, SendTo, SendFrom, NextLink, pTYPE_CD, TYPE_CD, TST_SKILL_LEVEL_CD, TST_CRSE_NAME As String
        Dim temp, ErrLvl, BasePage, testingURL As String
        Dim EXM_STATUS_CD, ClassLink, MSG1, TST_NAME, PAUSE, TST_ID, LaunchProtocol As String
        Dim EXM_CRSE_TST_ID, EXM_PASSED_FLG, EXM_TEST_DT, EXM_CRSE_OFFR_ID, EXM_FINAL_SCORE, EXM_GRADED_DT, EXM_CON_ID As String
        Dim EXM_ORDER_ITEM_ID, EXM_PAR_TSTRUN_ID, EXM_CURRCLM_PER_ID, TST_CRSE_ID, TST_STATUS_CD, TST_VERSION, TST_SURVEY_FLG, TST_MAX_POINTS As String
        Dim TST_PASSING_SCORE, TST_CONTINUABLE_FLG, TST_CATEGORY, TST_RETAKE_FLG, TST_REDIRECT_URL, TST_ORDER_ID, CURRCLM_ID, CRSE_TYPE_CD As String
        Dim TRNR_REG_ID, SESS_PART_ID, RETAKE_AUTH, RESOLUTION, VAL_UID, SCORM_FLG, POPUP_FLG, RESTART_RESET, TEST_FORMAT, X_ALT_RETAKE_FLG, SR_RETAKE_FLG As String
        Dim sleepcount, NumIts As Integer
        Dim EncodedUID As String

        ' ============================================
        ' Variable setup
        HOME_PAGE = ""
        myprotocol = ""
        LANG_CD = ""
        CURRENT_PAGE = ""
        RefreshID = ""
        callback = ""
        Debug = "N"
        Logging = "Y"
        errmsg = ""
        DupeCall = False
        ClassLink = ""
        MSG1 = ""
        TST_NAME = ""
        PAUSE = ""
        CRSE_TSTRUN_ID = ""
        TST_ID = ""
        UID = ""
        SessID = ""
        NextLink = ""
        DOMAIN = "TIPS"
        TST_ID = ""
        REG_ID = ""
        outdata = ""
        LOGGED_IN = "N"
        CONTACT_ID = ""
        SUB_ID = ""
        CONTACT_OU_ID = ""
        DOMAIN = "TIPS"
        EXM_CRSE_TST_ID = ""
        EXM_PASSED_FLG = ""
        EXM_TEST_DT = ""
        EXM_CRSE_OFFR_ID = ""
        EXM_FINAL_SCORE = ""
        EXM_GRADED_DT = ""
        EXM_CON_ID = ""
        EXM_STATUS_CD = ""
        EXM_PART_ID = ""
        EXM_ORDER_ITEM_ID = ""
        EXM_PAR_TSTRUN_ID = ""
        EXM_CURRCLM_PER_ID = ""
        TST_ID = ""
        TST_NAME = ""
        TST_CRSE_ID = ""
        TST_SKILL_LEVEL_CD = ""
        TST_STATUS_CD = ""
        TST_VERSION = ""
        TST_SURVEY_FLG = ""
        TST_MAX_POINTS = ""
        TST_PASSING_SCORE = ""
        TST_CONTINUABLE_FLG = ""
        TST_CATEGORY = ""
        TST_RETAKE_FLG = ""
        TST_REDIRECT_URL = ""
        TST_ORDER_ID = ""
        CURRCLM_ID = ""
        CRSE_TYPE_CD = ""
        CRSE_NAME = ""
        TRNR_REG_ID = ""
        REG_ID = ""
        SESS_PART_ID = ""
        RETAKE_AUTH = ""
        RESOLUTION = ""
        VAL_UID = ""
        SCORM_FLG = ""
        POPUP_FLG = ""
        RESTART_RESET = ""
        TEST_FORMAT = ""
        EMAIL_ADDR = ""
        temp = ""
        CRSE_CONTENT_URL = ""
        X_ALT_RETAKE_FLG = ""
        SR_RETAKE_FLG = ""
        TST_CRSE_NAME = ""
        LaunchProtocol = "http:"
        BasePage = ""
        testingURL = ""
        NumIts = 0
        sleepcount = 2
        EncodedUID = ""
        TYPE_CD = ""
        ErrLvl = "Error"
        
        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
            tempdebug = System.Configuration.ConfigurationManager.AppSettings.Get("LeaveAssessment_debug")
            If tempdebug = "Y" And Debug <> "T" Then Debug = "Y"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            GoTo CloseOut2
        End Try
        
        ' ============================================
        ' Open log file if applicable
        If Logging = "Y" Then
            logfile = "C:\Logs\WsLeaveAssessment.log"
            Try
                log4net.GlobalContext.Properties("LeaveAssessmentLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                GoTo CloseOut2
            End Try
        End If
        
        ' ============================================
        ' Get parameters    
        If Not context.Request.QueryString("ID") Is Nothing Then
            CRSE_TSTRUN_ID = context.Request.QueryString("ID")
        End If
        If CRSE_TSTRUN_ID = "" Then
            If Not context.Request.QueryString("EID") Is Nothing Then
                CRSE_TSTRUN_ID = context.Request.QueryString("EID")
            End If
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

        If Not context.Request.QueryString("RFR") Is Nothing Then
            RefreshID = context.Request.QueryString("RFR")
        End If
 
        If Not context.Request.QueryString("CUR") Is Nothing Then
            CURRENT_PAGE = context.Request.QueryString("CUR")
        End If
        
        If Not context.Request.QueryString("LANG") Is Nothing Then
            LANG_CD = UCase(context.Request.QueryString("LANG"))
        End If
 
        If Not context.Request.QueryString("callback") Is Nothing Then
            callback = context.Request.QueryString("callback")
        End If

        ' Validate parameters
        If HOME_PAGE = "web.certegrity.com" Then HOME_PAGE = "web.gettips.com"
        If myprotocol = "" Then myprotocol = "http:"
        CURRENT_PAGE = Replace(CURRENT_PAGE, "/ValFinishExam?OpenAgent", "")
        If LANG_CD <> "ESN" And LANG_CD <> "ENU" Then LANG_CD = "ENU"
        If callback = "" Then callback = "?"
        
        If Debug = "Y" Then
            mydebuglog.Debug("----------------------------------")
            mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
            mydebuglog.Debug("Parameters-")
            mydebuglog.Debug("  Debug: " & Debug)
            mydebuglog.Debug("  UID: " & UID)
            mydebuglog.Debug("  cookieid: " & cookieid)
            mydebuglog.Debug("  SessID: " & SessID)
            mydebuglog.Debug("  CRSE_TSTRUN_ID : " & CRSE_TSTRUN_ID)
            mydebuglog.Debug("  HOME_PAGE: " & HOME_PAGE)
            mydebuglog.Debug("  myprotocol: " & myprotocol)
            mydebuglog.Debug("  CURRENT_PAGE : " & CURRENT_PAGE)
            mydebuglog.Debug("  RefreshID : " & RefreshID)
            mydebuglog.Debug("  LANG_CD: " & LANG_CD)
            mydebuglog.Debug("  PrevLink: " & PrevLink)
            mydebuglog.Debug("  BROWSER: " & BROWSER)
            mydebuglog.Debug("  callback: " & callback)
        End If
        If CRSE_TSTRUN_ID = "" Then GoTo AccessError
        If cookieid <> UID Then GoTo AccessError
        If InStr(1, PrevLink, "WsLeaveAssessment") > 0 Then
            If Debug = "Y" Then mydebuglog.Debug(">> Duplicate call to this agent")
            DupeCall = True
        End If
        
        ' ============================================
        ' Open database connection 
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            errmsg = errmsg & "Unable to open the database connection. " & vbCrLf
            GoTo CloseOut
        End If

        ' ============================================
        '  Process Exit
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
                    mydebuglog.Debug("  .. DOMAIN: " & DOMAIN & vbCrLf)
                End If

                If CONTACT_ID = "" Then
                    Dim pdoc As XmlDocument
                    pdoc = Cmp.GetUserProfile(UID, SessID, Debug)
                    If pdoc Is Nothing Then
                    Else
                        Dim oNodeList As XmlNodeList = pdoc.SelectNodes("//profile")
                        For i = 0 To oNodeList.Count - 1
                            CONTACT_ID = GetNodeValue("CONTACT_ID ", oNodeList.Item(i))
                            CONTACT_OU_ID = GetNodeValue("CONTACT_OU_ID", oNodeList.Item(i))
                            SUB_ID = GetNodeValue("SUB_ID", oNodeList.Item(i))
                            DOMAIN = GetNodeValue("DOMAIN", oNodeList.Item(i))
                            LOGGED_IN = GetNodeValue("LOGGED_IN", oNodeList.Item(i))
                        Next
                        If Debug = "Y" Then
                            mydebuglog.Debug("  GetUserProfile-")
                            mydebuglog.Debug("  .. LOGGED_IN: " & LOGGED_IN)
                            mydebuglog.Debug("  .. CONTACT_ID: " & CONTACT_ID)
                            mydebuglog.Debug("  .. CONTACT_OU_ID: " & CONTACT_OU_ID)
                            mydebuglog.Debug("  .. DOMAIN: " & DOMAIN & vbCrLf)
                        End If
                    End If
                End If
                If LOGGED_IN <> "Y" Then GoTo AccessError
            End If
            
            ' Fix HOME_PAGE to remove domain if applicable
            If InStr(HOME_PAGE, LCase(DOMAIN)) > 0 Then
                HOME_PAGE = Replace(HOME_PAGE, "/" & LCase(DOMAIN), "")
                If Debug = "Y" Then mydebuglog.Debug("  .. HOME_PAGE: " & HOME_PAGE & vbCrLf)
            End If

ExecQuery:            
            ' ================================================
            ' RETRIEVE DATA FROM THE S_CRSE_TSTRUN RECORD 
            SqlS = "SELECT TOP 1 TR.CRSE_TST_ID, TR.PASSED_FLG, TR.TEST_DT, TR.CRSE_OFFR_ID, TR.FINAL_SCORE, " & _
              "TR.GRADED_DT, TR.PERSON_ID, TR.STATUS_CD, TR.X_PART_ID, TR.X_ORDER_ITEM_ID, T.ROW_ID, T.NAME, T.CRSE_ID, T.CURRCLM_ID, " & _
              "T.SKILL_LEVEL_CD, T.STATUS_CD, T.X_VERSION, T.X_SURVEY_FLG, T.MAX_POINTS, T.PASSING_SCORE, " & _
              "TR.X_PAR_TSTRUN_ID, T.X_CONTINUABLE_FLG, P.ROW_ID, T.X_CATEGORY, C.TYPE_CD, T.X_RETAKE_FLG, TR.X_REDIRECT_URL , " & _
              "OI.ROW_ID, C.NAME, R.ROW_ID, SP.ROW_ID, SR.ROW_ID, SR.RETAKE_AUTH, C.X_RESOLUTION, CN.X_REGISTRATION_NUM, " & _
              "C.X_SCORM_FLG, TRX.POPUP_FLG, T.X_RESTART_RESET, T.X_FORMAT, CN.EMAIL_ADDR, CN.X_PR_LANG_CD, C.X_CRSE_CONTENT_URL, " & _
              "C.X_ALT_RETAKE_FLG, SR.RETAKE_FLG, C.NAME " & _
              "FROM siebeldb.dbo.S_CRSE_TSTRUN TR " & _
              "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TSTRUN_X TRX ON TRX.PAR_ROW_ID=TR.ROW_ID " & _
              "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TST T ON T.ROW_ID=TR.CRSE_TST_ID " & _
              "LEFT OUTER JOIN siebeldb.dbo.S_CURRCLM_PER P ON P.X_CRSE_TSTRUN_ID=TR.ROW_ID " & _
              "LEFT OUTER JOIN siebeldb.dbo.S_CRSE C ON C.ROW_ID=T.CRSE_ID " & _
              "LEFT OUTER JOIN siebeldb.dbo.S_ORDER_ITEM OI ON OI.ROW_ID=TR.X_ORDER_ITEM_ID " & _
              "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_REG R ON R.CRSE_OFFR_ID=TR.CRSE_OFFR_ID AND R.PERSON_ID=TR.PERSON_ID " & _
              "LEFT OUTER JOIN siebeldb.dbo.CX_SESS_PART_X SP ON SP.CRSE_TSTRUN_ID=TR.ROW_ID " & _
              "LEFT OUTER JOIN siebeldb.dbo.CX_SESS_REG SR ON SR.SESS_PART_ID=SP.ROW_ID " & _
              "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT CN ON CN.ROW_ID=TR.PERSON_ID " & _
              "WHERE TR.ROW_ID='" & CRSE_TSTRUN_ID & "'"
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  RETRIEVE TEST INFORMATION: " & vbCrLf & "  " & SqlS)
            Try
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        EXM_CRSE_TST_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                        EXM_PASSED_FLG = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                        EXM_TEST_DT = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                        EXM_CRSE_OFFR_ID = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                        EXM_FINAL_SCORE = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                        EXM_GRADED_DT = Trim(CheckDBNull(dr(5), enumObjectType.StrType))
                        EXM_CON_ID = Trim(CheckDBNull(dr(6), enumObjectType.StrType))
                        EXM_STATUS_CD = Trim(CheckDBNull(dr(7), enumObjectType.StrType))
                        EXM_PART_ID = Trim(CheckDBNull(dr(8), enumObjectType.StrType))
                        EXM_ORDER_ITEM_ID = Trim(CheckDBNull(dr(9), enumObjectType.StrType))
                        EXM_PAR_TSTRUN_ID = Trim(CheckDBNull(dr(20), enumObjectType.StrType))
                        EXM_CURRCLM_PER_ID = Trim(CheckDBNull(dr(22), enumObjectType.StrType))
                        TST_ID = Trim(CheckDBNull(dr(10), enumObjectType.StrType))
                        TST_NAME = Trim(CheckDBNull(dr(11), enumObjectType.StrType))
                        TST_CRSE_ID = Trim(CheckDBNull(dr(12), enumObjectType.StrType))
                        TST_SKILL_LEVEL_CD = Trim(CheckDBNull(dr(14), enumObjectType.StrType))
                        TST_STATUS_CD = Trim(CheckDBNull(dr(15), enumObjectType.StrType))
                        TST_VERSION = Trim(CheckDBNull(dr(16), enumObjectType.StrType))
                        TST_SURVEY_FLG = Trim(CheckDBNull(dr(17), enumObjectType.StrType))
                        TST_MAX_POINTS = Trim(CheckDBNull(dr(18), enumObjectType.StrType))
                        TST_PASSING_SCORE = Trim(CheckDBNull(dr(19), enumObjectType.StrType))
                        TST_CONTINUABLE_FLG = Trim(CheckDBNull(dr(21), enumObjectType.StrType))
                        TST_CATEGORY = Trim(CheckDBNull(dr(23), enumObjectType.StrType))
                        TST_RETAKE_FLG = Trim(CheckDBNull(dr(25), enumObjectType.StrType))
                        TST_REDIRECT_URL = Trim(CheckDBNull(dr(26), enumObjectType.StrType))
                        TST_ORDER_ID = Trim(CheckDBNull(dr(27), enumObjectType.StrType))
                        CURRCLM_ID = Trim(CheckDBNull(dr(13), enumObjectType.StrType))
                        CRSE_TYPE_CD = Trim(CheckDBNull(dr(24), enumObjectType.StrType))
                        CRSE_NAME = Trim(CheckDBNull(dr(28), enumObjectType.StrType))
                        TRNR_REG_ID = Trim(CheckDBNull(dr(29), enumObjectType.StrType))
                        REG_ID = Trim(CheckDBNull(dr(31), enumObjectType.StrType))
                        SESS_PART_ID = Trim(CheckDBNull(dr(30), enumObjectType.StrType))
                        RETAKE_AUTH = Trim(CheckDBNull(dr(32), enumObjectType.StrType))
                        RESOLUTION = LCase(Trim(CheckDBNull(dr(33), enumObjectType.StrType)))
                        VAL_UID = Trim(CheckDBNull(dr(34), enumObjectType.StrType))
                        SCORM_FLG = Trim(CheckDBNull(dr(35), enumObjectType.StrType))
                        If POPUP_FLG = "" Then POPUP_FLG = Trim(CheckDBNull(dr(36), enumObjectType.StrType))
                        RESTART_RESET = Trim(CheckDBNull(dr(37), enumObjectType.StrType))
                        TEST_FORMAT = Trim(CheckDBNull(dr(38), enumObjectType.StrType))
                        EMAIL_ADDR = Trim(CheckDBNull(dr(39), enumObjectType.StrType))
                        temp = Trim(CheckDBNull(dr(40), enumObjectType.StrType))
                        If temp <> LANG_CD And LANG_CD = "" Then LANG_CD = temp
                        CRSE_CONTENT_URL = Trim(CheckDBNull(dr(41), enumObjectType.StrType))
                        X_ALT_RETAKE_FLG = Trim(CheckDBNull(dr(42), enumObjectType.StrType))
                        If TST_SURVEY_FLG = "Y" Then X_ALT_RETAKE_FLG = "N"
                        SR_RETAKE_FLG = Trim(CheckDBNull(dr(43), enumObjectType.StrType))
                        TST_CRSE_NAME = Trim(CheckDBNull(dr(44), enumObjectType.StrType))
                        If InStr(CRSE_CONTENT_URL, "https:") > 0 Then
                            LaunchProtocol = "https:"
                        End If                        
                    End While
                End If
            Catch ex As Exception
                GoTo AccessError
            End Try
            dr.Close()
            If Debug = "Y" Then
                mydebuglog.Debug("  .. EXM_CRSE_TST_ID: " & EXM_CRSE_TST_ID)
                mydebuglog.Debug("  .. EXM_PASSED_FLG: " & EXM_PASSED_FLG)
                mydebuglog.Debug("  .. EXM_TEST_DT: " & EXM_TEST_DT)
                mydebuglog.Debug("  .. EXM_CRSE_OFFR_ID: " & EXM_CRSE_OFFR_ID)
                mydebuglog.Debug("  .. EXM_FINAL_SCORE: " & EXM_FINAL_SCORE)
                mydebuglog.Debug("  .. EXM_GRADED_DT: " & EXM_GRADED_DT)
                mydebuglog.Debug("  .. EXM_CON_ID: " & EXM_CON_ID)
                mydebuglog.Debug("  .. EXM_STATUS_CD: " & EXM_STATUS_CD)
                mydebuglog.Debug("  .. EXM_PART_ID: " & EXM_PART_ID)
                mydebuglog.Debug("  .. EXM_ORDER_ITEM_ID: " & EXM_ORDER_ITEM_ID)
                mydebuglog.Debug("  .. EXM_PAR_TSTRUN_ID: " & EXM_PAR_TSTRUN_ID)
                mydebuglog.Debug("  .. EXM_CURRCLM_PER_ID: " & EXM_CURRCLM_PER_ID)
                mydebuglog.Debug("  .. TST_ID: " & TST_ID)
                mydebuglog.Debug("  .. TST_NAME: " & TST_NAME)
                mydebuglog.Debug("  .. TST_CRSE_ID: " & TST_CRSE_ID)
                mydebuglog.Debug("  .. TST_CRSE_NAME: " & TST_CRSE_NAME)
                mydebuglog.Debug("  .. TST_SKILL_LEVEL_CD: " & TST_SKILL_LEVEL_CD)
                mydebuglog.Debug("  .. TST_STATUS_CD: " & TST_STATUS_CD)
                mydebuglog.Debug("  .. TST_VERSION: " & TST_VERSION)
                mydebuglog.Debug("  .. TST_SURVEY_FLG: " & TST_SURVEY_FLG)
                mydebuglog.Debug("  .. TST_MAX_POINTS: " & TST_MAX_POINTS)
                mydebuglog.Debug("  .. TST_PASSING_SCORE: " & TST_PASSING_SCORE)
                mydebuglog.Debug("  .. TST_CONTINUABLE_FLG: " & TST_CONTINUABLE_FLG)
                mydebuglog.Debug("  .. CURRCLM_ID: " & CURRCLM_ID)
                mydebuglog.Debug("  .. CRSE_TYPE_CD: " & CRSE_TYPE_CD)
                mydebuglog.Debug("  .. REG_ID: " & REG_ID)
                mydebuglog.Debug("  .. SESS_PART_ID: " & SESS_PART_ID)
                mydebuglog.Debug("  .. RETAKE_AUTH: " & RETAKE_AUTH)
                mydebuglog.Debug("  .. RESOLUTION: " & RESOLUTION)
                mydebuglog.Debug("  .. VAL_UID: " & VAL_UID)
                mydebuglog.Debug("  .. SCORM_FLG: " & SCORM_FLG)
                mydebuglog.Debug("  .. POPUP_FLG: " & POPUP_FLG)
                mydebuglog.Debug("  .. RESTART_RESET: " & RESTART_RESET)
                mydebuglog.Debug("  .. TEST_FORMAT: " & TEST_FORMAT)
                mydebuglog.Debug("  .. EMAIL_ADDR: " & EMAIL_ADDR)
                mydebuglog.Debug("  .. LANG_CD: " & LANG_CD)
                mydebuglog.Debug("  .. CRSE_CONTENT_URL: " & CRSE_CONTENT_URL)
                mydebuglog.Debug("  .. SR_RETAKE_FLG: " & SR_RETAKE_FLG)
                mydebuglog.Debug("  .. X_ALT_RETAKE_FLG: " & X_ALT_RETAKE_FLG)
                mydebuglog.Debug("  .. LaunchProtocol: " & LaunchProtocol & vbCrLf)
            End If
            
            ' ================================================
            ' SET BASIC VALUES
            If TST_SURVEY_FLG = "Y" Then
                TYPE_CD = "Survey"
                Select Case LANG_CD
                    Case "ESN"
                        pTYPE_CD = "Encuesta"
                    Case Else
                        pTYPE_CD = "Survey"
                End Select
            Else
                TYPE_CD = "Exam"
                Select Case LANG_CD
                    Case "ESN"
                        pTYPE_CD = "Examen"
                    Case Else
                        pTYPE_CD = "Exam"
                End Select
            End If
            If Debug = "Y" Then mydebuglog.Debug("  .. TYPE_CD: " & TYPE_CD & ", pTYPE_CD: " & pTYPE_CD & vbCrLf)
            
            ' Compute base page
            If Debug = "Y" Then
                mydebuglog.Debug("  COMPUTING BASEPAGE: ")
                mydebuglog.Debug("  .. HOME_PAGE: " & HOME_PAGE)
                mydebuglog.Debug("  .. CURRENT_PAGE: " & CURRENT_PAGE)
                mydebuglog.Debug("  .. TYPE_CD: " & TYPE_CD)
                mydebuglog.Debug("  .. TST_SKILL_LEVEL_CD: " & TST_SKILL_LEVEL_CD)
            End If
            If HOME_PAGE = "certegrity.com" And CURRENT_PAGE <> "" Then
                BasePage = CURRENT_PAGE
                If RefreshID <> "" Then
                    BasePage = BasePage & "&RFR=" & RefreshID
                End If
            Else
                If LANG_CD <> "ENU" Then
                    BasePage = "https://" & HOME_PAGE & "/mobile/" & LANG_CD & "index.html"
                Else
                    BasePage = "https://" & HOME_PAGE & "/mobile/index.html"
                End If
                If InStr(BasePage, "UID=") = 0 Then BasePage = BasePage & "?UID=" & UID
            End If
            If InStr(BasePage, "SES=") = 0 Then BasePage = BasePage & "&SES=" & SessID
            If InStr(BasePage, "PP=") = 0 Then BasePage = BasePage & "&PP=" & DOMAIN
            If RefreshID = "" Then
                If TYPE_CD = "Survey" Then
                    If InStr(BasePage, "#cert") = 0 Then BasePage = BasePage & "#cert"
                Else
                    If InStr(BasePage, "#reg") = 0 Then BasePage = BasePage & "#reg"
                End If                
            End If
            If Debug = "Y" Then mydebuglog.Debug("  .. BasePage: " & BasePage & vbCrLf)
            
            ' ================================================
            ' PROCESS COMPLETED EXAMS   
            If EXM_STATUS_CD = "Graded" And TYPE_CD <> "Survey" Then
                PAUSE = "0"
                If SCORM_FLG = "Y" And REG_ID <> "" And POPUP_FLG = "" Then POPUP_FLG = "N"
      
                ' If given an automatic retake, then redurect to registration screen
                If EXM_PASSED_FLG = "N" And X_ALT_RETAKE_FLG = "Y" And SR_RETAKE_FLG <> "Y" Then
                    PAUSE = "60"
                    If HOME_PAGE = "certegrity.com" And CURRENT_PAGE <> "" Then
                        ClassLink = CURRENT_PAGE
                        If RefreshID <> "" Then
                            ClassLink = ClassLink & "&RFR=" & RefreshID
                        End If
                    Else
                        If LANG_CD <> "ENU" Then
                            ClassLink = BasePage & "?UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "#reg"
                        Else
                            ClassLink = BasePage & "?UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "#reg"
                        End If
                    End If
                Else
                    If LANG_CD <> "ENU" Then
                        ClassLink = "https://hciscorm.certegrity.com/ls/" & LANG_CD & "/OpenCertificate.html?RID=" & REG_ID & "&TID=" & CRSE_TSTRUN_ID & "&UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "&LANG=" & LANG_CD & "&HP=" & HOME_PAGE & "&RFR=" & RefreshID & "&CUR=" & CURRENT_PAGE
                    Else
                        ClassLink = "https://hciscorm.certegrity.com/ls/OpenCertificate.html?RID=" & REG_ID & "&TID=" & CRSE_TSTRUN_ID & "&UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "&LANG=" & LANG_CD & "&HP=" & HOME_PAGE & "&RFR=" & RefreshID & "&CUR=" & CURRENT_PAGE
                    End If
                End If
                If Debug = "Y" Then mydebuglog.Debug("  .. ClassLink: " & ClassLink & vbCrLf)
         
                ' Update count of exams in activity summary
                If Not DupeCall Then
                    SqlS = "UPDATE siebeldb.dbo.CX_SUB_CON " & _
                    "SET NEW_EXM=U.CNT " & _
                    "FROM (SELECT COUNT(*) AS CNT " & _
                    "FROM siebeldb.dbo.S_CRSE_TSTRUN A " & _
                    "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TST T ON A.CRSE_TST_ID=T.ROW_ID " & _
                    "WHERE A.PERSON_ID='" & EXM_CON_ID & "' AND T.X_SURVEY_FLG='N' AND " & _
                    "(A.X_PAR_TSTRUN_ID IS NULL OR A.X_PAR_TSTRUN_ID='') AND " & _
                    "A.STATUS_CD IN ('Pending','Paymt Reqd','Unsubmitted','Incomplete')) U " & _
                    "WHERE CON_ID='" & EXM_CON_ID & "'"
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  UPDATE ACTIVITY SUMMARY: " & vbCrLf & "  " & SqlS)
                    temp = ExecQuery("Update", "CX_SUB_CON", cmd, SqlS, mydebuglog, "N")
                End If
            End If
            
            ' ================================================
            ' PROCESS INCOMPLETE EXAMS
            '   Allow the user to open the assessment
            If LaunchProtocol = "" Then LaunchProtocol = "http:"
            If Debug = "Y" Then
                mydebuglog.Debug("  COMPUTING TESTINGURL: ")
                mydebuglog.Debug("  .. EXM_STATUS_CD: " & EXM_STATUS_CD)
                mydebuglog.Debug("  .. LaunchProtocol: " & LaunchProtocol)
                mydebuglog.Debug("  .. VAL_UID: " & VAL_UID)
                mydebuglog.Debug("  .. UID$: " & UID$)
            End If
            If EXM_STATUS_CD = "Pending" Or EXM_STATUS_CD = "Unsubmitted" Then
                If VAL_UID = UID$ Then
                    If RESOLUTION = "" Then RESOLUTION = "800x600"
                    RESOLUTION = Replace(RESOLUTION, "x", ",")
                    If LANG_CD <> "ENU" Then
                        testingURL = LaunchProtocol & "//hciscorm.certegrity.com/ls/" & LANG_CD & "/OpenAssessment.html?RID=" & REG_ID & "&EID=" & CRSE_TSTRUN_ID & "&TID=" & TST_ID & "&UID=" & UID$ & "&SES=" & SessID$ & "&PP=" & DOMAIN & "&LANG=" & LANG_CD & "&HP=" & HOME_PAGE & "&RFR=" & RefreshID & "&FNC=Y&CUR=" & CURRENT_PAGE
                    Else
                        testingURL = LaunchProtocol & "//hciscorm.certegrity.com/ls/OpenAssessment.html?RID=" & REG_ID & "&EID=" & CRSE_TSTRUN_ID & "&TID=" & TST_ID & "&UID=" & UID$ & "&SES=" & SessID$ & "&PP=" & DOMAIN & "&LANG=" & LANG_CD & "&HP=" & HOME_PAGE & "&RFR=" & RefreshID & "&FNC=Y&CUR=" & CURRENT_PAGE
                    End If
                    PAUSE = "20"
                    If Debug = "Y" Then mydebuglog.Debug("  .. testingURL: " & testingURL & vbCrLf)
                    GoTo GenerateResults
                End If
            End If
            
            ' ================================================
            ' PROCESS COMPLETED SURVEYS
            If EXM_STATUS_CD = "Graded" And TYPE_CD = "Survey" Then
                PAUSE = "30"
                If CURRENT_PAGE <> "" Then
                    ClassLink = CURRENT_PAGE
                    If RefreshID <> "" Then
                        ClassLink = ClassLink & "&RFR=" & RefreshID
                    Else
                        If InStr(ClassLink, "UID=") = 0 Then ClassLink = ClassLink & "?UID=" & UID
                        If InStr(ClassLink, "SES=") = 0 Then ClassLink = ClassLink & "&SES=" & SessID
                        If InStr(ClassLink, "PP=") = 0 Then ClassLink = ClassLink & "&PP=" & DOMAIN
                        If InStr(ClassLink, "#cert") = 0 Then ClassLink = ClassLink & "#cert"
                    End If
                Else
                    If HOME_PAGE = "certegrity.com" And CURRENT_PAGE <> "" Then
                        ClassLink = CURRENT_PAGE
                        If RefreshID <> "" Then
                            ClassLink = ClassLink & "&RFR=" & RefreshID
                        End If
                    Else
                        If LANG_CD <> "ENU" Then
                            ClassLink = BasePage & "?UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "#cert"
                        Else
                            ClassLink = BasePage & "?UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "#cert"
                        End If
                    End If
                End If
                If Debug = "Y" Then mydebuglog.Debug("  .. ClassLink: " & ClassLink & vbCrLf)
         
                ' Update count of surveys in activity summary
                If Not DupeCall Then
                    ' Update survey count
                    SqlS = "UPDATE siebeldb.dbo.CX_SUB_CON " & _
                    "SET NEW_SRV=U.CNT " & _
                    "FROM (SELECT COUNT(*) AS CNT " & _
                    "FROM siebeldb.dbo.S_CRSE_TSTRUN A " & _
                    "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TST T ON A.CRSE_TST_ID=T.ROW_ID " & _
                    "WHERE A.PERSON_ID='" & EXM_CON_ID & "' AND T.X_SURVEY_FLG='Y' AND A.STATUS_CD IN ('Unsubmitted','Pending')) U " & _
                    "WHERE CON_ID='" & EXM_CON_ID & "'"
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  UPDATE ACTIVITY SUMMARY: " & vbCrLf & "  " & SqlS)
                    temp = ExecQuery("Update", "CX_SUB_CON", cmd, SqlS, mydebuglog, "N")
         
                    ' Send thank you email
                    SendFrom = "techsupport@gettips.com"
                    SendTo = EMAIL_ADDR
                    If LANG_CD = "" Then LANG_CD = "ENU"
                    If TST_CRSE_NAME = "None" Then TST_CRSE_ID = ""
                    MergeData = "<messages>" & _
                    "<message send_to=""" & SendTo & """ send_from=""" & SendFrom & """ from_name=""Customer Service"" from_id="""" to_id=""" & CONTACT_ID & """>" & _
                    "<REG_ID>" & TST_CRSE_ID & "</REG_ID>" & _
                    "<DOMAIN>" & DOMAIN & "</DOMAIN>" & _
                    "</message>" & _
                    "</messages>"
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  MESSAGE 0025- " & vbCrLf & MergeData)
                    Results = Processing.EmailXsltMerge(MergeData, "669", "669", "EMAIL", LANG_CD, Debug)
                    If Debug = "Y" Then mydebuglog.Debug("  .. EmailXsltMerge: " & Results & vbCrLf)
                    GoTo GenerateResults
                End If
            End If
            
            ' ================================================
            ' PROCESS UNGRADED ASSESSMENTS
            '   Score the submitted assessment by invoking the ExamScoreStore service
            If Not DupeCall Then
                If EXM_STATUS_CD = "Submitted" Then
                    NumIts = NumIts + 1
                    If Debug = "Y" Then mydebuglog.Debug("  PROCESSING SUBMITTED EXAM ATTEMPT #: " & Str(NumIts))
         
                    ' Pause and then re-check status 
                    Threading.Thread.Sleep(sleepcount * 100)
                    SqlS = "SELECT STATUS_CD " & _
                    "FROM siebeldb.dbo.S_CRSE_TSTRUN " & _
                    "WHERE ROW_ID='" & CRSE_TSTRUN_ID & "'"
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  RETRIEVE TEST STATUS: " & vbCrLf & "  " & SqlS)
                    cmd.CommandText = SqlS
                    Try
                        EXM_STATUS_CD = CheckDBNull(cmd.ExecuteScalar(), enumObjectType.StrType)
                    Catch ex As Exception
                        GoTo AccessError
                    End Try
                    If Debug = "Y" Then mydebuglog.Debug("  .. EXM_STATUS_CD : " & EXM_STATUS_CD)

                    ' If not processed all the way then remove S_CRSE_TSTRUN_ACCESS 
                    ' records so we can re-run the process
                    If EXM_STATUS_CD = "Submitted" Then
                        SqlS = "DELETE FROM siebeldb.dbo.S_CRSE_TSTRUN_ACCESS WHERE CRSE_TSTRUN_ID='" & CRSE_TSTRUN_ID & "' AND EXIT_FLG='Y'"
                        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  REMOVE S_CRSE_TSTRUN_ACCESS RECORDS: " & vbCrLf & "  " & SqlS)
                        temp = ExecQuery("Delete", "S_CRSE_TSTRUN_ACCESS", cmd, SqlS, mydebuglog, "N")
            
                        ' Execute ExamScoreStore
                        '	service = "http://cloudsvc.certegrity.com/processing/service.asmx/ExamScoreStore?TestId=" & ExamId & 
                        '   "&RegId=" & RegId & "&UID=" & temp2 & "&BatchTable=&BatchRec=&ProcessType=C&debug=" & debugflag
                        
                        EncodedUID = EncodeUID(UID)
                        If Debug = "Y" Then mydebuglog.Debug("  .. EncodedUID : " & EncodedUID)
                        Results = Processing.ExamScoreStore(CRSE_TSTRUN_ID, REG_ID, "", "", "C", EncodedUID, Debug)
                        If Debug = "Y" Then mydebuglog.Debug("  .. ExamScoreStore: " & Results & vbCrLf)
                        sleepcount = sleepcount + 1
                        If NumIts < 6 Then GoTo ExecQuery
                    End If
                End If
            End If
            
            ' ================================================
            ' REPORT STATUS
            ' Presume that the status code needs no further processing - performed by ExamScoreStore
GenerateResults:
            ' Status dependent messages
            Select Case EXM_STATUS_CD
      
                ' ------
                ' Test completed - go to results screen
                Case "Graded"
                    If PAUSE <> "0" Then
                        If EXM_PASSED_FLG = "N" And X_ALT_RETAKE_FLG = "Y" And SR_RETAKE_FLG <> "Y" Then
                            Select Case LANG_CD
                                Case "ESN"
                                    MSG1 = "Su examen fue completado Lamentablemente, usted no pas&oacute;.<br/>Un examen de repetici&oacute;n est&aacute; disponible para usted. <br><br>"
                                    MSG1 = MSG1 & "<br><br><a href=""" & ClassLink & """ data-role=""button"" id=""ProceedBtn"" rel=""external"" data-theme=""b"">Haga clic para volver al portal</a>"
                                Case Else
                                    MSG1 = "Your exam was completed. Unfortunately, you did not pass.<br/>A retake exam is available to you.<br><br>"
                                    MSG1 = MSG1 & "<br><br><a href=""" & ClassLink & """ data-role=""button"" id=""ProceedBtn"" rel=""external"" data-theme=""b"">Click to return to the portal</a>"
                            End Select
                        Else
                            Select Case LANG_CD
                                Case "ESN"
                                    MSG1 = "Tu " & LCase(pTYPE_CD) & " se complet&oacute;. <br><br>"
                                    MSG1 = MSG1 & "Gracias"
                                    If TYPE_CD = "Survey" Then
                                        MSG1 = MSG1 & " para proporcionar comentarios."
                                        MSG1 = MSG1 & "<br><br><a href=""" & ClassLink & """ data-role=""button"" id=""ProceedBtn"" rel=""external"" data-theme=""b"">Haga clic para volver al portal</a>"
                                    End If
                                Case Else
                                    MSG1 = "Your " & LCase(pTYPE_CD) & " was completed. <br><br>"
                                    MSG1 = MSG1 & "Thank you"
                                    If TYPE_CD = "Survey" Then
                                        MSG1 = MSG1 & " for providing feedback."
                                        MSG1 = MSG1 & "<br><br><a href=""" & ClassLink & """ data-role=""button"" id=""ProceedBtn"" rel=""external"" data-theme=""b"">Click to return to the portal</a>"
                                    End If
                            End Select
                        End If
                    Else
                        Select Case LANG_CD
                            Case "ESN"
                                MSG1 = MSG1 & "Un momento por favor..."
                            Case Else
                                MSG1 = MSG1 & "One moment please..."
                        End Select
                    End If
      
                    ' ------
                    ' Test not taken.. re-enter the testing system
                Case "Pending"
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = "Este " & LCase(pTYPE_CD) & " no se ha iniciado.<br><br>"
                            MSG1 = MSG1 & "<a href=""" & testingURL & """ data-role=""button"" id=""ProceedBtn"" rel=""external"" data-theme=""b"">Haga clic aqu&iacute; para comenzar este " & TYPE_CD & "</a><br>"
                        Case Else
                            MSG1 = "This " & LCase(pTYPE_CD) & " has not been started.<br><br>"
                            MSG1 = MSG1 & "<a href=""" & testingURL & """ data-role=""button"" id=""ProceedBtn"" rel=""external"" data-theme=""b"">Click here to begin this " & pTYPE_CD & "</a><br>"
                    End Select
                    If HOME_PAGE = "certegrity.com" And CURRENT_PAGE <> "" Then
                        NextLink = CURRENT_PAGE
                        If RefreshID <> "" Then
                            NextLink = NextLink & "&RFR=" & RefreshID
                        End If
                    Else
                        If Debug = "Y" Then
                            mydebuglog.Debug("  COMPUTING EXIT: ")
                            mydebuglog.Debug("  .. CURRENT_PAGE: " & CURRENT_PAGE)
                            mydebuglog.Debug("  .. TYPE_CD: " & TYPE_CD)
                            mydebuglog.Debug("  .. TST_SKILL_LEVEL_CD: " & TST_SKILL_LEVEL_CD)
                        End If
                        If CURRENT_PAGE <> "" Then
                            NextLink = CURRENT_PAGE
                            If InStr(NextLink, "UID=") = 0 Then NextLink = NextLink & "?UID=" & UID
                            If InStr(NextLink, "SES=") = 0 Then NextLink = NextLink & "&SES=" & SessID
                            If InStr(NextLink, "PP=") = 0 Then NextLink = NextLink & "&PP=" & DOMAIN
                            If RefreshID <> "" Then NextLink = NextLink & "&RFR=" & RefreshID
                            If TYPE_CD = "Survey" Then
                                If InStr(NextLink, "#cert") = 0 Then NextLink = NextLink & "#cert"
                            Else
                                If InStr(NextLink, "#reg") = 0 Then NextLink = NextLink & "#reg"
                            End If
                        Else
                            If LANG_CD <> "ENU" Then
                                NextLink = "https://" & HOME_PAGE & "/mobile/" & LANG_CD & "/index.html" & "?UID=" & UID & "&SES=" & SessID
                            Else
                                NextLink = "https://" & HOME_PAGE & "/mobile/index.html" & "?UID=" & UID & "&SES=" & SessID
                            End If
                            If TYPE_CD = "Survey" Then NextLink = NextLink & "#cert" Else NextLink = NextLink & "#reg"
                        End If
                    End If
                    
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = MSG1 & "<br><br><a href=" & NextLink & " data-role=""button"" id=""ProceedBtn"" rel=""external"" data-theme=""b"">Haga clic para volver al portal</a>"
                        Case Else
                            MSG1 = MSG1 & "<br><br><a href=" & NextLink & " data-role=""button"" id=""ProceedBtn"" rel=""external"" data-theme=""b"">Click to return to the portal</a>"
                    End Select

                    ' ------
                    ' Test not completed.. re-enter the testing system
                Case "Unsubmitted"
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = "Este " & LCase(pTYPE_CD) & " no se ha completado, "
                            NextLink = "<a href=""" & LaunchProtocol & "//hciscorm.certegrity.com/ls/" & LANG_CD & "/OpenAssessment.html?EID=" & CRSE_TSTRUN_ID & "&CID=&UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "&LANG=" & LANG_CD & "&SRC=&CUR=" & CURRENT_PAGE & """ data-role=""button"" rel=""external"" data-theme=""b"" title=\u0027Haga clic para realizar este examen.\u0027>"
                        Case Else
                            MSG1 = "This " & LCase(pTYPE_CD) & " has not been completed, "
                            NextLink = "<a href=""" & LaunchProtocol & "//hciscorm.certegrity.com/ls/OpenAssessment.html?EID=" & CRSE_TSTRUN_ID & "&UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "&CID=&LANG=" & LANG_CD & "&SRC=&CUR=" & CURRENT_PAGE & """ data-role=""button"" rel=""external"" data-theme=""b"" title=\u0027Click to take this exam\u0027>"
                    End Select
         
                    If TST_CONTINUABLE_FLG = "Y" Then
                        If RESTART_RESET = "Y" Then
                            Select Case LANG_CD
                                Case "ESN"
                                    MSG1 = MSG1 & "y tendr&aacute;s que reiniciar el examen desde el principio porque saliste.<br><br>"
                                Case Else
                                    MSG1 = MSG1 & "and you will need to restart the exam from the beginning because you exited it.<br><br>"
                            End Select
                        Else
                            Select Case LANG_CD
                                Case "ESN"
                                    MSG1 = MSG1 & "pero puedes continuarla donde fue interrumpida.<br><br>"
                                Case Else
                                    MSG1 = MSG1 & "but you may continue it where it was interrupted.<br><br>"
                            End Select
                        End If
                        If testingURL <> "" Then
                            Select Case LANG_CD
                                Case "ESN"
                                    MSG1 = MSG1 & "<a href=" & testingURL & " title=""Haga clic para iniciar o continuar su " & pTYPE_CD & """ id=""ProceedBtn"" data-role=""button"" rel=""external"" data-theme=""b"">Haga clic aqu&iacute; para continuar esto " & pTYPE_CD & "</a><br>"
                                Case Else
                                    MSG1 = MSG1 & "<a href=" & testingURL & " title=""Click to start or continue your " & pTYPE_CD & """ id=""ProceedBtn"" data-role=""button"" rel=""external"" data-theme=""b"">Click here to continue this " & pTYPE_CD & "</a><br>"
                            End Select
                        End If
                    Else
                        Select Case LANG_CD
                            Case "ESN"
                                MSG1 = MSG1 & "y debe ser completado en una sola sesi&oacute;n.<br><br>"
                                If testingURL <> "" Then
                                    MSG1 = MSG1 & "<a href=" & testingURL & " title=""Haga clic para iniciar o continuar su " & pTYPE_CD & """ id=""ProceedBtn"" data-role=""button"" rel=""external"" data-theme=""b"">Haga clic aqu&iacute; para reiniciar esto " & pTYPE_CD & "</A>"
                                End If
                            Case Else
                                MSG1 = MSG1 & "and must be completed in one sitting.<br><br>"
                                If testingURL <> "" Then
                                    MSG1 = MSG1 & "<a href=" & testingURL & " title=""Click to start or continue your " & pTYPE_CD & """ id=""ProceedBtn"" data-role=""button"" rel=""external"" data-theme=""b"">Click here to restart this " & pTYPE_CD & "</A>"
                                End If
                        End Select
                    End If
                    If HOME_PAGE = "certegrity.com" And CURRENT_PAGE <> "" Then
                        NextLink = CURRENT_PAGE
                        If RefreshID <> "" Then
                            NextLink = NextLink & "&RFR=" & RefreshID
                        End If
                    Else
                        If Debug = "Y" Then
                            mydebuglog.Debug("  COMPUTING EXIT: ")
                            mydebuglog.Debug("  .. CURRENT_PAGE: " & CURRENT_PAGE)
                            mydebuglog.Debug("  .. TYPE_CD: " & TYPE_CD)
                            mydebuglog.Debug("  .. TST_SKILL_LEVEL_CD: " & TST_SKILL_LEVEL_CD)
                        End If
                        If CURRENT_PAGE <> "" Then
                            NextLink = CURRENT_PAGE
                            If InStr(NextLink, "UID=") = 0 Then NextLink = NextLink & "?UID=" & UID
                            If InStr(NextLink, "SES=") = 0 Then NextLink = NextLink & "&SES=" & SessID
                            If InStr(NextLink, "PP=") = 0 Then NextLink = NextLink & "&PP=" & DOMAIN
                            If RefreshID <> "" Then NextLink = NextLink & "&RFR=" & RefreshID
                            If TYPE_CD = "Survey" Then
                                If InStr(NextLink, "#cert") = 0 Then NextLink = NextLink & "#cert"
                            Else
                                If InStr(NextLink, "#reg") = 0 Then NextLink = NextLink & "#reg"
                            End If
                        Else
                            If LANG_CD <> "ENU" Then
                                NextLink = "https://" & HOME_PAGE & "/mobile/" & LANG_CD & "/index.html" & "?UID=" & UID & "&SES=" & SessID
                            Else
                                NextLink = "https://" & HOME_PAGE & "/mobile/index.html" & "?UID=" & UID & "&SES=" & SessID
                            End If
                            If TYPE_CD = "Survey" Then NextLink = NextLink & "#cert" Else NextLink = NextLink & "#reg"
                        End If
                    End If

                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = MSG1 & "<a href=""" & NextLink & """ data-role=""button"" id=""ProceedBtn"" rel=""external"" data-theme=""b"">Haga clic para volver al portal</a>"
                        Case Else
                            MSG1 = MSG1 & "<a href=""" & NextLink & """ data-role=""button"" id=""ProceedBtn"" rel=""external"" data-theme=""b"">Click to return to the portal</a>"
                    End Select
      
                    ' ------
                    ' We should never get this status here - if we do it means that ExamScoreStore is broken
                Case "Submitted"
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = "Ha habido un error<br><br>"
                        Case Else
                            MSG1 = "There has been an error<br><br>"
                    End Select
                    If HOME_PAGE = "certegrity.com" And CURRENT_PAGE <> "" Then
                        NextLink = CURRENT_PAGE
                        If RefreshID <> "" Then
                            NextLink = NextLink & "&RFR=" & RefreshID
                        End If
                    Else
                        If Debug = "Y" Then
                            mydebuglog.Debug("  COMPUTING EXIT: ")
                            mydebuglog.Debug("  .. CURRENT_PAGE: " & CURRENT_PAGE)
                            mydebuglog.Debug("  .. TYPE_CD: " & TYPE_CD)
                            mydebuglog.Debug("  .. TST_SKILL_LEVEL_CD: " & TST_SKILL_LEVEL_CD)
                        End If
                        If CURRENT_PAGE <> "" Then
                            NextLink = CURRENT_PAGE
                            If InStr(NextLink, "UID=") = 0 Then NextLink = NextLink & "?UID=" & UID
                            If InStr(NextLink, "SES=") = 0 Then NextLink = NextLink & "&SES=" & SessID
                            If InStr(NextLink, "PP=") = 0 Then NextLink = NextLink & "&PP=" & DOMAIN
                            If RefreshID <> "" Then NextLink = NextLink & "&RFR=" & RefreshID
                            If TYPE_CD = "Survey" Then
                                If InStr(NextLink, "#cert") = 0 Then NextLink = NextLink & "#cert"
                            Else
                                If InStr(NextLink, "#reg") = 0 Then NextLink = NextLink & "#reg"
                            End If
                        Else
                            If LANG_CD <> "ENU" Then
                                NextLink = "https://" & HOME_PAGE & "/mobile/" & LANG_CD & "/index.html" & "?UID=" & UID & "&SES=" & SessID
                            Else
                                NextLink = "https://" & HOME_PAGE & "/mobile/index.html" & "?UID=" & UID & "&SES=" & SessID
                            End If
                            If TYPE_CD = "Survey" Then NextLink = NextLink & "#cert" Else NextLink = NextLink & "#reg"
                        End If
                    End If

                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = MSG1 & "<a href=""" & NextLink & """ data-role=""button"" id=""ProceedBtn"" rel=""external"" data-theme=""b"">Haga clic para volver al portal</a>"
                        Case Else
                            MSG1 = MSG1 & "<a href=""" & NextLink & """ data-role=""button"" id=""ProceedBtn"" rel=""external"" data-theme=""b"">Click to return to the portal</a>"
                    End Select
      
                    ' ------
                    ' We should never get this status here - if we do it means that ExamScoreStore is broken
                Case "Ungraded"
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = "Ha habido un error<br><br>"
                        Case Else
                            MSG1 = "There has been an error<br><br>"
                    End Select
                    If HOME_PAGE = "certegrity.com" And CURRENT_PAGE <> "" Then
                        NextLink = CURRENT_PAGE
                        If RefreshID <> "" Then
                            NextLink = NextLink & "&RFR=" & RefreshID
                        End If
                    Else
                        If Debug = "Y" Then
                            mydebuglog.Debug("  COMPUTING EXIT: ")
                            mydebuglog.Debug("  .. CURRENT_PAGE: " & CURRENT_PAGE)
                            mydebuglog.Debug("  .. TYPE_CD: " & TYPE_CD)
                            mydebuglog.Debug("  .. TST_SKILL_LEVEL_CD: " & TST_SKILL_LEVEL_CD)
                        End If
                        If CURRENT_PAGE <> "" Then
                            NextLink = CURRENT_PAGE
                            If InStr(NextLink, "UID=") = 0 Then NextLink = NextLink & "?UID=" & UID
                            If InStr(NextLink, "SES=") = 0 Then NextLink = NextLink & "&SES=" & SessID
                            If InStr(NextLink, "PP=") = 0 Then NextLink = NextLink & "&PP=" & DOMAIN
                            If RefreshID <> "" Then NextLink = NextLink & "&RFR=" & RefreshID
                            If TYPE_CD = "Survey" Then
                                If InStr(NextLink, "#cert") = 0 Then NextLink = NextLink & "#cert"
                            Else
                                If InStr(NextLink, "#reg") = 0 Then NextLink = NextLink & "#reg"
                            End If
                        Else
                            If LANG_CD <> "ENU" Then
                                NextLink = "https://" & HOME_PAGE & "/mobile/" & LANG_CD & "/index.html" & "?UID=" & UID & "&SES=" & SessID
                            Else
                                NextLink = "https://" & HOME_PAGE & "/mobile/index.html" & "?UID=" & UID & "&SES=" & SessID
                            End If
                            If TYPE_CD = "Survey" Then NextLink = NextLink & "#cert" Else NextLink = NextLink & "#reg"
                        End If
                    End If

                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = MSG1 & "<a href=""" & NextLink & """ data-role=""button"" id=""ProceedBtn"" rel=""external"" data-theme=""b"">Haga clic para volver al portal</a>"
                        Case Else
                            MSG1 = MSG1 & "<a href=""" & NextLink & """ data-role=""button"" id=""ProceedBtn"" rel=""external"" data-theme=""b"">Click to return to the portal</a>"
                    End Select
      
                    ' ------
                Case "Cancelled"         ' Cannot grade
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = "Este " & pTYPE_CD & " ha sido cancelado.<br><br>"
                        Case Else
                            MSG1 = "This " & pTYPE_CD & " has been cancelled.<br><br>"
                    End Select
                    If HOME_PAGE = "certegrity.com" And CURRENT_PAGE <> "" Then
                        NextLink = CURRENT_PAGE
                        If RefreshID <> "" Then
                            NextLink = NextLink & "&RFR=" & RefreshID
                        End If
                    Else
                        If Debug = "Y" Then
                            mydebuglog.Debug("  COMPUTING EXIT: ")
                            mydebuglog.Debug("  .. CURRENT_PAGE: " & CURRENT_PAGE)
                            mydebuglog.Debug("  .. TYPE_CD: " & TYPE_CD)
                            mydebuglog.Debug("  .. TST_SKILL_LEVEL_CD: " & TST_SKILL_LEVEL_CD)
                        End If
                        If CURRENT_PAGE <> "" Then
                            NextLink = CURRENT_PAGE
                            If InStr(NextLink, "UID=") = 0 Then NextLink = NextLink & "?UID=" & UID
                            If InStr(NextLink, "SES=") = 0 Then NextLink = NextLink & "&SES=" & SessID
                            If InStr(NextLink, "PP=") = 0 Then NextLink = NextLink & "&PP=" & DOMAIN
                            If RefreshID <> "" Then NextLink = NextLink & "&RFR=" & RefreshID
                            If TYPE_CD = "Survey" Then
                                If InStr(NextLink, "#cert") = 0 Then NextLink = NextLink & "#cert"
                            Else
                                If InStr(NextLink, "#reg") = 0 Then NextLink = NextLink & "#reg"
                            End If
                        Else
                            If LANG_CD <> "ENU" Then
                                NextLink = "https://" & HOME_PAGE & "/mobile/" & LANG_CD & "/index.html" & "?UID=" & UID & "&SES=" & SessID
                            Else
                                NextLink = "https://" & HOME_PAGE & "/mobile/index.html" & "?UID=" & UID & "&SES=" & SessID
                            End If
                            If TYPE_CD = "Survey" Then NextLink = NextLink & "#cert" Else NextLink = NextLink & "#reg"
                        End If
                    End If
                    
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = MSG1 & "<a href=""" & NextLink & """ data-role=""button"" id=""ProceedBtn"" rel=""external"" data-theme=""b"">Haga clic para volver al portal</a>"
                        Case Else
                            MSG1 = MSG1 & "<a href=""" & NextLink & """ data-role=""button"" id=""ProceedBtn"" rel=""external"" data-theme=""b"">Click to return to the portal</a>"
                    End Select
            
                    ' ------
                Case "Incomplete"
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = "Este " & LCase(pTYPE_CD) & " no se ha completado, pero ha alcanzado el n&uacute;mero permitido de intentos y es posible que no contin&uacute;e.<br><br>"
                        Case Else
                            MSG1 = "This " & LCase(pTYPE_CD) & " has not been completed but you have reached the allowed number of attempts and may not continue.<br><br>"
                    End Select
         
                    If HOME_PAGE = "certegrity.com" And CURRENT_PAGE <> "" Then
                        NextLink = CURRENT_PAGE
                        If RefreshID <> "" Then
                            NextLink = NextLink & "&RFR=" & RefreshID
                        End If
                    Else
                        If Debug = "Y" Then
                            mydebuglog.Debug("  COMPUTING EXIT: ")
                            mydebuglog.Debug("  .. CURRENT_PAGE: " & CURRENT_PAGE)
                            mydebuglog.Debug("  .. TYPE_CD: " & TYPE_CD)
                            mydebuglog.Debug("  .. TST_SKILL_LEVEL_CD: " & TST_SKILL_LEVEL_CD)
                        End If
                        If CURRENT_PAGE <> "" Then
                            NextLink = CURRENT_PAGE
                            If InStr(NextLink, "UID=") = 0 Then NextLink = NextLink & "?UID=" & UID
                            If InStr(NextLink, "SES=") = 0 Then NextLink = NextLink & "&SES=" & SessID
                            If InStr(NextLink, "PP=") = 0 Then NextLink = NextLink & "&PP=" & DOMAIN
                            If RefreshID <> "" Then NextLink = NextLink & "&RFR=" & RefreshID
                            If TYPE_CD = "Survey" Then
                                If InStr(NextLink, "#cert") = 0 Then NextLink = NextLink & "#cert"
                            Else
                                If InStr(NextLink, "#reg") = 0 Then NextLink = NextLink & "#reg"
                            End If
                        Else
                            If LANG_CD <> "ENU" Then
                                NextLink = "https://" & HOME_PAGE & "/mobile/" & LANG_CD & "/index.html" & "?UID=" & UID & "&SES=" & SessID
                            Else
                                NextLink = "https://" & HOME_PAGE & "/mobile/index.html" & "?UID=" & UID & "&SES=" & SessID
                            End If
                            If TYPE_CD = "Survey" Then NextLink = NextLink & "#cert" Else NextLink = NextLink & "#reg"
                        End If
                    End If
                    
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = MSG1 & "<a href=""" & NextLink & """ data-role=""button"" id=""ProceedBtn"" rel=""external"" data-theme=""b"">Haga clic para volver al portal</a>"
                        Case Else
                            MSG1 = MSG1 & "<a href=""" & NextLink & """ data-role=""button"" id=""ProceedBtn"" rel=""external"" data-theme=""b"">Click to return to the portal</a>"
                    End Select
      
                    ' ------
                    ' The assessment has been put on hold due to a KBA issue 
                Case "On-Hold"
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = "Su evaluaci&#243;n ha sido puesta en espera pendiente de revisi&#243;n<br><br>"
                        Case Else
                            MSG1 = "Your assessment has been placed On Hold pending a review<br><br>"
                    End Select
                    
                    If HOME_PAGE = "certegrity.com" And CURRENT_PAGE <> "" Then
                        NextLink = CURRENT_PAGE
                        If RefreshID <> "" Then
                            NextLink = NextLink & "&RFR=" & RefreshID
                        End If
                    Else
                        If Debug = "Y" Then
                            mydebuglog.Debug("  COMPUTING EXIT: ")
                            mydebuglog.Debug("  .. CURRENT_PAGE: " & CURRENT_PAGE)
                            mydebuglog.Debug("  .. TYPE_CD: " & TYPE_CD)
                            mydebuglog.Debug("  .. TST_SKILL_LEVEL_CD: " & TST_SKILL_LEVEL_CD)
                        End If
                        If CURRENT_PAGE <> "" Then
                            NextLink = CURRENT_PAGE
                            If InStr(NextLink, "UID=") = 0 Then NextLink = NextLink & "?UID=" & UID
                            If InStr(NextLink, "SES=") = 0 Then NextLink = NextLink & "&SES=" & SessID
                            If InStr(NextLink, "PP=") = 0 Then NextLink = NextLink & "&PP=" & DOMAIN
                            If RefreshID <> "" Then NextLink = NextLink & "&RFR=" & RefreshID
                            If TYPE_CD = "Survey" Then
                                If InStr(NextLink, "#cert") = 0 Then NextLink = NextLink & "#cert"
                            Else
                                If InStr(NextLink, "#reg") = 0 Then NextLink = NextLink & "#reg"
                            End If
                        Else
                            If LANG_CD <> "ENU" Then
                                NextLink = "https://" & HOME_PAGE & "/mobile/" & LANG_CD & "/index.html" & "?UID=" & UID & "&SES=" & SessID
                            Else
                                NextLink = "https://" & HOME_PAGE & "/mobile/index.html" & "?UID=" & UID & "&SES=" & SessID
                            End If
                            If TYPE_CD = "Survey" Then NextLink = NextLink & "#cert" Else NextLink = NextLink & "#reg"
                        End If
                    End If
                    
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = MSG1 & "<a href=""" & NextLink & """ data-role=""button"" id=""ProceedBtn"" rel=""external"" data-theme=""b"">Haga clic para volver al portal</a>"
                        Case Else
                            MSG1 = MSG1 & "<a href=""" & NextLink & """ data-role=""button"" id=""ProceedBtn"" rel=""external"" data-theme=""b"">Click to return to the portal</a>"
                    End Select
                    
                    ' ------
                    ' We should never get this status here - if we do it means that ExamScoreStore is broken
                Case "Error"
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = "Ha habido un error<br><br>"
                        Case Else
                            MSG1 = "There has been an error<br><br>"
                    End Select
      
                    If HOME_PAGE = "certegrity.com" And CURRENT_PAGE <> "" Then
                        NextLink = CURRENT_PAGE
                        If RefreshID <> "" Then
                            NextLink = NextLink & "&RFR=" & RefreshID
                        End If
                    Else
                        If Debug = "Y" Then
                            mydebuglog.Debug("  COMPUTING EXIT: ")
                            mydebuglog.Debug("  .. CURRENT_PAGE: " & CURRENT_PAGE)
                            mydebuglog.Debug("  .. TYPE_CD: " & TYPE_CD)
                            mydebuglog.Debug("  .. TST_SKILL_LEVEL_CD: " & TST_SKILL_LEVEL_CD)
                        End If
                        If CURRENT_PAGE <> "" Then
                            NextLink = CURRENT_PAGE
                            If InStr(NextLink, "UID=") = 0 Then NextLink = NextLink & "?UID=" & UID
                            If InStr(NextLink, "SES=") = 0 Then NextLink = NextLink & "&SES=" & SessID
                            If InStr(NextLink, "PP=") = 0 Then NextLink = NextLink & "&PP=" & DOMAIN
                            If RefreshID <> "" Then NextLink = NextLink & "&RFR=" & RefreshID
                            If TYPE_CD = "Survey" Then
                                If InStr(NextLink, "#cert") = 0 Then NextLink = NextLink & "#cert"
                            Else
                                If InStr(NextLink, "#reg") = 0 Then NextLink = NextLink & "#reg"
                            End If
                        Else
                            If LANG_CD <> "ENU" Then
                                NextLink = "https://" & HOME_PAGE & "/mobile/" & LANG_CD & "/index.html" & "?UID=" & UID & "&SES=" & SessID
                            Else
                                NextLink = "https://" & HOME_PAGE & "/mobile/index.html" & "?UID=" & UID & "&SES=" & SessID
                            End If
                            If TYPE_CD = "Survey" Then NextLink = NextLink & "#cert" Else NextLink = NextLink & "#reg"
                        End If
                    End If
                    
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = MSG1 & "<a href=""" & NextLink & """ data-role=""button"" id=""ProceedBtn"" rel=""external"" data-theme=""b"">Haga clic para volver al portal</a>"
                        Case Else
                            MSG1 = MSG1 & "<a href=""" & NextLink & """ data-role=""button"" id=""ProceedBtn"" rel=""external"" data-theme=""b"">Click to return to the portal</a>"
                    End Select
            End Select
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
            mydebuglog.Debug("  .. EXM_STATUS_CD: " & EXM_STATUS_CD)
            mydebuglog.Debug("  .. ClassLink: " & ClassLink)
            mydebuglog.Debug("  .. MSG1: " & MSG1)
            mydebuglog.Debug("  .. TST_NAME: " & TST_NAME)
            mydebuglog.Debug("  .. NextLink: " & NextLink)
            mydebuglog.Debug("  .. PAUSE: " & PAUSE)
            mydebuglog.Debug("  .. DOMAIN: " & DOMAIN)
            mydebuglog.Debug("  .. REG_ID: " & REG_ID)
            mydebuglog.Debug("  .. TST_ID: " & TST_ID)
            mydebuglog.Debug("  .. ErrMsg: " & errmsg)
        End If
        GoTo PrepareExit
        
TestNotFound:
        ErrLvl = "Warning"
        If Debug = "Y" Then mydebuglog.Debug(">>TestNotFound")
        If pTYPE_CD <> "" Then
            errmsg = "Este " & pTYPE_CD & " no se encontr&oacute;"
        Else
            errmsg = "The exam record was not found"
        End If
        GoTo PrepareExit

Incomplete:
        If Debug = "Y" Then mydebuglog.Debug(">>Incomplete")
        ErrLvl = "Warning"
        Select Case LANG_CD
            Case "ESN"
                errmsg = "No hemos podido procesar su salida de clase. Por favor cont&aacute;ctenos para asistencia."
            Case Else
                errmsg = "We were unable to process your class exit.  Please contact us for assistance."
        End Select
        GoTo PrepareExit
   
TimedOut:
        If Debug = "Y" Then mydebuglog.Debug(">>TimedOut")
        ErrLvl = "Warning"
        Select Case LANG_CD
            Case "ESN"
                errmsg = "Su sesi&oacute;n del portal web ha caducado. Por favor inicie sesi&oacute;n para continuar."
            Case Else
                errmsg = "Your web portal session timed out.  Please login to continue."
        End Select
        GoTo PrepareExit
      
NotFound:
        If Debug = "Y" Then mydebuglog.Debug(">>NotFound")
        ErrLvl = "Warning"
        Select Case LANG_CD
            Case "ESN"
                errmsg = pTYPE_CD & " error no encontrado. Por favor cont&aacute;ctenos para asistencia."
            Case Else
                errmsg = pTYPE_CD & " not found error.  Please contact us for assistance."
        End Select
        GoTo PrepareExit
   
DBError:
        If Debug = "Y" Then mydebuglog.Debug(">>DBError")
        Select Case LANG_CD
            Case "ESN"
                errmsg = "El sistema puede no estar disponible ahora. Por favor, int&eacute;ntelo de nuevo m&aacute;s tarde"
            Case Else
                errmsg = "The system may be unavailable now.  Please try again later"
        End Select
        GoTo PrepareExit
   
AccessError:
        If Debug = "Y" Then mydebuglog.Debug(">>AccessError")
        ErrLvl = "Warning"
        Select Case LANG_CD
            Case "ESN"
                errmsg = "No tienes acceso a este examen o encuesta. Por favor cont&aacute;ctenos para asistencia."
            Case Else
                errmsg = "You do not have access to this exam or survey.  Please contact us for assistance."
        End Select
        GoTo PrepareExit
   
DataError:
        If Debug = "Y" Then mydebuglog.Debug(">>DataError")
        ErrLvl = "Warning"
        Select Case LANG_CD
            Case "ESN"
                errmsg = "No podemos abrir su clase debido a un problema con su registro. Por favor cont&aacute;ctenos para asistencia."
            Case Else
                errmsg = "We are unable to open your class due to a problem with your registration.  Please contact us for assistance."
        End Select
   
PrepareExit:
        ' Locate specific variables
        If (DOMAIN = "" Or UID = "") And REG_ID <> "" Then
            SqlS = "SELECT C.X_REGISTRATION_NUM, R.DOMAIN " & _
            "FROM siebeldb.dbo.S_CONTACT C " & _
            "LEFT OUTER JOIN siebeldb.dbo.CX_SESS_REG R ON R.CONTACT_ID=C.ROW_ID " & _
            "WHERE R.ROW_ID='" & REG_ID & "'"
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Get User: " & vbCrLf & "  " & SqlS)
            Try
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        UID = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                        DOMAIN = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                    End While
                End If
            Catch ex As Exception
            End Try
            dr.Close()
        End If
        If DOMAIN = "" Then DOMAIN = "TIPS"
        
        ' Calculate redirect
        If HOME_PAGE = "certegrity.com" And CURRENT_PAGE <> "" Then
            NextLink = CURRENT_PAGE
            If RefreshID <> "" Then
                NextLink = NextLink & "&RFR=" & RefreshID
            End If
        Else
            If CURRENT_PAGE <> "" Then
                NextLink = CURRENT_PAGE
                If InStr(NextLink, "UID=") = 0 Then NextLink = NextLink & "?UID=" & UID
                If InStr(NextLink, "SES=") = 0 Then NextLink = NextLink & "&SES=" & SessID
                If InStr(NextLink, "PP=") = 0 Then NextLink = NextLink & "&PP=" & DOMAIN
                If RefreshID <> "" Then NextLink = NextLink & "&RFR=" & RefreshID
                If TYPE_CD = "Survey" Then
                    If InStr(NextLink, "#cert") = 0 Then NextLink = NextLink & "#cert"
                Else
                    If InStr(NextLink, "#reg") = 0 Then NextLink = NextLink & "#reg"
                End If
            Else
                If LANG_CD <> "ENU" Then
                    NextLink = "https://" & HOME_PAGE & "/mobile/" & LANG_CD & "/index.html" & "?UID=" & UID & "&SES=" & SessID
                Else
                    NextLink = "https://" & HOME_PAGE & "/mobile/index.html" & "?UID=" & UID & "&SES=" & SessID
                End If
                If TYPE_CD = "Survey" Then NextLink = NextLink & "#cert" Else NextLink = NextLink & "#reg"
            End If
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
        outdata = ""
        outdata = outdata & """EXM_STATUS_CD"":""" & EXM_STATUS_CD & ""","
        outdata = outdata & """ClassLink"":""" & EscapeJSON(ClassLink) & ""","
        outdata = outdata & """MSG1"":""" & EscapeJSON(MSG1) & ""","
        outdata = outdata & """TST_NAME"":""" & EscapeJSON(TST_NAME) & ""","
        outdata = outdata & """NextLink"":""" & EscapeJSON(NextLink) & ""","
        outdata = outdata & """PAUSE"":""" & PAUSE & ""","
        outdata = outdata & """DOMAIN"":""" & EscapeJSON(DOMAIN) & ""","
        outdata = outdata & """TST_ID"":""" & EscapeJSON(TST_ID) & ""","
        outdata = outdata & """REG_ID"":""" & EscapeJSON(REG_ID) & ""","
        outdata = outdata & """ErrMsg"":""" & errmsg & """ "
        outdata = callback & "({""ResultSet"": {" & outdata & "} })"
        
        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("WsLeaveAssessment.ashx : " & ErrLvl & ": " & Trim(errmsg))
        myeventlog.Info("WsLeaveAssessment.ashx : UID: " & UID & ", Sess Id: " & SessID & ", Exam Id: " & CRSE_TSTRUN_ID & " - NextLink: " & NextLink)
        If Debug = "Y" Or (Logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("  outdata: " & outdata & vbCrLf)
                mydebuglog.Debug("Results:  UID: " & UID & ", Sess Id: " & SessID & ", Exam Id: " & CRSE_TSTRUN_ID & " - NextLink: " & NextLink)
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
                LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, "WSLEAVEASSESSMENT", LogStartTime, VersionNum, Debug)
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
    Public Function EncodeUID(ByVal UID As String) As String
        Dim btUID() As Byte
        Dim encText As New System.Text.UTF8Encoding()
        btUID = encText.GetBytes(UID)
        EncodeUID = ReverseString(ToBase64(btUID))
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
    ' XML DOCUMENT MANAGEMENT
    Private Sub AddXMLChild(ByVal xmldoc As XmlDocument, ByVal root As XmlElement, _
        ByVal childname As String, ByVal childvalue As String)
        Dim resultsItem As System.Xml.XmlElement

        resultsItem = xmldoc.CreateElement(childname)
        resultsItem.InnerText = childvalue
        root.AppendChild(resultsItem)
    End Sub

    Private Sub CreateXMLChild(ByVal xmldoc As XmlDocument, ByVal root As XmlElement, _
        ByVal childname As String, ByVal childvalue As String)
        Dim resultsItem As System.Xml.XmlElement

        resultsItem = xmldoc.CreateElement(childname)
        resultsItem.InnerText = childvalue
    End Sub

    Private Sub AddXMLAttribute(ByVal xmldoc As XmlDocument, _
        ByVal xmlnode As XmlElement, ByVal attribute As String, _
        ByVal attributevalue As String)
        ' Used to add an attribute to a specified node

        Dim newAtt As XmlAttribute

        newAtt = xmldoc.CreateAttribute(attribute)
        newAtt.Value = attributevalue
        xmlnode.Attributes.Append(newAtt)
    End Sub
    
    Private Function GetNodeValue(ByVal sNodeName As String, ByVal oParentNode As XmlNode) As String
        ' Generic function to return the value of a node in an XML document
        Dim oNode As XmlNode = oParentNode.SelectSingleNode(".//" + sNodeName)
        If oNode Is Nothing Then
            Return String.Empty
        Else
            Return oNode.InnerText
        End If
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