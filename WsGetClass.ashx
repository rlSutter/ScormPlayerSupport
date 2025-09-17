<%@ WebHandler Language="VB" Class="WsGetClass" %>

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

Public Class WsGetClass : Implements IHttpHandler
    
    Enum enumObjectType
        StrType = 0
        IntType = 1
        DblType = 2
        DteType = 3
    End Enum
    
    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        
        ' This web service is invoked from the OpenClass.html page, validates a user's access to a class, and generates information necessary for 
        ' launching a class. If applicable, this agent also returns any related course documents, as well as KBA questions to be asked. 
        ' This service approximates the functionality provided by PsWsOpenClass, which in turn was designed to replace CMOpenSClass
        
        ' Parameter Declarations
        Dim REG_ID, UID, SessID, CURRENT_PAGE, HOME_PAGE, LANG_CD, callback, myprotocol As String
        Dim Debug As String
        
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
        mydebuglog = log4net.LogManager.GetLogger("GetClassDebugLog")
        Dim logfile, tempdebug As String
        Dim Logging As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "103"
        
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
        Dim errmsg, EOL, ErrLvl, MSG1, MSG2, LAST_INST As String
        Dim KBA_FLAG, OldScorm, AutoStart, AnotherWindow, Olark, HCIPlayer As Boolean
        Dim LOGGED_IN, CONTACT_ID, SUB_ID, CONTACT_OU_ID, DOMAIN, LogoutUser As String
        Dim REG_STATUS_CD, CRSE_CONTENT_URL, TRAIN_TYPE, SCORM, CRSE_ID, REGISTRANT_ID, CRSE_TST_ID, EXAM_ENGINE As String
        Dim COURSE, RESOLUTION, JURIS_ID, KBA_REQD, temp, REF_CON_ID, TEST_FLG, USER_NAME As String
        Dim EMAIL_ADDR, X_FORMAT, ltemp, FST_NAME, LAST_NAME, PAUSE, from_db As String
        Dim RES_X, RES_Y As String
        Dim KBA_QUESTIONS As Integer
        Dim ClassLink, PCKG_ID As String
        Dim ENTER_FLG, EXIT_FLG, TIME_ELAPSED As String
        Dim ACTIVITY_COUNT, ntemp As Integer
        Dim SCORM_ACTIVITY, ReturnLink, mailpath As String
        Dim KBA_COUNT, TO_ASK, NUM_ANSRD, TotalDocs As Integer
        Dim JURIS, JURIS_LVL, KBA_NOTICE, ConfirmEmail, SCORM_FLG As String
        Dim Q_ID(50) As String
        Dim Q_TEXT(50) As String
        Dim DocId, DocName, DocDesc, CPage, attachlink, AccessToken As String
        Dim eCreateClass, eAcceptClass, PRE_CRSE_DOC_ID, POST_CRSE_DOC_ID As String
        Dim lockoutcount As Integer
 
        ' ============================================
        ' Variable setup
        Debug = "N"
        lockoutcount = 0
        errmsg = ""
        LOGGED_IN = "N"
        CONTACT_ID = ""
        SUB_ID = ""
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
        RES_X = ""
        RES_Y = ""
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
        JURIS_ID = ""
        KBA_COUNT = 0
        KBA_QUESTIONS = 0
        TO_ASK = 0
        NUM_ANSRD = 0
        JURIS_LVL = ""
        JURIS = ""
        ConfirmEmail = ""
        SCORM_FLG = ""
        KBA_NOTICE = ""
        CPage = ""
        attachlink = ""
        AccessToken = ""
        EOL = Chr(13)
        KBA_FLAG = False
        OldScorm = False
        AutoStart = True
        Olark = False
        HCIPlayer = False
        AnotherWindow = False
        ErrLvl = "Error"
        PAUSE = "0"
        LAST_INST = ""
        eCreateClass = ""
        eAcceptClass = ""
        UID = ""
        SessID = ""
        REG_ID = ""
        mailpath = ""
        CRSE_ID = ""
        SCORM = ""
        jdoc = ""
        RESOLUTION = ""
        USER_NAME = ""
        ClassLink = ""
        from_db = ""
        PRE_CRSE_DOC_ID = ""
        POST_CRSE_DOC_ID = ""
        
        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server=DB_SERVER;uid=DB_USER;pwd=DB_PASSWORD;database=siebeldb"
            tempdebug = System.Configuration.ConfigurationManager.AppSettings.Get("GetClass_debug")
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
            logfile = "C:\Logs\WsGetClass.log"
            Try
                log4net.GlobalContext.Properties("GetClassLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                GoTo CloseOut2
            End Try
        End If
        
        ' ============================================
        ' Get parameters    
        '   REG_ID          - The ROW_ID of the CX_SESS_REG record
        '   UID             - The user's S_CONTACT.X_REGISTRATION_NUM
        '   SessID          - The user's session id
        '   CURRENT_PAGE    - Optional. The page from which this service was called
        '   LANG_CD         - Language code.  Default "ENU"
        '   HOME_PAGE       - Optional. The root domain that invoked this agent/service
        '   callback        - The name of the Javascript callback function in which to wrap the resulting JSON
        '   myprotocol      - Whether the service was called via http or https
        '   Debug           - "Y", "N" or "T"      

        If Not context.Request.QueryString("RID") Is Nothing Then
            REG_ID = context.Request.QueryString("RID")
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

        If Not context.Request.QueryString("CUR") Is Nothing Then
            CURRENT_PAGE = context.Request.QueryString("CUR")
        End If
 
        If Not context.Request.QueryString("LANG") Is Nothing Then
            LANG_CD = UCase(context.Request.QueryString("LANG"))
        End If
        
        If Not context.Request.QueryString("callback") Is Nothing Then
            callback = context.Request.QueryString("callback")
        End If

        If Not context.Request.QueryString("PROT") Is Nothing Then
            myprotocol = LCase(context.Request.QueryString("PROT"))
        End If

        ' Validate parameters
        If InStr(1, CURRENT_PAGE, "http:") = 0 And InStr(1, CURRENT_PAGE, "https:") = 0 Then
            CURRENT_PAGE = "https:" & CURRENT_PAGE
        End If
        If InStr(1, PrevLink, "?UID") = 0 Then PrevLink = PrevLink & "?UID=" & UID & "&SES=" & SessID
        PrevLink = Replace(PrevLink, "#reg", "")
        If LANG_CD <> "ESN" And LANG_CD <> "ENU" Then LANG_CD = "ENU"
        If callback = "" Then callback = "?"
        If Left(HOME_PAGE, 4) <> "web." And Left(HOME_PAGE, 4) <> "www." and HOME_PAGE<>"certegrity.com" Then
            If InStr(1, PrevLink, "web.") > 0 Then HOME_PAGE = "web." & HOME_PAGE Else HOME_PAGE = "www." & HOME_PAGE
        End If
        If myprotocol = "" Then myprotocol = "http:"

        If Debug = "Y" Then
            mydebuglog.Debug("----------------------------------")
            mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
            mydebuglog.Debug("Parameters-")
            mydebuglog.Debug("  lockoutcount: " & Str(lockoutcount))
            mydebuglog.Debug("  REG_ID: " & REG_ID)
            mydebuglog.Debug("  UID: " & UID)
            mydebuglog.Debug("  cookieid: " & cookieid)
            mydebuglog.Debug("  SessID: " & SessID)
            mydebuglog.Debug("  CURRENT_PAGE : " & CURRENT_PAGE)
            mydebuglog.Debug("  HOME_PAGE: " & HOME_PAGE)
            mydebuglog.Debug("  myprotocol: " & myprotocol)
            mydebuglog.Debug("  PrevLink: " & PrevLink)
            mydebuglog.Debug("  BROWSER: " & BROWSER)
            mydebuglog.Debug("  LANG_CD: " & LANG_CD)
            mydebuglog.Debug("  callback: " & callback & vbCrLf)
        End If
        
        If REG_ID = "" Or UID = "" Or SessID = "" Then            
            GoTo DataError
        End If
        If cookieid <> UID Then
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
        ' Process
        If Not cmd Is Nothing Then
            
            ' ================================================   
            ' GET USER PROFILE
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
                        DOMAIN = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                    End While
                End If
            Catch ex As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Unable to locate credentials. Error: " & vbCrLf & ex.ToString & vbCrLf)
                LogoutUser = "Y"
                GoTo AccessError
            End Try
            dr.Close()
            If Debug = "Y" Then
                mydebuglog.Debug("  .. LOGGED_IN: " & LOGGED_IN)
                mydebuglog.Debug("  .. CONTACT_ID: " & CONTACT_ID)
                mydebuglog.Debug("  .. CONTACT_OU_ID: " & CONTACT_OU_ID)
                mydebuglog.Debug("  .. DOMAIN: " & DOMAIN)
            End If
            If LOGGED_IN = "Y" Then
                LogoutUser = "N"
            Else
                LogoutUser = "Y"
                GoTo AccessError
            End If

            ' ================================================
            ' SET SCREEN DISPLAY DEFAULTS
            MSG1 = ""
            MSG2 = ""

            ' ================================================
            ' CHECK REGISTRATION AND CLASS
            SqlS = "SELECT TOP 1 R.STATUS_CD, CR.X_CRSE_CONTENT_URL, R.ROW_ID, T.TRAIN_TYPE, " & _
                "CR.X_SCORM_FLG, CR.ROW_ID, R.CONTACT_ID, E.ROW_ID, E.X_ENGINE, CR.NAME, CR.X_RESOLUTION, " & _
                "R.JURIS_ID, JC.KBA_REQD, JC.KBA_QUESTIONS, R.REF_CON_ID, R.TEST_FLG, " & _
                "C.FST_NAME+' '+C.LAST_NAME, C.EMAIL_ADDR, CR.X_FORMAT, CR.X_LANG_CD, C.FST_NAME, C.LAST_NAME, " & _
                "SA.PRE_CRSE_DOC_ID, SA.POST_CRSE_DOC_ID " & _
                "FROM siebeldb.dbo.CX_SESS_REG R " & _
                "LEFT OUTER JOIN siebeldb.dbo.CX_TRAIN_OFFR T ON T.ROW_ID=R.TRAIN_OFFR_ID " & _
                "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=R.CONTACT_ID " & _
                "LEFT OUTER JOIN siebeldb.dbo.S_CRSE CR ON CR.ROW_ID=R.CRSE_ID " & _
                "LEFT OUTER JOIN siebeldb.dbo.CX_CRSE_SCORM_ATTR SA ON SA.CRSE_ID=CR.ROW_ID " & _
                "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TST E ON E.CRSE_ID=CR.ROW_ID " & _
                "LEFT OUTER JOIN siebeldb.dbo.CX_JURIS_CRSE JC ON JC.JURIS_ID=R.JURIS_ID AND JC.CRSE_ID=T.CRSE_ID " & _
                "WHERE R.ROW_ID='" & REG_ID & "' AND C.X_REGISTRATION_NUM='" & UID & "' AND (E.X_SURVEY_FLG<>'Y' OR E.X_SURVEY_FLG IS NULL) AND E.STATUS_CD='Active'"
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  CHECK REGISTRATION AND CLASS: " & vbCrLf & "  " & SqlS)
            Try
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        REG_STATUS_CD = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                        CRSE_CONTENT_URL = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                        REG_ID = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                        TRAIN_TYPE = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                        SCORM = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                        CRSE_ID = Trim(CheckDBNull(dr(5), enumObjectType.StrType))
                        REGISTRANT_ID = Trim(CheckDBNull(dr(6), enumObjectType.StrType))
                        CRSE_TST_ID = Trim(CheckDBNull(dr(7), enumObjectType.StrType))
                        EXAM_ENGINE = Trim(CheckDBNull(dr(8), enumObjectType.StrType))
                        COURSE = Trim(CheckDBNull(dr(9), enumObjectType.StrType))
                        RESOLUTION = Trim(CheckDBNull(dr(10), enumObjectType.StrType))
                        JURIS_ID = Trim(CheckDBNull(dr(11), enumObjectType.StrType))
                        KBA_REQD = Trim(CheckDBNull(dr(12), enumObjectType.StrType))
                        temp = Trim(CheckDBNull(dr(13), enumObjectType.StrType))
                        If temp = "" Then KBA_QUESTIONS = 0 Else KBA_QUESTIONS = Val(temp)
                        REF_CON_ID = Trim(CheckDBNull(dr(14), enumObjectType.StrType))
                        TEST_FLG = Trim(CheckDBNull(dr(15), enumObjectType.StrType))
                        USER_NAME = Trim(CheckDBNull(dr(16), enumObjectType.StrType))
                        EMAIL_ADDR = Trim(CheckDBNull(dr(17), enumObjectType.StrType))
                        X_FORMAT = UCase(Trim(CheckDBNull(dr(18), enumObjectType.StrType)))
                        ltemp = UCase(Trim(CheckDBNull(dr(19), enumObjectType.StrType)))
                        If LANG_CD = "" And LANG_CD <> "" Then LANG_CD = ltemp
                        If LANG_CD = "" Then LANG_CD = "ENU"
                        FST_NAME = Trim(CheckDBNull(dr(20), enumObjectType.StrType))
                        LAST_NAME = Trim(CheckDBNull(dr(21), enumObjectType.StrType))
                        PRE_CRSE_DOC_ID = Trim(Str(CheckDBNull(dr(22), enumObjectType.IntType)))
                        If PRE_CRSE_DOC_ID = "0" Then PRE_CRSE_DOC_ID = ""
                        POST_CRSE_DOC_ID = Trim(Str(CheckDBNull(dr(23), enumObjectType.IntType)))
                        If POST_CRSE_DOC_ID = "0" Then POST_CRSE_DOC_ID = ""
                        If SCORM = "Y" Then
                            RES_X = Trim(Str(Val(RES_X) - 20))
                            RES_Y = Trim(Str(Val(RES_Y) - 20))
                        End If
                        If X_FORMAT = "HTML5" Then
                            RES_Y = Trim(Str(Val(RES_Y) * 0.85))
                        End If
                        If InStr(1, UCase(CRSE_CONTENT_URL), "HCIPLAYER") > 0 Then HCIPlayer = True
                    End While
                Else
                    GoTo AccessError
                End If
                dr.Close()
            Catch ex As Exception
                GoTo AccessError
            End Try
            
            If Debug = "Y" Then
                mydebuglog.Debug("  .. REGISTRANT_ID: " & REGISTRANT_ID)
                mydebuglog.Debug("  .. CRSE_TST_ID: " & CRSE_TST_ID)
                mydebuglog.Debug("  .. EXAM_ENGINE: " & EXAM_ENGINE)
                mydebuglog.Debug("  .. CRSE_CONTENT_URL: " & CRSE_CONTENT_URL)
                mydebuglog.Debug("  .. HCIPlayer: " & HCIPlayer)
                mydebuglog.Debug("  .. TEST_FLG: " & TEST_FLG)
                mydebuglog.Debug("  .. KBA_REQD: " & KBA_REQD)
                mydebuglog.Debug("  .. LANG_CD: " & LANG_CD)
                mydebuglog.Debug("  .. COURSE: " & COURSE)
                mydebuglog.Debug("  .. NAME: " & FST_NAME & " " & LAST_NAME)
                mydebuglog.Debug("  .. EMAIL_ADDR: " & EMAIL_ADDR)
                mydebuglog.Debug("  .. X_FORMAT: " & X_FORMAT)
                mydebuglog.Debug("  .. PRE_CRSE_DOC_ID: " & PRE_CRSE_DOC_ID)
                mydebuglog.Debug("  .. POST_CRSE_DOC_ID: " & POST_CRSE_DOC_ID)
            End If
            
            ' ================================================
            ' CHECK REGISTRANT
            If REGISTRANT_ID <> CONTACT_ID Then
                MSG1 = "<table width=""100%"" height=""100%""><tr valign=""Middle""><td class=""BigHeader""><font color=""Gray"">You are not the registered student for this course and may not take it.  If you believe that this is in error, please contact our <a href=""JavaScript:openNewWindow('" & mailpath & "/message?OpenForm&+param1=Technical+Supportparam2=Course+Access+Question',525,375)""><b>Technical Support</b></a> department</FONT></td></tr></table>"
                ClassLink = ""
                GoTo CloseOut
            End If
            
            ' ================================================
            ' DOUBLE CHECK PLAYER
            If CRSE_ID <> "" And REG_ID <> "" Then
                SqlS = "SELECT from_db FROM elearning.dbo.Elearning_Player_Data WHERE reg_id='" & REG_ID & "' AND crse_id='" & CRSE_ID & "' ORDER BY update_dt DESC"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  DOUBLE CHECK PLAYER: " & vbCrLf & "  " & SqlS)
                cmd.CommandText = SqlS
                Try
                    from_db = Trim(CheckDBNull(cmd.ExecuteScalar(), enumObjectType.StrType))
                Catch ex As Exception
                End Try
                
                If EXAM_ENGINE = "" Then
                    ' If no engine found, then set based on from_db
                    If from_db = "elearning" Or from_db = "" Then
                        EXAM_ENGINE = "HCIPLAYER"
                        HCIPlayer = True
                    End If
                    If from_db = "hciscorm" Then
                        EXAM_ENGINE = "SCORM"
                        HCIPlayer = False
                    End If
                Else
                    ' Engine is known, then test based on from_db                    
                    If from_db = "elearning" Then
                        HCIPlayer = True
                    Else
                        If from_db <> "hciscorm" Then HCIPlayer = True
                        If from_db = "hciscorm" Then HCIPlayer = False
                    End If
                    If EXAM_ENGINE = "SCORM" And HCIPlayer Then
                        HCIPlayer = False
                    End If
                End If
                If Debug = "Y" Then
                    mydebuglog.Debug("  .. from_db: " & from_db)
                    mydebuglog.Debug("  .. HCIPlayer: " & Str(HCIPlayer))
                End If
            End If
            
            ' ================================================
            ' CHECK ACCESS STATUS 
            If SCORM = "Y" Then
                SqlS = "SELECT package_id " & _
                    "FROM elearning.dbo.Elearning_Package_Data " & _
                    "WHERE crse_id='" & CRSE_ID & "'"
                If from_db <> "" Then SqlS = SqlS & " AND from_db='" & from_db & "'"
                SqlS = SqlS & " ORDER BY version_id DESC"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Get SCORM package id: " & vbCrLf & "  " & SqlS)
                cmd.CommandText = SqlS
                Try
                    PCKG_ID = CheckDBNull(cmd.ExecuteScalar(), enumObjectType.StrType)
                Catch ex As Exception
                    GoTo AccessError
                End Try
                If Debug = "Y" Then mydebuglog.Debug("  .. PCKG_ID : " & PCKG_ID)
                
                ' Get access semaphor count
                Dim accessfound As Integer
                SqlS = "SELECT COUNT(*) FROM siebeldb.dbo.CX_TRAIN_OFFR_ACCESS WHERE REG_ID='" & REG_ID & "'"
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
                    SqlS = "UPDATE siebeldb.dbo.CX_SESS_REG SET STATUS_CD='On-Hold' WHERE ROW_ID='" & REG_ID & "'"
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Locking student out because access count exceeded: " & vbCrLf & "  " & SqlS)
                    temp = ExecQuery("Update", "CX_SESS_REG", cmd, SqlS, mydebuglog, "N")
                    GoTo AccessCountError
                End If
                
                ' Get access semaphor
                ENTER_FLG = ""
                EXIT_FLG = ""
                TIME_ELAPSED = ""
                SCORM_ACTIVITY = 0
                SqlS = "SELECT TOP 1 ENTER_FLG, EXIT_FLG, DATEDIFF(MINUTE,CREATED,GETDATE()) AS TIME_ELAPSED FROM siebeldb.dbo.CX_TRAIN_OFFR_ACCESS " & _
                       "WHERE REG_ID='" & REG_ID & "' ORDER BY CREATED DESC "
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Get access semaphor: " & vbCrLf & "  " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            ENTER_FLG = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            EXIT_FLG = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            TIME_ELAPSED = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                            If TIME_ELAPSED = "" Then TIME_ELAPSED = "0"
                        Catch ex As Exception
                        End Try
                    End While
                Else
                    GoTo DataError
                End If
                dr.Close()
                If Debug = "Y" Then
                    mydebuglog.Debug("  .. ENTER_FLG : " & ENTER_FLG)
                    mydebuglog.Debug("  .. EXIT_FLG : " & EXIT_FLG)
                    mydebuglog.Debug("  .. TIME_ELAPSED : " & TIME_ELAPSED)
                End If
                
                ' If the enter state is already set, then do not restart the class - it is open in another browser window
                If ENTER_FLG = "Y" Then
                    SqlS = "SELECT COUNT(*) " & _
                    "FROM elearning.dbo.Elearning_Player_Data " & _
                    "WHERE reg_id='" & REG_ID & "' AND crse_id='" & CRSE_ID & "'"
                    If from_db <> "" Then SqlS = SqlS & " AND from_db='" & from_db & "'"
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Get course state: " & vbCrLf & "  " & SqlS)
                    cmd.CommandText = SqlS
                    Try
                        ACTIVITY_COUNT = CheckDBNull(cmd.ExecuteScalar(), enumObjectType.IntType)
                        If Not IsNumeric(ACTIVITY_COUNT) Then ACTIVITY_COUNT = 0
                    Catch ex As Exception
                    End Try
                    If Debug = "Y" Then
                        mydebuglog.Debug("  .. ACTIVITY_COUNT : " & Str(ACTIVITY_COUNT))
                    End If
                End If
                
                ' If activity found check to see how the "active" and "suspended" flags were set - this implies a browser crash
                If ACTIVITY_COUNT > 0 Then
                    SqlS = "SELECT RTRIM(CAST(active AS CHAR))+RTRIM(CAST(suspended AS CHAR)) " & _
                        "FROM elearning.dbo.Elearning_Player_Data " & _
                        "WHERE reg_id='" & Trim(REG_ID) & "'"
                    If from_db <> "" Then SqlS = SqlS & " AND from_db='" & from_db & "'"
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  SCORM CHECK FLAGS QUERY: " & vbCrLf & "  " & SqlS)
                    cmd.CommandText = SqlS
                    Try
                        SCORM_ACTIVITY = Trim(CheckDBNull(cmd.ExecuteScalar(), enumObjectType.StrType))
                    Catch ex As Exception
                    End Try
                    If Debug = "Y" Then
                        mydebuglog.Debug("  .. SCORM_ACTIVITY : " & Str(SCORM_ACTIVITY))
                    End If
                    If SCORM_ACTIVITY = "10" And TIME_ELAPSED < 10 Then
                        ClassLink = ""
                        MSG1 = "You may have this class already open in another browser window or on another device.  This error sometimes occurs if your prior attempt to take this class was disrupted due to a sudden loss of network access or power, or due to a browser failure.  If this is the case, if you wait 10 minutes before trying this again your access will be reset and you should be able to return to your class.  Our apologies for this inconvenience. <br><br>" & _
                        "If you believe that this is in error, please contact our <a href=""JavaScript:openNewWindow('" & mailpath & "/message?OpenForm&param1=Technical+Support&param2=Course+Access+Question&DOM=" & DOMAIN & "',525,375)""><b>Technical Support</b></a> department<br><br>" & _
                        "<a href=[RETURN] data-role=""button"" rel=""external"" data-theme=""b"">Return to the previous screen</a>"
                        If CURRENT_PAGE <> "" Then
                            ReturnLink = CURRENT_PAGE
                            If InStr(ReturnLink, "UID=") = 0 Then ReturnLink = ReturnLink & "?UID=" & UID$
                            If InStr(ReturnLink, "SES=") = 0 Then ReturnLink = ReturnLink & "&SES=" & SessID$
                            If InStr(ReturnLink, "#reg") = 0 Then ReturnLink = ReturnLink & "#reg"
                        Else
                            ReturnLink = "http://" & HOME_PAGE & "/mobile/index.html" & "?UID=" & UID$ & "&SES=" & SessID$ & "#reg"
                        End If
                        MSG1 = Replace(MSG1, "[RETURN]", ReturnLink)
                        AnotherWindow = True
                    End If

                    ' Zap flag - it was set spuriously
                    'SqlS = "DELETE FROM siebeldb.dbo.CX_TRAIN_OFFR_ACCESS " & _
                    '     "WHERE REG_ID='" & REG_ID & "'"
                    'temp = ExecQuery("Update", "CX_TRAIN_OFFR_ACCESS", cmd, SqlS, mydebuglog, Debug)
                End If
            End If
            
            ' ================================================
            ' IF THE CLASS IS NOT COMPLETED        
            If (REG_STATUS_CD = "Accepted" Or REG_STATUS_CD = "In Progress" Or REG_STATUS_CD = "Retake") And CRSE_CONTENT_URL <> "" Then
                
                ' ----------
                ' FIX JURISDICTION IF BLANK
                If JURIS_ID = "" Then
                    SqlS = "SELECT (SELECT CASE WHEN A.X_JURIS_ID IS NOT NULL THEN A.X_JURIS_ID ELSE " & _
                        "(SELECT CASE WHEN BA.X_JURIS_ID IS NOT NULL THEN BA.X_JURIS_ID ELSE " & _
                        "(SELECT CASE WHEN P.X_JURIS_ID IS NOT NULL THEN P.X_JURIS_ID ELSE " & _
                        "(SELECT CASE WHEN CP.X_JURIS_ID IS NOT NULL THEN CP.X_JURIS_ID ELSE '' END) END) END) END) AS JURIS_ID  " & _
                        "FROM siebeldb.dbo.CX_SESS_REG R " & _
                        "LEFT OUTER JOIN siebeldb.dbo.S_ORG_EXT O ON O.ROW_ID=R.OU_ID " & _
                        "LEFT OUTER JOIN siebeldb.dbo.S_ADDR_ORG A ON A.ROW_ID=R.ADDR_ID " & _
                        "LEFT OUTER JOIN siebeldb.dbo.S_ADDR_ORG BA ON BA.ROW_ID=O.PR_ADDR_ID " & _
                        "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=R.CONTACT_ID " & _
                        "LEFT OUTER JOIN siebeldb.dbo.S_ADDR_PER P ON P.ROW_ID=R.PER_ADDR_ID " & _
                        "LEFT OUTER JOIN siebeldb.dbo.S_ADDR_PER CP ON CP.ROW_ID=C.PR_PER_ADDR_ID " & _
                        "WHERE R.ROW_ID='" & REG_ID & "'"
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  GET NEW JURISDICTION ID: " & vbCrLf & "  " & SqlS)
                    cmd.CommandText = SqlS
                    Try
                        JURIS_ID = Trim(CheckDBNull(cmd.ExecuteScalar(), enumObjectType.StrType))
                        If JURIS_ID = "" Then GoTo DataError
                    Catch ex As Exception
                    End Try
                    mydebuglog.Debug("  .. JURIS_ID : " & JURIS_ID)
                    
                    ' FIX RECORD
                    SqlS = "UPDATE siebeldb.dbo.CX_SESS_REG " & _
                        "SET JURIS_ID='" & JURIS_ID & "' WHERE ROW_ID='" & REG_ID & "'"
                    temp = ExecQuery("Update", "CX_TRAIN_OFFR_ACCESS", cmd, SqlS, mydebuglog, Debug)
                    
                End If
                
                ' ----------
                ' CHECK KBA STATUS
                If KBA_REQD = "Y" Then
                    SqlS = "SELECT COUNT(*) " & _
                        "FROM elearning.dbo.KBA_QUES Q " & _
                        "LEFT OUTER JOIN elearning.dbo.KBA_JURIS J ON J.QUES_ID=Q.ROW_ID " & _
                        "LEFT OUTER JOIN elearning.dbo.KBA_CRSE C ON C.QUES_ID=Q.ROW_ID " & _
                        "WHERE Q.ROW_ID IS NOT NULL AND C.CRSE_ID='" & CRSE_ID & "' AND J.JURIS_ID='" & JURIS_ID & "' AND Q.ACTIVE_FLG='Y' AND Q.LANG_CD='" & LANG_CD & "'"
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  COUNT NUMBER OF KBA QUESTIONS FOR A JURISDICTION: " & vbCrLf & "  " & SqlS)
                    cmd.CommandText = SqlS
                    Try
                        KBA_COUNT = CheckDBNull(cmd.ExecuteScalar(), enumObjectType.IntType)
                    Catch ex As Exception
                    End Try
                    mydebuglog.Debug("  .. KBA_COUNT : " & Str(KBA_COUNT))
                    
                    ' CHECK TO SEE HOW MANY QUESTIONS HAVE ACTUALLY BEEN ANSWERED IF NOT PROVIDED
                    '   NUM_ANSRD = number the student actually answered already
                    SqlS = "SELECT COUNT(*) " & _
                        "FROM elearning.dbo.KBA_ANSR A " & _
                        "WHERE A.REG_ID='" & REG_ID & "'"
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  CHECK TO SEE HOW MANY QUESTIONS ANSWERED: " & vbCrLf & "  " & SqlS)
                    cmd.CommandText = SqlS
                    Try
                        NUM_ANSRD = Trim(CheckDBNull(cmd.ExecuteScalar(), enumObjectType.IntType))
                    Catch ex As Exception
                    End Try
                    mydebuglog.Debug("  .. NUM_ANSRD : " & Str(NUM_ANSRD))
                
                    ' Determine if we need to ask KBA questions
                    '      KBA_QUESTIONS = Number of questions required by the jurisdiction - "0" means all of the questions available
                    '      KBA_COUNT = Number of questions available for a jurisdiction
                    '      NUM_ANSRD = Number of questions answered by the student
                    '   Redirect to the KBA question screen if the number asked is less than the number required, and the number required is greater than zero
                    If KBA_QUESTIONS = 0 Then KBA_QUESTIONS = KBA_COUNT ' Set the number of questions to the jurisdiction count if zero
                    ntemp = 0
                    If NUM_ANSRD < KBA_QUESTIONS Then
                        KBA_FLAG = True
                        ' GET THE LIST OF QUESTIONS AND STORE IN AN ARRAY THAT HAVE NOT BEEN ASKED YET
                        '   TO_ASK = number of actual questions to ask - computed
                        SqlS = "SELECT Q.ROW_ID, Q.QUES_TEXT, JC.KBA_QUESTIONS, JN.NAME, JN.JURIS_LVL, CR.NAME, " & _
                        "CN.FST_NAME, CN.LAST_NAME, CN.EMAIL_ADDR, JC.KBA_NOTICE, DM.CS_EMAIL, CR.X_SCORM_FLG " & _
                        "FROM siebeldb.dbo.CX_SESS_REG S " & _
                        "LEFT OUTER JOIN elearning.dbo.KBA_JURIS J ON J.JURIS_ID=S.JURIS_ID " & _
                        "LEFT OUTER JOIN elearning.dbo.KBA_CRSE C ON C.CRSE_ID=S.CRSE_ID " & _
                        "LEFT OUTER JOIN elearning.dbo.KBA_QUES Q ON Q.ROW_ID=C.QUES_ID AND Q.ROW_ID=J.QUES_ID " & _
                        "LEFT OUTER JOIN siebeldb.dbo.CX_JURIS_CRSE JC ON JC.JURIS_ID=S.JURIS_ID AND JC.CRSE_ID=C.CRSE_ID " & _
                        "LEFT OUTER JOIN siebeldb.dbo.CX_JURISDICTION_X JN ON JN.ROW_ID=S.JURIS_ID " & _
                        "LEFT OUTER JOIN siebeldb.dbo.S_CRSE CR ON CR.ROW_ID=S.CRSE_ID " & _
                        "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT CN ON CN.ROW_ID=S.CONTACT_ID " & _
                        "LEFT OUTER JOIN siebeldb.dbo.CX_SUB_DOMAIN DM ON DM.DOMAIN='" & DOMAIN & "' " & _
                        "WHERE S.ROW_ID='" & REG_ID & "' AND Q.LANG_CD=CN.X_PR_LANG_CD AND Q.ROW_ID IS NOT NULL AND Q.ACTIVE_FLG='Y' AND NOT EXISTS (" & _
                        "SELECT QUES_ID FROM elearning.dbo.KBA_ANSR WHERE QUES_ID=Q.ROW_ID AND REG_ID=S.ROW_ID) " & _
                        "ORDER BY NEWID()"
                        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  LOCATE ALL KBA QUESTIONS: " & vbCrLf & "  " & SqlS)
                        cmd.CommandText = SqlS
                        dr = cmd.ExecuteReader()
                        If Not dr Is Nothing Then
                            While dr.Read()
                                Try
                                    ntemp = ntemp + 1
                                    If ntemp > 100 Then
                                        ReDim Q_ID(ntemp)
                                        ReDim Q_TEXT(ntemp)
                                    End If
                                    Q_ID(ntemp) = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                                    Q_TEXT(ntemp) = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                                    If ntemp = 1 Then
                                        KBA_QUESTIONS = Trim(CheckDBNull(dr(2), enumObjectType.IntType))
                                        JURIS = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                                        JURIS_LVL = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                                        KBA_NOTICE = Trim(CheckDBNull(dr(9), enumObjectType.StrType))
                                        ConfirmEmail = Trim(CheckDBNull(dr(10), enumObjectType.StrType))
                                        SCORM_FLG = Trim(CheckDBNull(dr(11), enumObjectType.StrType))
                                    End If
                                Catch ex As Exception
                                End Try
                            End While
                        Else
                            GoTo DataError
                        End If
                        dr.Close()
                        If Debug = "Y" Then
                            mydebuglog.Debug("  .. ntemp : " & Str(ntemp))
                            mydebuglog.Debug("  .. JURIS : " & JURIS)
                            mydebuglog.Debug("  .. JURIS_LVL : " & JURIS_LVL)
                            mydebuglog.Debug("  .. ACTIVITY_COUNT : " & ACTIVITY_COUNT)
                        End If

                        ' ================================================
                        ' FIND OUT NUMBER OF QUESTIONS THAT SHOULD BE ASKED
                        If KBA_QUESTIONS = 0 Then
                            TO_ASK = ntemp
                        Else
                            TO_ASK = KBA_QUESTIONS - Val(NUM_ANSRD)
                        End If
                        If Debug = "Y" Then mydebuglog.Debug("  .. TO_ASK : " & Str(TO_ASK))
                    End If
      
                End If

                ' GENERATE A NEW ACCESS TOKEN
                AccessToken = LoggingService.GeneratePassword(Debug)

                ' ----------
                ' GET ANY REQUIRED DOCUMENTS
                CPage = "<table cellpadding=""1"" border=""0"" cellspacing=""0"" width=""98%""><tr><td><table border=0 width=""100%"" cellpadding=2 cellspacing=1>"
                SqlS = "SELECT D.row_id, D.name, D.description, DT.name as type, FORMAT(D.created,'M/d/yyyy'), FORMAT(D.last_upd,'M/d/yyyy') " & _
                    "FROM DMS.dbo.Documents D " & _
                    "LEFT OUTER JOIN DMS.dbo.Document_Categories DC ON DC.doc_id=D.row_id and DC.pr_flag='Y' " & _
                    "LEFT OUTER JOIN DMS.dbo.Categories C ON C.row_id=DC.cat_id " & _
                    "LEFT OUTER JOIN DMS.dbo.Document_Associations DA on DA.doc_id=D.row_id " & _
                    "LEFT OUTER JOIN DMS.dbo.Associations A on A.row_id=DA.association_id " & _
                    "LEFT OUTER JOIN DMS.dbo.Document_Types DT on DT.row_id=D.data_type_id " & _
                    "WHERE D.row_id IS NOT NULL AND D.deleted IS NULL AND A.name='Course' AND DA.fkey='" & CRSE_ID & "' AND C.row_id=102"
                If POST_CRSE_DOC_ID <> "" Then
                    SqlS = SqlS & " AND D.row_id<>" & POST_CRSE_DOC_ID
                End If
                If PRE_CRSE_DOC_ID <> "" Then
                    SqlS = SqlS & " AND D.row_id=" & PRE_CRSE_DOC_ID
                End If
                SqlS = SqlS & " AND EXISTS ( " & _
                    "SELECT oda.doc_id FROM DMS.dbo.Document_Associations oda " & _
                    "LEFT OUTER JOIN DMS.dbo.Associations oa ON oa.row_id=oda.association_id " & _
                    "WHERE oa.name='Jurisdiction' AND oda.fkey='" & JURIS_ID & "' AND oda.doc_id=D.row_id) " & _
                    "GROUP BY D.row_id, D.name, D.description, DT.name, C.name, C.row_id, DC.pr_flag, A.name, A.row_id, " & _
                    "DA.pr_flag, DA.fkey, D.created, D.last_upd " & _
                    "ORDER BY D.name, D.last_upd"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Document Query: " & vbCrLf & "  " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            TotalDocs = TotalDocs + 1
                            DocId = Trim(CheckDBNull(dr(0), enumObjectType.StrType))                    ' Documents.row_id
                            DocName = Trim(CheckDBNull(dr(1), enumObjectType.StrType))                  ' Documents.name
                            DocDesc = Trim(CheckDBNull(dr(2), enumObjectType.StrType))                  ' Documents.description
                            If Debug = "Y" Then mydebuglog.Debug("  .. DocId : " & DocId & " - " & DocName)
                            CPage = CPage & "<tr valign=""Top"">" & EOL
                            Select Case LANG_CD
                                Case "ESN"
                                    attachlink = "http://hciscorm.certegrity.com/ls/" & LANG_CD & "/OpenDocument.html?DMI=" & DocId & "&TK=" & AccessToken & "&RG=" & UID & "&SS=" & SessID & "&PROT=" & myprotocol
                                Case Else
                                    attachlink = "http://hciscorm.certegrity.com/ls/OpenDocument.html?DMI=" & DocId & "&TK=" & AccessToken & "&RG=" & UID & "&SS=" & SessID & "&PROT=" & myprotocol
                            End Select
                            If DOMAIN <> "" Then attachlink = attachlink & "&PUB=Y&PP=" & DOMAIN
                            attachlink = "<a href=""JavaScript:openNewWin2('" & attachlink & "',1024,768)"" data-role=""button"" data-icon=""star"" rel=""external"" data-theme=""a"" id=""DocBtn" & Trim(Str(TotalDocs)) & """>"
                            CPage = CPage & "<td valign=""top"" width=""100%"" class=""BigHeader"" align=""center"">" & attachlink & DocName & "</a></font></td></tr>" & EOL
                        Catch ex As Exception
                        End Try
                    End While
                    If CPage <> "" Then CPage = CPage & "</table></td></tr></table>" & EOL
                Else
                    errmsg = errmsg & "The exam record was not found." & vbCrLf
                    GoTo CloseOut
                End If
                dr.Close()

                ' SAVE ACCESS TOKEN
                If TotalDocs > 0 Then
                    SqlS = "UPDATE siebeldb.dbo.CX_SUB_CON SET TOKEN='" & AccessToken & "' WHERE CON_ID='" & CONTACT_ID & "'"
                    temp = ExecQuery("Update", "CX_SUB_CON", cmd, SqlS, mydebuglog, Debug)
                    AutoStart = False
                End If
                
                ' SETUP PLAYER URL
                If HCIPlayer Then
                    If InStr(1, CRSE_CONTENT_URL, myprotocol & ":") = 0 Then
                        If myprotocol = "http" Then CRSE_CONTENT_URL = Replace(CRSE_CONTENT_URL, "https:", "http:")
                        If myprotocol = "https" Then CRSE_CONTENT_URL = Replace(CRSE_CONTENT_URL, "http:", "https:")
                    End If
                    If Debug = "Y" Then mydebuglog.Debug("  .. HCIPlayer CRSE_CONTENT_URL : " & CRSE_CONTENT_URL)
                Else
                    If SCORM = "Y" Then
                        CRSE_CONTENT_URL = myprotocol & "//hciscorm.certegrity.com/scorm/defaultui/launch.aspx?"
                    End If
                    If TEST_FLG = "Y" Then
                        CRSE_CONTENT_URL = myprotocol & "//hciscorm.certegrity.com/scorm/defaultui/launch.aspx?"
                    End If
                    If Debug = "Y" Then mydebuglog.Debug("  .. CRSE_CONTENT_URL : " & CRSE_CONTENT_URL)
                End If
                
                ' GENERATE THE LINK
                If HCIPlayer Then
                    ClassLink = CRSE_CONTENT_URL & "CrseId=" & CRSE_ID & "&CrseType=C&VersionId=0&RegId=" & REG_ID & "&UserId=" & UID & "&InstanceId=0&Debug=Y"
                Else
                    ClassLink = CRSE_CONTENT_URL & "registration=CrseId|" & CRSE_ID & "!CrseType|C!RegId|" & REG_ID & "!UserId|" & UID & "!InstanceId|0&configuration=Popup|false!DiagnosticsLog|false!DiagnosticsDetailedLog|false&forceFrameset=true&player=modern"
                End If
                If Debug = "Y" Then mydebuglog.Debug("  .. ClassLink : " & ClassLink)
                
                ' LOG THE ENTRANCE TO THE COURSE
                'SqlS = "INSERT INTO siebeldb.dbo.CX_TRAIN_OFFR_ACCESS(ROW_ID, CREATED, CREATED_BY, LAST_UPD, LAST_UPD_BY, " & _
                '    "MODIFICATION_NUM, CONFLICT_ID, REG_ID, ENTER_FLG, EXIT_FLG, MOBILE) " & _
                '    "SELECT TOP 1 '" & REG_ID & "-'+LTRIM(CAST(COUNT(*)+1 AS VARCHAR)) ,GETDATE(), '0-1', GETDATE(), " & _
                '    "'0-1', 0, 0, '" & REG_ID & "','Y','N','Y' " & _
                '    "FROM siebeldb.dbo.CX_TRAIN_OFFR_ACCESS " & _
                '    "WHERE REG_ID = '" & REG_ID & "'"
                SqlS = "IF (SELECT TOP 1 ENTER_FLG FROM siebeldb.dbo.CX_TRAIN_OFFR_ACCESS WHERE REG_ID='" & REG_ID & "' ORDER BY CREATED DESC)='N' OR " &
                        "(SELECT TOP 1 ENTER_FLG FROM siebeldb.dbo.CX_TRAIN_OFFR_ACCESS WHERE REG_ID='" & REG_ID & "' ORDER BY CREATED DESC) IS NULL BEGIN; " &
                        "INSERT INTO siebeldb.dbo.CX_TRAIN_OFFR_ACCESS(ROW_ID, CREATED, CREATED_BY, LAST_UPD, LAST_UPD_BY, " &
                        "MODIFICATION_NUM, CONFLICT_ID, REG_ID, ENTER_FLG, EXIT_FLG, MOBILE) " &
                        "SELECT '" & REG_ID & "-'+LTRIM(CAST(COUNT(*)+1 AS VARCHAR)),GETDATE(),'0-1',GETDATE(),'0-1'," &
                        "0,0,'" & REG_ID & "','Y','N','Y' " &
                        "FROM siebeldb.dbo.CX_TRAIN_OFFR_ACCESS " &
                        "WHERE REG_ID='" & REG_ID & "'; END;"
                temp = ExecQuery("Insert", "CX_TRAIN_OFFR_ACCESS", cmd, SqlS, mydebuglog, Debug)
                
                ' Log to CM activity log
                SqlS = "INSERT INTO reports.dbo.CM_LOG(REG_ID, SESSION_ID, RECORD_ID, REMOTE_ADDR, ACTION, BROWSER) " & _
                    "VALUES('" & UID & "','" & SessID & "','" & REG_ID & "','" & callip & "','ENTERED COURSE','" & Left(BROWSER, 200) & "')"
                temp = ExecQuery("Insert", "CM_LOG", cmd, SqlS, mydebuglog, Debug)
                
                ' UPDATE REGISTRATION STATUS IF NECESSARY
                If REG_STATUS_CD = "Accepted" Then
                    SqlS = "UPDATE siebeldb.dbo.CX_SESS_REG SET STATUS_CD='In Progress', LAST_UPD=GETDATE() WHERE ROW_ID='" & REG_ID & "'"
                    temp = ExecQuery("Update", "CX_SESS_REG", cmd, SqlS, mydebuglog, Debug)
                End If
                
                ' VERIFY LAST_INST IN SUB_CON_ID - NEEDED FOR REDIRECT BACK
                If LANG_CD <> "ENU" Then
                    LAST_INST = myprotocol & "//hciscorm.certegrity.com/ls/" & LANG_CD & "/FinishClass.html?RID=" & REG_ID & "&UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "&HP=" & HOME_PAGE & "&LANG=" & LANG_CD & "&CUR=" & CURRENT_PAGE
                Else
                    LAST_INST = myprotocol & "//hciscorm.certegrity.com/ls/FinishClass.html?RID=" & REG_ID & "&UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "&HP=" & HOME_PAGE & "&CUR=" & CURRENT_PAGE
                End If
                SqlS = "UPDATE siebeldb.dbo.CX_SUB_CON SET LAST_INST='" & LAST_INST & "' " & _
                "WHERE CON_ID='" & CONTACT_ID & "'"
                temp = ExecQuery("Update", "CX_SESS_REG", cmd, SqlS, mydebuglog, Debug)

            Else
                ' ====================
                ' Registration Status does not allow class to be started

                ' ALTERNATIVE MESSAGE   
                ClassLink = ""
                Select Case LANG_CD
                    Case "ESN"
                        MSG1 = "<table width=""100%"" height=""100%""><tr valign=""Middle""><td class=""BigHeader""><font color=""Gray"">Esta clase no es accesible para ti. Si cree que esto es un error, póngase en contacto con nuestro <A HREF=""JavaScript:openNewWindow('" & mailpath & "/message?OpenForm&param1=Technical+Support&param2=Curso+Acceso+Pregunta&DOM=" & DOMAIN & "',525,375)""><b>Soporte técnico</b></a><br><br>" & _
                        "Haga clic en <a href=[RETURN] class=""button"">Regrese a la pantalla anterior</a></font></td></tr></table>"
                    Case Else
                        MSG1 = "<table width=""100%"" height=""100%""><tr valign=""Middle""><td class=""BigHeader""><font color=""Gray"">This class is not accessible to you.  If you believe that this is in error, please contact our <A HREF=""JavaScript:openNewWindow('" & mailpath & "/message?OpenForm&param1=Technical+Support&param2=Course+Access+Question&DOM=" & DOMAIN & "',525,375)""><b>Technical Support</b></a> department<br><br>" & _
                        "Click to <a href=[RETURN] class=""button"">Return to the previous screen</a></font></td></tr></table>"
                End Select
                    
                ' REGISTRATION NOT COMPLETE
                If REG_STATUS_CD = "Tentative" Then
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = "<table width=""100%"" height=""100%""><tr valign=""Middle""><td class=""BigHeader""><font color=""Gray"">Hay un problema con su registro: puede que no esté completo o que no se haya recibido el pago. Si cree que esto es un error, póngase en contacto con nuestro <A HREF=""JavaScript:openNewWindow('" & mailpath & "/message?OpenForm&param1=Customer+Service&param2=Curso+Acceso+Pregunta&DOM=" & DOMAIN & "',525,375)""><b>Servicio al Cliente</b></a>" & _
                            "Haga clic en <a href=[RETURN] class=""button"">Regrese a la pantalla anterior</a></font></td></tr></table>"
                        Case Else
                            MSG1 = "<table width=""100%"" height=""100%""><tr valign=""Middle""><td class=""BigHeader""><font color=""Gray"">There is a problem with your registration - it may not be complete or payment may not have been received. If you believe that this is in error, please contact our <A HREF=""JavaScript:openNewWindow('" & mailpath & "/message?OpenForm&param1=Customer+Service&param2=Course+Access+Question&DOM=" & DOMAIN & "',525,375)""><b>Customer Service</b></a> department" & _
                            "Click to <a href=[RETURN] class=""button"">Return to the previous screen</a></font></td></tr></table>"
                    End Select
                End If
                
                ' REGISTRATION CANCELLED/DECLINED
                If REG_STATUS_CD = "Cancelled" Or REG_STATUS_CD = "Declined" Then
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = "<table width=""100%"" height=""100%""><tr valign=""Middle""><td class=""BigHeader""><font color=""Gray"">Hay un problema con su registro. Si cree que esto es un error, póngase en contacto con nuestro <A HREF=""JavaScript:openNewWindow('" & mailpath & "/message?OpenForm&param1=Customer+Service&param2=Curso+Acceso+Pregunta&DOM=" & DOMAIN & "',525,375)""><b>Servicio al Cliente</b></a></font><br><br>" & _
                            "Haga clic en <a href=[RETURN] class=""button"">Regrese a la pantalla anterior</a></td></tr></table>"
                        Case Else
                            MSG1 = "<table width=""100%"" height=""100%""><tr valign=""Middle""><td class=""BigHeader""><font color=""Gray"">There is a problem with your registration. If you believe that this is in error, please contact our <A HREF=""JavaScript:openNewWindow('" & mailpath & "/message?OpenForm&param1=Customer+Service&param2=Course+Access+Question&DOM=" & DOMAIN & "',525,375)""><b>Customer Service</b></a> department</font><br><br>" & _
                            "Click to <a href=[RETURN] class=""button"">Return to the previous screen</a></td></tr></table>"
                    End Select
                End If
                
                ' IF THE CLASS IS FINISHED BUT THE EXAM IS PENDING
                If REG_STATUS_CD = "Exam Reqd" Then
                    AutoStart = False
                    If LANG_CD <> "ENU" Then
                        ClassLink = myprotocol & "//hciscorm.certegrity.com/ls/" & LANG_CD & "/OpenAssessment.html?RID=" & REG_ID & "&UID=" & UID & "&SES=" & SessID & "&HP=" & HOME_PAGE & "&LANG=" & LANG_CD & "&PP=" & DOMAIN
                    Else
                        ClassLink = myprotocol & "//hciscorm.certegrity.com/ls/OpenAssessment.html?RID=" & REG_ID & "&UID=" & UID & "&SES=" & SessID & "&HP=" & HOME_PAGE & "&PP=" & DOMAIN
                    End If
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = "<center><span class=""BigHeader""><font color=""Gray"">Ha completado la clase pero aún no ha completado el examen de certificación.<br><br><a href=" & ClassLink & " class=""button"">Haga clic aquí para rendir el examen de certificación</a></font></span></center>"
                        Case Else
                            MSG1 = "<center><span class=""BigHeader""><font color=""Gray"">You have completed the class but have not yet completed the certification exam for it.<br><br><a href=" & ClassLink & " class=""button"">Click here to take the certification exam</a></font></span></center>"
                    End Select
                End If

                ' IF THE CLASS IS FINISHED AND EXAM COMPLETED
                If REG_STATUS_CD = "Completed" Then
                    PAUSE = "3"
                    If LANG_CD <> "ENU" Then
                        ClassLink = myprotocol & "//hciscorm.certegrity.com/ls/" & LANG_CD & "/FinishClass.html?RID=" & REG_ID & "&UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "&LANG=" & LANG_CD & "&HP=" & HOME_PAGE & "&CUR=" & CURRENT_PAGE
                    Else
                        ClassLink = myprotocol & "//hciscorm.certegrity.com/ls/FinishClass.html?RID=" & REG_ID & "&UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "&HP=" & HOME_PAGE & "&CUR=" & CURRENT_PAGE
                    End If
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = "<span class=""BigHeader""><font color=""Gray"">Has completado la clase. Un momento por favor..</font></span>"
                        Case Else
                            MSG1 = "<span class=""BigHeader""><font color=""Gray"">You have completed the class.  One moment please..</font></span>"
                    End Select
                End If
                
            End If

            ' GENERATE CONDITIONAL FLAGS
            If X_FORMAT = "HTML5" Then
                If HCIPlayer Then
                    If TotalDocs = 0 Then
                        eCreateClass = "CreateCourse"
                    End If
                Else
                    If TotalDocs = 0 Then
                        eCreateClass = "CreateCourse"
                        eAcceptClass = "AcceptClass"
                    Else
                        eAcceptClass = "AcceptClass"
                    End If
                End If
            Else
                eCreateClass = "LoginCourse"
                eAcceptClass = "AcceptClass"
            End If
            If Debug = "Y" Then
                mydebuglog.Debug("  .. PAUSE : " & PAUSE)
                mydebuglog.Debug("  .. MSG1 : " & MSG1)
                mydebuglog.Debug("  .. ClassLink : " & ClassLink)
                mydebuglog.Debug("  .. AutoStart : " & AutoStart)
                mydebuglog.Debug("  .. eCreateClass : " & eCreateClass)
                mydebuglog.Debug("  .. eAcceptClass : " & eAcceptClass)
            End If
            
            ' PREPARE OUTPUT
            ' Booleans-
            jdoc = jdoc & """KBA_FLAG"":""" & KBA_FLAG & ""","
            jdoc = jdoc & """OldScorm"":""" & OldScorm & ""","
            jdoc = jdoc & """AutoStart"":""" & AutoStart & ""","
            jdoc = jdoc & """Olark"":""" & Olark & ""","
            jdoc = jdoc & """HCIPlayer"":""" & HCIPlayer & ""","
            jdoc = jdoc & """AnotherWindow"":""" & AnotherWindow & ""","
      
            ' Computed Strings-
            jdoc = jdoc & """ClassLink"":""" & EscapeJSON(ClassLink) & ""","
            jdoc = jdoc & """MSG1"":""" & EscapeJSON(MSG1) & ""","
            jdoc = jdoc & """MSG2"":""" & EscapeJSON(MSG2) & ""","
            jdoc = jdoc & """eCreateClass"":""" & eCreateClass & ""","
            jdoc = jdoc & """eAcceptClass"":""" & eAcceptClass & ""","
            If TotalDocs > 0 Then
                jdoc = jdoc & """DOCLIST"":""" & EscapeJSON(CPage) & ""","
            Else
                jdoc = jdoc & """DOCLIST"":"""","
            End If
            jdoc = jdoc & """LAST_INST"":""" & EscapeJSON(LAST_INST) & ""","
            
            ' Retrieved strings-
            jdoc = jdoc & """JURIS"":""" & JURIS & ""","
            jdoc = jdoc & """JURIS_LVL"":""" & JURIS_LVL & ""","
            jdoc = jdoc & """JURIS_ID"":""" & EscapeJSON(JURIS_ID) & ""","
            jdoc = jdoc & """COURSE"":""" & COURSE & ""","
            jdoc = jdoc & """FST_NAME"":""" & FST_NAME & ""","
            jdoc = jdoc & """LAST_NAME"":""" & EscapeJSON(LAST_NAME) & ""","
            jdoc = jdoc & """EMAIL_ADDR"":""" & EscapeJSON(EMAIL_ADDR) & ""","
            jdoc = jdoc & """KBA_NOTICE"":""" & EscapeJSON(KBA_NOTICE) & ""","
            jdoc = jdoc & """ConfirmEmail"":""" & EscapeJSON(ConfirmEmail) & ""","
            jdoc = jdoc & """X_FORMAT"":""" & EscapeJSON(X_FORMAT) & ""","
            jdoc = jdoc & """REG_STATUS_CD"":""" & EscapeJSON(REG_STATUS_CD) & ""","
            jdoc = jdoc & """CRSE_ID"":""" & EscapeJSON(CRSE_ID) & ""","
            jdoc = jdoc & """RESOLUTION"":""" & EscapeJSON(RESOLUTION) & ""","
            jdoc = jdoc & """RES_X"":""" & EscapeJSON(RES_X) & ""","
            jdoc = jdoc & """RES_Y"":""" & EscapeJSON(RES_Y) & ""","
            jdoc = jdoc & """KBA_REQD"":""" & EscapeJSON(KBA_REQD) & ""","
            jdoc = jdoc & """USER_NAME"":""" & EscapeJSON(USER_NAME) & ""","
            jdoc = jdoc & """LANG_CD"":""" & EscapeJSON(LANG_CD) & ""","
            
            ' KBA-
            '   Numbers-
            '       KBA_COUNT : Number of questions available for a jurisdiction
            '       KBA_QUESTIONS : Number of questions required by the jurisdiction - "0" means all of the questions available
            '       TO_ASK : Number of questions to ask
            '       NUM_ANSRD : Number of questions answered by the student      
            '    Arrays-
            '      Q_ID(TO_ASK)
            '      Q_TEXT(TO_ASK)
            jdoc = jdoc & """KBA_COUNT"":""" & Trim(Str(KBA_COUNT)) & ""","
            jdoc = jdoc & """KBA_QUESTIONS"":""" & Trim(Str(KBA_QUESTIONS)) & ""","
            jdoc = jdoc & """TO_ASK"":""" & Trim(Str(TO_ASK)) & ""","
            jdoc = jdoc & """NUM_ANSRD"":""" & Trim(Str(NUM_ANSRD)) & ""","
            Dim kba_items As String
            Dim i As Integer
            If TO_ASK = 0 Then
                kba_items = """kba"": []"
            Else
                If Debug = "Y" Then mydebuglog.Debug("  KBA:")
                kba_items = """kba"": ["
                For i = 1 To TO_ASK
                    If Debug = "Y" Then mydebuglog.Debug("   " & Q_ID(i) & " : " & Q_TEXT(i))
                    kba_items = kba_items & "{ ""Q_ID"": """ & Q_ID(i) & """,""Q_TEXT"":""" & Q_TEXT(i) & """}, "
                Next
                kba_items = Left(kba_items, Len(kba_items) - 2) & " ]"
            End If
            jdoc = jdoc & kba_items & ","
   
            ' Identity fields-      
            jdoc = jdoc & """REG_ID"":""" & REG_ID & ""","
            jdoc = jdoc & """Id"":""" & UID & ""","
            jdoc = jdoc & """SessId"":""" & SessID & ""","
            jdoc = jdoc & """HOME_PAGE"":""" & HOME_PAGE & ""","
            jdoc = jdoc & """DOMAIN"":""" & DOMAIN & ""","
        Else
            GoTo DBError
        End If
        GoTo CloseOut
        
AccessCountError:
        If Debug = "Y" Then mydebuglog.Debug(">>AccessCountError")
        Select Case LANG_CD
            Case "ESN"
                errmsg = "Ha superado el acceso normal para su registro y se ha puesto en espera en espera de una revisi&oacute;n."
            Case Else
                errmsg = "You exceeded normal access for your registration and it has been placed On Hold pending a review."
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
        jdoc = jdoc & """ErrMsg"":""" & errmsg & """"
        jdoc = callback & "({""ResultSet"": {" & jdoc & "} })"
        
        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("WsGetClass.ashx : " & ErrLvl & ": " & Trim(errmsg))
        myeventlog.Info("WsGetClass.ashx : UID: " & UID & ", SessID: " & SessID & ", and RegID:" & REG_ID & ", Classlink: " & ClassLink)
        If Debug = "Y" Or (Logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("  JDOC: " & jdoc & vbCrLf)
                mydebuglog.Debug("Results:  UID: " & UID & ", SessID: " & SessID & ", and RegID:" & REG_ID & ", Classlink: " & ClassLink)
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
                LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, "WSGETCLASS", LogStartTime, VersionNum, Debug)
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