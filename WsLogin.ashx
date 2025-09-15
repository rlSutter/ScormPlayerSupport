<%@ WebHandler Language="VB" Class="WsLogin" %>

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

Public Class WsLogin : Implements IHttpHandler

    Enum enumObjectType
        StrType = 0
        IntType = 1
        DblType = 2
        DteType = 3
    End Enum

    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest

        ' Parameter Declarations
        Dim Debug, temp As String

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String
        Dim DBisOpen As Boolean = False

        ' Logging declarations
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("WsLoginDebugLog")
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
        Dim REF_URL As String = Trim(context.Request.ServerVariables("HTTP_REFERER"))
        Dim HTTP_REFERER As String = Trim(context.Request.ServerVariables("HTTP_REFERER"))
        Dim REMOTE_ADDR As String = Trim(context.Request.ServerVariables("REMOTE_ADDR"))
        Dim HTTP_HOST As String = Trim(context.Request.ServerVariables("HTTP_HOST"))
        Dim BROWSER As String = Trim(context.Request.ServerVariables("HTTP_USER_AGENT"))
        Dim QueryString As String = Trim(context.Request.QueryString.ToString)
        Dim UserID, SessID As String
        Try
            UserID = Trim(context.Request.Cookies.Item("ID").Value.ToString())
        Catch ex As Exception
            UserID = ""
        End Try
        Try
            SessID = Trim(context.Request.Cookies.Item("Sess").Value.ToString())
        Catch ex As Exception
            SessID = ""
        End Try

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service

        ' Variable declarations
        Dim errmsg, ErrLvl As String
        Dim REG_NUM, HOME_PAGE, LANG_CD, UID, CURRENT_PAGE, callback, DOMAIN As String
        Dim LOGGED_IN, CONTACT_ID, SUB_ID, CONTACT_OU_ID As String
        Dim MSG1, ReDirect, NextLink, ExamLink, ClassLink, Refresh, LaunchProtocol As String
        Dim LOGIN_REDIRECT, pUserID, pSessID, REDIRECT_REC, ALT_RECORD_ID, REF_CON_ID, FORM_TYPE As String
        Dim PROCTOR_ID, CURRCLM_ID, INVOICE_NUM, PUBLIC_FLG, HIST_ID, sUserID, output As String
        Dim TERM_FLG, SUB_CON_ID, SVC_TYPE, CON_ID, TRAINING_ACCESS, TRAINER_ACC_FLG, PAID_USER_FLG, NO_HOME As String
        Dim HOME_URL, UNSUB_URL, LOGOUT_URL, ETIPS_DOMAIN, SRC_URL, Path, SessionID, NXT, sReDirect, mReDirect As String
        Dim ReLogin, Student, Mobile, logout_flag As Boolean

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
        REG_NUM = ""
        UID = ""
        SessID = ""
        CURRENT_PAGE = ""
        ErrLvl = "Error"
        MSG1 = ""
        ReDirect = ""
        NextLink = ""
        ExamLink = ""
        ClassLink = ""
        Refresh = ""
        LaunchProtocol = "http:"
        pUserID = ""
        pSessID = ""
        REF_URL = ""
        LOGIN_REDIRECT = ""
        REDIRECT_REC = ""
        ALT_RECORD_ID = ""
        REF_CON_ID = ""
        FORM_TYPE = ""
        PROCTOR_ID = ""
        CURRCLM_ID = ""
        INVOICE_NUM = ""
        PUBLIC_FLG = ""
        HIST_ID = ""
        ReLogin = False
        sUserID = ""
        output = ""
        HOME_URL = ""
        UNSUB_URL = ""
        LOGOUT_URL = ""
        ETIPS_DOMAIN = ""
        SRC_URL = ""
        TERM_FLG = ""
        SUB_CON_ID = ""
        SVC_TYPE = ""
        CON_ID = ""
        TRAINING_ACCESS = ""
        TRAINER_ACC_FLG = ""
        PAID_USER_FLG = ""
        NO_HOME = ""
        Path = ""
        SessionID = ""
        NXT = "1"
        sReDirect = ""
        mReDirect = ""
        Student = False
        Mobile = False
        logout_flag = False

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
            tempdebug = System.Configuration.ConfigurationManager.AppSettings.Get("WsLogin_debug")
            If tempdebug = "Y" And Debug <> "T" Then Debug = "Y"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("LaunchProtocol")
            If temp <> "" Then LaunchProtocol = temp
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open log file if applicable
        If Logging = "Y" Then
            logfile = "C:\Logs\WsLogin.log"
            Try
                log4net.GlobalContext.Properties("WsLoginLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                GoTo CloseOut2
            End Try
        End If

        ' ============================================
        ' Get parameters    
        If Not context.Request.QueryString("ID") Is Nothing Then
            pUserID = Trim(context.Request.QueryString("ID"))
            pUserID = SimpleString(pUserID)
        End If

        If Not context.Request.QueryString("SESS") Is Nothing Then
            pSessID = Trim(context.Request.QueryString("SESS"))
            pSessID = SimpleString(pSessID)
        End If

        If Not context.Request.QueryString("REF") Is Nothing Then
            REF_URL = Trim(context.Request.QueryString("REF"))
            REF_URL = SimpleString2(REF_URL)
        End If

        If Not context.Request.QueryString("RD") Is Nothing Then
            ReDirect = Trim(context.Request.QueryString("RD"))
            ReDirect = SimpleString2(ReDirect)
        End If

        If Not context.Request.QueryString("DOM") Is Nothing Then
            DOMAIN = Trim(UCase(context.Request.QueryString("DOM")))
        End If

        If Not context.Request.QueryString("PP") Is Nothing And DOMAIN = "" Then
            DOMAIN = Trim(UCase(context.Request.QueryString("PP")))
        End If
        If InStr(DOMAIN, ",") > 0 Then
            DOMAIN = Left(DOMAIN, InStr(DOMAIN, ",") - 1)
        End If
        DOMAIN = SimpleString(DOMAIN)

        If Not context.Request.QueryString("NH") Is Nothing Then
            NO_HOME = Trim(context.Request.QueryString("NH"))
            If NO_HOME <> "Y" Then NO_HOME = ""
        End If

        If Not context.Request.QueryString("RNL") Is Nothing Then
            LOGIN_REDIRECT = Trim(UCase(context.Request.QueryString("RNL")))
            LOGIN_REDIRECT = SimpleString2(LOGIN_REDIRECT)
        End If
        If InStr(LOGIN_REDIRECT, ",") > 0 Then
            LOGIN_REDIRECT = Left(LOGIN_REDIRECT, InStr(LOGIN_REDIRECT, ",") - 1)
            LOGIN_REDIRECT = SimpleString2(LOGIN_REDIRECT)
        End If

        If Not context.Request.QueryString("RID") Is Nothing Then
            REDIRECT_REC = Trim(UCase(context.Request.QueryString("RID")))
            REDIRECT_REC = SimpleString2(REDIRECT_REC)
        End If
        If InStr(REDIRECT_REC, ",") > 0 Then
            REDIRECT_REC = Left(REDIRECT_REC, InStr(REDIRECT_REC, ",") - 1)
            REDIRECT_REC = SimpleString2(REDIRECT_REC)
        End If

        If Not context.Request.QueryString("AID") Is Nothing Then
            ALT_RECORD_ID = Trim(UCase(context.Request.QueryString("AID")))
            ALT_RECORD_ID = SimpleString2(ALT_RECORD_ID)
        End If

        If Not context.Request.QueryString("RCN") Is Nothing Then
            REF_CON_ID = Trim(UCase(context.Request.QueryString("RCN")))
            REF_CON_ID = SimpleString2(REF_CON_ID)
        End If

        If Not context.Request.QueryString("LANG") Is Nothing Then
            LANG_CD = UCase(context.Request.QueryString("LANG"))
            LANG_CD = SimpleString(LANG_CD)
            If Len(LANG_CD) > 5 Then GoTo AccessError
        End If
        If LANG_CD <> "ESN" And LANG_CD <> "ENU" Then LANG_CD = "ENU"

        If Not context.Request.QueryString("FT") Is Nothing Then
            FORM_TYPE = UCase(context.Request.QueryString("FT"))
        End If

        If Debug = "Y" Then
            mydebuglog.Debug("----------------------------------")
            mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
            mydebuglog.Debug("Parameters-")
            mydebuglog.Debug("  callip: " & callip)
            mydebuglog.Debug("  LaunchProtocol: " & LaunchProtocol)
            mydebuglog.Debug("  HTTP_HOST: " & HTTP_HOST)
            mydebuglog.Debug("  REMOTE_ADDR: " & REMOTE_ADDR)
            mydebuglog.Debug("  BROWSER: " & BROWSER)
            mydebuglog.Debug("  REF_URL: " & REF_URL)
            mydebuglog.Debug("  QueryString: " & QueryString)
            mydebuglog.Debug("  Cookie User Id: " & UserID)
            mydebuglog.Debug("  Cookie Session Id: " & SessID)
            mydebuglog.Debug("  Parameter User Id: " & pUserID)
            mydebuglog.Debug("  Parameter Session Id: " & pSessID)
            mydebuglog.Debug("  ReDirect: " & ReDirect)
            mydebuglog.Debug("  DOMAIN: " & DOMAIN)
            mydebuglog.Debug("  LOGIN_REDIRECT: " & LOGIN_REDIRECT)
            mydebuglog.Debug("  REDIRECT_REC: " & REDIRECT_REC)
            mydebuglog.Debug("  ALT_RECORD_ID: " & ALT_RECORD_ID)
            mydebuglog.Debug("  REF_CON_ID: " & REF_CON_ID)
            mydebuglog.Debug("  LANG_CD: " & LANG_CD)
            mydebuglog.Debug("  NO_HOME: " & NO_HOME)
            mydebuglog.Debug("  FORM_TYPE: " & FORM_TYPE & vbCrLf)
        End If

        If Len(pUserID) > 16 Then GoTo AccessError
        If Len(pSessID) > 16 Then GoTo AccessError
        If Len(DOMAIN) > 15 Then GoTo AccessError
        If Len(ALT_RECORD_ID) > 30 Then GoTo AccessError
        If Len(REDIRECT_REC) > 30 Then GoTo AccessError
        If Len(REF_CON_ID) > 16 Then GoTo AccessError

        ' ============================================
        ' Form specific parameters
        If Not context.Request.QueryString("PCN") Is Nothing Then
            PROCTOR_ID = UCase(context.Request.QueryString("PCN"))
        End If

        If Not context.Request.QueryString("CON") Is Nothing Then
            CONTACT_ID = UCase(context.Request.QueryString("CON"))
        End If

        If Not context.Request.QueryString("CUR") Is Nothing Then
            CURRCLM_ID = UCase(context.Request.QueryString("CUR"))
        End If

        If Not context.Request.QueryString("INV") Is Nothing Then
            INVOICE_NUM = UCase(context.Request.QueryString("INV"))
        End If

        If Not context.Request.QueryString("PUB") Is Nothing Then
            PUBLIC_FLG = UCase(context.Request.QueryString("PUB"))
        End If

        If Debug = "Y" Then
            mydebuglog.Debug("Form specific parameters-")
            mydebuglog.Debug("  PROCTOR_ID: " & PROCTOR_ID)
            mydebuglog.Debug("  CONTACT_ID: " & CONTACT_ID)
            mydebuglog.Debug("  CURRCLM_ID: " & REF_CON_ID)
            mydebuglog.Debug("  INVOICE_NUM: " & INVOICE_NUM)
            mydebuglog.Debug("  PUBLIC_FLG: " & PUBLIC_FLG & vbCrLf)
        End If

        ' If not logged in already and not referred, then access error
        If pUserID = "" And UserID = "" Then
            ' Exceptions to requiring login information
            Select Case Trim(LCase(ReDirect))
                Case "newexam"
                    GoTo OpenDB
                Case "openexam"
                    GoTo OpenDB
                Case "openreg"
                    GoTo OpenDB
                Case "openexamreg?openagent"
                    GoTo OpenDB
            End Select

            ' Exceptions to requiring login information
            Select Case Trim(LCase(LOGIN_REDIRECT))
                Case "newexam"
                    GoTo OpenDB
                Case "openexam"
                    GoTo OpenDB
                Case "openreg"
                    GoTo OpenDB
                Case "openexamreg?openagent"
                    GoTo OpenDB
            End Select
            GoTo AccessError
        End If

        ' ============================================
        ' Open database connection 
OpenDB:
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            errmsg = errmsg & "Unable to open the database connection. " & vbCrLf
            GoTo CloseOut
        End If
        DBisOpen = True

        ' ============================================
        ' Prepare results
        If Not cmd Is Nothing Then

            ' ================================================   
            ' Checks IDs
            If pUserID <> "" And UserID <> "" Then
                If pUserID <> UserID Then

                    ' Check To see if the parameter supplied is for a valid login
                    SqlS = "SELECT ROW_ID " & _
                    "FROM siebeldb.dbo.CX_SUB_CON_HIST " & _
                    "WHERE USER_ID='" & pUserID & "' AND SESSION_ID='" & pSessID & "'"
                    If Debug = "Y" Then mydebuglog.Debug("User Query: " & vbCrLf & "  " & SqlS)
                    Try
                        cmd.CommandText = SqlS
                        dr = cmd.ExecuteReader()
                        If Not dr Is Nothing Then
                            While dr.Read()
                                HIST_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            End While
                        End If
                    Catch ex As Exception
                        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Unable to locate credentials. Error: " & vbCrLf & ex.ToString & vbCrLf)
                        GoTo AccessError
                    End Try
                    dr.Close()
                    If Debug = "Y" Then mydebuglog.Debug("  .. LOGGED_IN: " & LOGGED_IN & vbCrLf)

                    ' Assume that the parameter is correct - that an old cookie exists due to an incomplete logout
                    If HIST_ID <> "" Then
                        UserID = pUserID
                        SessID = pSessID
                        ReLogin = True
                    End If

                End If
            End If

            ' ================================================   
            ' LOOKUP SUBSCRIPTION
            ' If unable to locate, then subscription not setup, error out
            sUserID = Trim(UserID)
            If sUserID = "" Then sUserID = Trim(pUserID)
            If sUserID = "" Then
                Select Case Trim(LCase(ReDirect))
                    Case "newexam"
                        GoTo LocateInstance
                    Case "openexam"
                        GoTo LocateInstance
                    Case "reviewexam"
                        GoTo LocateInstance
                    Case "openreg"
                        GoTo LocateInstance
                    Case "openexamreg?openagent"
                        GoTo LocateInstance
                End Select

                Select Case Trim(LCase(LOGIN_REDIRECT))
                    Case "newexam"
                        GoTo LocateInstance
                    Case "openexam"
                        GoTo LocateInstance
                    Case "openreg"
                        GoTo LocateInstance
                    Case "openexamreg?openagent"
                        GoTo LocateInstance
                End Select
                GoTo AccessError
            End If

            SqlS = "SELECT TOP 1 C.SUB_ID, (SELECT CASE WHEN (S.SVC_TERM_DT<GETDATE() OR (C.USER_EXP_DATE<GETDATE() AND C.USER_EXP_DATE IS NOT NULL)) AND S.SVC_TYPE<>'PUBLIC ACCESS' THEN 'Y' ELSE 'N' END) AS TERM_FLG, " & _
                   "C.ROW_ID AS SUB_CON_ID, S.SVC_TYPE, S.DOMAIN, P.ROW_ID, C.TRAINING_ACCESS, C.TRAINER_ACC_FLG, C.PAID_USER_FLG " & _
                   " FROM siebeldb.dbo.S_CONTACT P " & _
                   "LEFT OUTER JOIN siebeldb.dbo.CX_SUB_CON C ON C.CON_ID=P.ROW_ID " & _
                   "LEFT OUTER JOIN siebeldb.dbo.CX_SUBSCRIPTION S ON S.ROW_ID=C.SUB_ID " & _
                   "WHERE P.X_REGISTRATION_NUM='" & sUserID & "'"
            If Debug = "Y" Then mydebuglog.Debug("Subscription Query: " & vbCrLf & "  " & SqlS)
            Try
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        SUB_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                        TERM_FLG = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                        SUB_CON_ID = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                        SVC_TYPE = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                        If SVC_TYPE = "PUBLIC ACCESS" Then TERM_FLG = "N"
                        temp = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                        If temp <> DOMAIN And DOMAIN <> "" Then
                            logout_flag = True
                        End If
                        CON_ID = Trim(CheckDBNull(dr(5), enumObjectType.StrType))
                        TRAINING_ACCESS = Trim(CheckDBNull(dr(6), enumObjectType.StrType))
                        TRAINER_ACC_FLG = Trim(CheckDBNull(dr(7), enumObjectType.StrType))
                        PAID_USER_FLG = Trim(CheckDBNull(dr(8), enumObjectType.StrType))
                    End While
                Else
                    GoTo SubscriptionError
                End If
            Catch ex As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Unable to locate credentials. Error: " & vbCrLf & ex.ToString & vbCrLf)
                GoTo SystemUnavailable
            End Try
            dr.Close()
            If Debug = "Y" Then
                mydebuglog.Debug("  .. TERM_FLG: " & TERM_FLG)
                mydebuglog.Debug("  .. CON_ID: " & CON_ID)
                mydebuglog.Debug("  .. DOMAIN: " & DOMAIN)
                mydebuglog.Debug("  .. SVC_TYPE: " & SVC_TYPE)
                mydebuglog.Debug("  .. TRAINING_ACCESS: " & TRAINING_ACCESS)
                mydebuglog.Debug("  .. TRAINER_ACC_FLG: " & TRAINER_ACC_FLG)
                mydebuglog.Debug("  .. PAID_USER_FLG: " & PAID_USER_FLG & vbCrLf)
            End If
            If SUB_CON_ID = "" Or SUB_ID = "" Then GoTo AccessError

            ' ================================================
            ' LOCATE DOMAIN INFORMATION            
            If DOMAIN <> "" Then
                SqlS = "SELECT HOME_URL, DEF_SUB_ID, UNSUB_URL, LOGOUT_URL, ETIPS_FLG, SRC_URL " & _
                       "FROM siebeldb.dbo.CX_SUB_DOMAIN WHERE DOMAIN='" & DOMAIN & "'"
                If Debug = "Y" Then mydebuglog.Debug("Domain Query: " & vbCrLf & "  " & SqlS)
                Try
                    cmd.CommandText = SqlS
                    dr = cmd.ExecuteReader()
                    If Not dr Is Nothing Then
                        While dr.Read()
                            HOME_URL = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            SUB_ID = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            UNSUB_URL = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                            LOGOUT_URL = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                            ETIPS_DOMAIN = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                            SRC_URL = Trim(CheckDBNull(dr(5), enumObjectType.StrType))
                        End While
                    Else
                        GoTo SubscriptionError
                    End If
                Catch ex As Exception
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Unable to locate credentials. Error: " & vbCrLf & ex.ToString & vbCrLf)
                    GoTo SystemUnavailable
                End Try
                dr.Close()
                If Debug = "Y" Then
                    mydebuglog.Debug("  .. HOME_URL: " & HOME_URL)
                    mydebuglog.Debug("  .. SUB_ID: " & SUB_ID & vbCrLf)
                End If
            End If

            ' ================================================
            ' DETERMINE IF WE NEED TO FIND AN INSTANCE
LocateInstance:
            If InStr(1, LCase(ReDirect), "/cp") > 0 Then GoTo ComputeNext
            Randomize()
            NXT = Chr(Str(Int(Rnd() * 3)) + 48)
            If Val(NXT) > 3 Then NXT = "1"
            If Val(NXT) < 1 Then NXT = "3"
            If NXT = "1" Then Path = LaunchProtocol & "//w2.certegrity.com/cp1.nsf"
            If NXT = "2" Then Path = LaunchProtocol & "//w3.certegrity.com/cp2.nsf"
            If NXT = "3" Then Path = LaunchProtocol & "//w3.certegrity.com/cp3.nsf"
            If NXT = "4" Then Path = LaunchProtocol & "//w4.certegrity.com/cp4.nsf"
            If NXT = "5" Then Path = LaunchProtocol & "//w4.certegrity.com/cp5.nsf"
            If Path = "" Then Path = LaunchProtocol & "//w2.certegrity.com/cp1.nsf"
            If Debug = "Y" Then
                mydebuglog.Debug("NXT: " & NXT)
                mydebuglog.Debug("Path: " & Path & vbCrLf)
            End If

ComputeNext:
            ' ==============================   
            ' GENERATE A SESSION ID
            If SessID = "" And pSessID = "" Then
                Randomize()
                SessionID = UCase(LoggingService.GeneratePassword(Debug)) & NXT & Chr(Str(Int(Rnd() * 26)) + 65) & Chr(Str(Int(Rnd() * 26)) + 65) & Chr(Str(Int(Rnd() * 26)) + 65)
            Else
                If SessID <> "" Then SessionID = Trim(SessID)
                If pSessID <> "" Then SessionID = Trim(pSessID)
            End If
            If Debug = "Y" Then mydebuglog.Debug("SessionID: " & SessionID & vbCrLf)

            ' ==============================      
            ' COMPUTE NEXT STEP
            '   If the referring hostname (REF_URL) is NOT for the domain that matches the domain supplied, then
            '   redirect (ReDirect) to the login system that is hosted for that domain first.  Otherwise, redirect to the
            '   specified page in Certification Manager

            ' Determine the source URL
            If SRC_URL = "" Then
                Select Case DOMAIN
                    Case "TIPS"
                        SRC_URL = "gettips.com"
                    Case "PBSA"
                        SRC_URL = "compliancetracking.com"
                    Case "IMIRDB"
                        SRC_URL = "ebevlaw.com"
                    Case "BARLAP"
                        SRC_URL = "barlap.us"
                    Case "HACCP"
                        SRC_URL = "haccp-training.com"
                    Case Else
                        SRC_URL = "compliancetracking.com"
                End Select
            End If
            If Debug = "Y" Then mydebuglog.Debug("SRC_URL: " & SRC_URL & vbCrLf)

            ' If the referenced url supplied doesn't match the source for the users domain, we need to go to that domain first 
            ' to login
            If REF_URL <> SRC_URL Then
                If Debug = "Y" Then mydebuglog.Debug("REF_URL: " & REF_URL & vbCrLf)
                Select Case REF_URL
                    Case "tipsuniversity.org"
                    Case "compliancetracking.com"
                    Case Else
                        Select Case DOMAIN
                            Case "TIPS"
                                'If InStr(1, ReDirect, "&") = 0 Then
                                'ReDirect = "https://w1.gettips.com/plogini.nsf/pWsLogin?OpenAgent&ID=" & sUserID & "&SESS=" & SessionID & "&RNL=" & ReDirect & "&REF=certegrity.com&DOM=" & DOMAIN
                                'Else
                                'ReDirect = "https://web.gettips.com/plogini.nsf/pWsLogin?OpenAgent&ID=" & sUserID & "&SESS=" & SessionID & "&RD=" & ReDirect & "&REF=certegrity.com&DOM=" & DOMAIN
                                'End If
                                'Path = ""
                            Case "BARLAP"
                            Case "IMIRDB"
                            Case "PBSA"
                            Case "HACCP"
                            Case Else
                        End Select
                End Select
                If Debug = "Y" Then mydebuglog.Debug("ReDirect: " & ReDirect & vbCrLf)
            End If

            ' ==============================   
            ' TRANSLATE REDIRECT
            '    If the initial redirection is to a service keyword, it has to be translated
            sReDirect = ReDirect
            Select Case ReDirect
                Case "home"
                    ReDirect = "Main?OpenForm&PP=" & DOMAIN
                Case "newexam"
                    ReDirect = "OpenExamReg?OpenAgent&PP=" & DOMAIN & "&ADD=Y&NPP=Y&PUB=Y&POP=N"
                    If pUserID = "" And UserID = "" Then ReDirect = ReDirect & "&LOG=Y"
                Case "openexam"
                    ReDirect = "OpenExamReg?OpenAgent&PP=" & DOMAIN & "&ADD=&NPP=Y&PUB=Y&POP=N"
                    If pUserID = "" And UserID = "" Then ReDirect = ReDirect & "&LOG=Y"
                Case "reviewexam"
                    ReDirect = "OpenExamReg?OpenAgent&PP=" & DOMAIN & "&ADD=N&NPP=Y&PUB=Y&POP=N"
                Case "surveylist"
                    ReDirect = "Main?OpenForm&PP=" & DOMAIN & "&RNL=" & ReDirect
                Case "assessments"
                    ReDirect = "Main?OpenForm&PP=" & DOMAIN & "&RNL=" & ReDirect
                Case "orderpass"
                    ReDirect = "Main?OpenForm&PP=" & DOMAIN & "&RNL=" & ReDirect
                Case "orderpartcards"
                    ReDirect = "Main?OpenForm&PP=" & DOMAIN & "&RNL=" & ReDirect
                Case "orderlogbook"
                    ReDirect = "Main?OpenForm&PP=" & DOMAIN & "&RNL=" & ReDirect
                Case "ordermaterials"
                    ReDirect = "Main?OpenForm&PP=" & DOMAIN & "&RNL=" & ReDirect
                Case "orderidguide"
                    ReDirect = "Main?OpenForm&PP=" & DOMAIN & "&RNL=" & ReDirect
                Case "orderreg"
                    ReDirect = "Main?OpenForm&PP=" & DOMAIN & "&RNL=" & ReDirect
                Case ""
                    If UserID <> "" Or pUserID <> "" Then
                        ReDirect = "Main?OpenForm&PP=" & DOMAIN
                    End If
            End Select
            If Debug = "Y" Then mydebuglog.Debug("ReDirect 1: " & ReDirect & vbCrLf)

            ' ==============================
            ' TRANSLATE LOGIN REDIRECT
            '    If the secondary redirection is to a service keyword, translation sometimes helps
            temp = LCase(LOGIN_REDIRECT)
            If InStr(1, temp, "http:") = 0 And InStr(1, temp, "https:") = 0 And temp <> "" And ReDirect = "" Then
                If InStr(temp, "home") > 0 Then
                    ReDirect = "Main?OpenForm&PP=" & DOMAIN
                End If
                If InStr(temp, "newexam") > 0 Then
                    ReDirect = "OpenExamReg?OpenAgent&PP=" & DOMAIN & "&ADD=Y&NPP=Y&PUB=Y&POP=N"
                    If pUserID = "" And UserID = "" Then ReDirect = ReDirect & "&LOG=Y"
                End If
                If InStr(temp, "openexam") > 0 Then
                    ReDirect = "OpenExamReg?OpenAgent&PP=" & DOMAIN & "&ADD=&NPP=Y&PUB=Y&POP=N"
                    If pUserID = "" And UserID = "" Then ReDirect = ReDirect & "&LOG=Y"
                End If
                If InStr(temp, "reviewexam") > 0 Then
                    If pUserID = "" And UserID = "" Then ReDirect = ReDirect & "&LOG=Y"
                End If
                If ReDirect = "" Then
                    If pUserID <> "" And UserID <> "" Then
                        ReDirect = "Main?OpenForm&PP=" & DOMAIN & "&RNL=" & LOGIN_REDIRECT
                    Else
                        ReDirect = ReDirect & temp
                    End If
                End If
            Else
                If LOGIN_REDIRECT <> "" Then
                    If InStr(1, LOGIN_REDIRECT, "&RNL") = 0 Then
                        ReDirect = ReDirect & "&RNL=" & Trim(LOGIN_REDIRECT)
                    Else
                        ReDirect = ReDirect & LOGIN_REDIRECT
                    End If
                End If
            End If
            If Debug = "Y" Then mydebuglog.Debug("ReDirect 2: " & ReDirect & vbCrLf)

            ' Add parameters
            If REDIRECT_REC <> "" Then ReDirect = ReDirect & "&RID=" & Trim(REDIRECT_REC)
            If ALT_RECORD_ID <> "" Then ReDirect = ReDirect & "&AID=" & ALT_RECORD_ID
            If REF_CON_ID <> "" Then ReDirect = ReDirect & "&RCN=" & REF_CON_ID

            ' Add form specific parameters
            If PROCTOR_ID <> "" Then ReDirect = ReDirect & "&PCN=" & PROCTOR_ID
            If CONTACT_ID <> "" Then ReDirect = ReDirect & "&CON=" & CONTACT_ID
            If CURRCLM_ID <> "" Then ReDirect = ReDirect & "&CUR=" & CURRCLM_ID
            If PUBLIC_FLG <> "" Then ReDirect = ReDirect & "&PUB=" & PUBLIC_FLG
            If INVOICE_NUM <> "" Then ReDirect = ReDirect & "&INV=" & INVOICE_NUM
            If DOMAIN <> "" And InStr(ReDirect, DOMAIN) = 0 And InStr(ReDirect, "RegisterPortal") = 0 Then ReDirect = ReDirect & "&PP=" & DOMAIN

            If Path <> "" And InStr(ReDirect, "RegisterPortal") = 0 Then ReDirect = Path & "/" & ReDirect
            If Path <> "" And InStr(LCase(ReDirect), "newrequest") > 0 Then
                Select Case FORM_TYPE
                    Case "PRC"
                        ReDirect = "https://www.gettips.com/forms/update_part_card.shtml"
                    Case "UNN"
                        ReDirect = "https://www.gettips.com/forms/unsubscribe_newsletter.shtml"
                    Case "UNS"
                        ReDirect = "https://www.gettips.com/forms/unsubscribe_emails.shtml"
                    Case "NOM"
                        ReDirect = "https://www.gettips.com/forms/trainer_recognition.shtml"
                    Case "NWS"
                        ReDirect = "https://www.gettips.com/forms/subscribe_newsletter.shtml"
                    Case "REF"
                        ReDirect = "https://www.gettips.com/forms/referral_trainer_request.shtml"
                    Case "PF"
                        ReDirect = "https://www.gettips.com/forms/order_feedback.shtml"
                    Case "NAM"
                        ReDirect = "https://www.gettips.com/forms/name_change.shtml"
                    Case Else
                        ReDirect = LaunchProtocol & "//w1.certegrity.com/forms.nsf/NewRequest?OpenAgent&FT=" & FORM_TYPE
                End Select
                If Debug = "Y" Then mydebuglog.Debug("ReDirect 3: " & ReDirect & vbCrLf)
            End If

            ' -----
            ' REDIRECTION TO MCM FOR ALL STUDENTS
            ' Compute Student flag
            mReDirect = ""
            If TRAINING_ACCESS = "1" Or SVC_TYPE = "PUBLIC ACCESS" Then Student = True
            If SVC_TYPE = "CERTIFICATION MANAGER ELEARNING ADMINISTRATOR" Then Student = True
            If SVC_TYPE = "CERTIFICATION MANAGER REPORTS" Then Student = False
            If SVC_TYPE = "CERTIFICATION MANAGER REG COMP" Then Student = False
            If SVC_TYPE = "CERTIFICATION MANAGER REG DB" Then Student = False
            If TRAINING_ACCESS <> "" And Val(TRAINING_ACCESS) > 1 Then Student = False
            If TRAINER_ACC_FLG = "Y" Or PAID_USER_FLG = "Y" Then Student = False
            If Debug = "Y" Then mydebuglog.Debug("Student: " & Str(Student) & vbCrLf)

            ' Process redirects
            If Student Then
                If ReDirect = Path & "/Main?OpenForm&PP=" & DOMAIN Or ReDirect = "" Then
                    If LANG_CD <> "ENU" Then
                        mReDirect = "https://www.gettips.com/mobile/" & LANG_CD & "/index.html?PP=" & DOMAIN & "&UID=" & sUserID & "&SES=" & SessionID
                    Else
                        mReDirect = "https://www.gettips.com/mobile/index.html?PP=" & DOMAIN & "&UID=" & sUserID & "&SES=" & SessionID
                    End If
                    If NO_HOME = "Y" Then mReDirect = mReDirect & "&NH=Y"
                    Mobile = True
                    GoTo ExitmCM
                End If
                If LCase(LOGIN_REDIRECT) = "invoice" Then
                    If LANG_CD <> "ENU" Then
                        mReDirect = "https://www.gettips.com/mobile/" & LANG_CD & "/index.html?UID=" & sUserID & "&SES=" & SessionID & "&OID=" & REDIRECT_REC & "&PP=" & DOMAIN
                    Else
                        mReDirect = "https://www.gettips.com/mobile/index.html?UID=" & sUserID & "&SES=" & SessionID & "&OID=" & REDIRECT_REC & "&PP=" & DOMAIN
                    End If
                    If NO_HOME = "Y" Then mReDirect = mReDirect & "&NH=Y"
                    mReDirect = mReDirect & "#ord"
                    Mobile = True
                    GoTo ExitmCM
                End If
                If LCase(LOGIN_REDIRECT) = "cardaddr" Then
                    If LANG_CD <> "ENU" Then
                        mReDirect = "https://www.gettips.com/mobile/" & LANG_CD & "/index.html?UID=" & sUserID & "&SES=" & SessionID & "&PP=" & DOMAIN
                    Else
                        mReDirect = "https://www.gettips.com/mobile/index.html?UID=" & sUserID & "&SES=" & SessionID & "&PP=" & DOMAIN
                    End If
                    If NO_HOME = "Y" Then mReDirect = mReDirect & "&NH=Y"
                    mReDirect = mReDirect & "#requpdcard"
                    Mobile = True
                    GoTo ExitmCM
                End If
                If LCase(LOGIN_REDIRECT) = "partrec" Then
                    If LANG_CD <> "ENU" Then
                        mReDirect = "https://www.gettips.com/mobile/" & LANG_CD & "/index.html?UID=" & sUserID & "&SES=" & SessionID & "&PD=" & REDIRECT_REC & "&PP=" & DOMAIN
                    Else
                        mReDirect = "https://www.gettips.com/mobile/index.html?UID=" & sUserID & "&SES=" & SessionID & "&PD=" & REDIRECT_REC & "&PP=" & DOMAIN
                    End If
                    If NO_HOME = "Y" Then mReDirect = mReDirect & "&NH=Y"
                    mReDirect = mReDirect & "#cert"
                    Mobile = True
                    GoTo ExitmCM
                End If
                If LCase(LOGIN_REDIRECT) = "openreg" Or LCase(LOGIN_REDIRECT) = "openwreg" Then
                    If LANG_CD <> "ENU" Then
                        mReDirect = "https://www.gettips.com/mobile/" & LANG_CD & "/index.html?UID=" & sUserID & "&SES=" & SessionID & "&PP=" & DOMAIN
                    Else
                        mReDirect = "https://www.gettips.com/mobile/index.html?UID=" & sUserID & "&SES=" & SessionID & "&PP=" & DOMAIN
                    End If
                    If REDIRECT_REC <> "" Then mReDirect = mReDirect & "&RG=" & REDIRECT_REC
                    If NO_HOME = "Y" Then mReDirect = mReDirect & "&NH=Y"
                    mReDirect = mReDirect & "#reg"
                    Mobile = True
                    GoTo ExitmCM
                End If
                If LCase(LOGIN_REDIRECT) = "reviewexam" Then
                    If LANG_CD <> "ENU" Then
                        mReDirect = "https://www.gettips.com/mobile/" & LANG_CD & "/index.html?UID=" & sUserID & "&SES=" & SessionID & "&PP=" & DOMAIN
                    Else
                        mReDirect = "https://www.gettips.com/mobile/index.html?UID=" & sUserID & "&SES=" & SessionID & "&PP=" & DOMAIN
                    End If
                    If REDIRECT_REC <> "" Then mReDirect = mReDirect & "&RG=" & REDIRECT_REC
                    If NO_HOME = "Y" Then mReDirect = mReDirect & "&NH=Y"
                    mReDirect = mReDirect & "#reg"
                    Mobile = True
                    GoTo ExitmCM
                End If
                If LCase(LOGIN_REDIRECT) = "orderreg" Then
                    If LANG_CD <> "ENU" Then
                        mReDirect = "https://www.gettips.com/mobile/" & LANG_CD & "/orders.html?UID=" & sUserID & "&SES=" & SessionID & "&OTY=REG&PP=" & DOMAIN
                    Else
                        mReDirect = "https://www.gettips.com/mobile/orders.html?UID=" & sUserID & "&SES=" & SessionID & "&OTY=REG&PP=" & DOMAIN
                    End If
                    If REDIRECT_REC <> "" Then mReDirect = mReDirect & "&TRN=" & REDIRECT_REC
                    If NO_HOME = "Y" Then mReDirect = mReDirect & "&NH=Y"
                    Mobile = True
                    GoTo ExitmCM
                End If
                If LCase(LOGIN_REDIRECT) = "orderpartcards" Then
                    If LANG_CD <> "ENU" Then
                        mReDirect = "https://www.gettips.com/mobile/" & LANG_CD & "/orders.html?UID=" & sUserID & "&SES=" & SessionID & "&OTY=PRT&PP=" & DOMAIN
                    Else
                        mReDirect = "https://www.gettips.com/mobile/orders.html?UID=" & sUserID & "&SES=" & SessionID & "&OTY=PRT&PP=" & DOMAIN
                    End If
                    If REDIRECT_REC <> "" Then mReDirect = mReDirect & "&TRN=" & REDIRECT_REC
                    If NO_HOME = "Y" Then mReDirect = mReDirect & "&NH=Y"
                    Mobile = True
                    GoTo ExitmCM
                End If
                If LCase(LOGIN_REDIRECT) = "orderlogbook" Then
                    If LANG_CD <> "ENU" Then
                        mReDirect = "https://www.gettips.com/mobile/" & LANG_CD & "/orders.html?UID=" & sUserID & "&SES=" & SessionID & "&OTY=AIR&PP=" & DOMAIN
                    Else
                        mReDirect = "https://www.gettips.com/mobile/orders.html?UID=" & sUserID & "&SES=" & SessionID & "&OTY=AIR&PP=" & DOMAIN
                    End If
                    If REDIRECT_REC <> "" Then mReDirect = mReDirect & "&TRN=" & REDIRECT_REC
                    If NO_HOME = "Y" Then mReDirect = mReDirect & "&NH=Y"
                    Mobile = True
                    GoTo ExitmCM
                End If
                If LCase(LOGIN_REDIRECT) = "ordermaterials" Then
                    If LANG_CD <> "ENU" Then
                        mReDirect = "https://www.gettips.com/mobile/" & LANG_CD & "/orders.html?UID=" & sUserID & "&SES=" & SessionID & "&OTY=PRS&PP=" & DOMAIN
                    Else
                        mReDirect = "https://www.gettips.com/mobile/orders.html?UID=" & sUserID & "&SES=" & SessionID & "&OTY=PRS&PP=" & DOMAIN
                    End If
                    If REDIRECT_REC <> "" Then mReDirect = mReDirect & "&TRN=" & REDIRECT_REC
                    If NO_HOME = "Y" Then mReDirect = mReDirect & "&NH=Y"
                    Mobile = True
                    GoTo ExitmCM
                End If
                If LCase(LOGIN_REDIRECT) = "orderpartcards" Then
                    If LANG_CD <> "ENU" Then
                        mReDirect = "https://www.gettips.com/mobile/" & LANG_CD & "/orders.html?UID=" & sUserID & "&SES=" & SessionID & "&OTY=PRT&PP=" & DOMAIN
                    Else
                        mReDirect = "https://www.gettips.com/mobile/orders.html?UID=" & sUserID & "&SES=" & SessionID & "&OTY=PRT&PP=" & DOMAIN
                    End If
                    If REDIRECT_REC <> "" Then mReDirect = mReDirect & "&TRN=" & REDIRECT_REC
                    If NO_HOME = "Y" Then mReDirect = mReDirect & "&NH=Y"
                    Mobile = True
                    GoTo ExitmCM
                End If
                If LCase(LOGIN_REDIRECT) = "orderpass" Then
                    If LANG_CD <> "ENU" Then
                        mReDirect = "https://www.gettips.com/mobile/" & LANG_CD & "/orders.html?UID=" & sUserID & "&SES=" & SessionID & "&OTY=PSP&PP=" & DOMAIN
                    Else
                        mReDirect = "https://www.gettips.com/mobile/orders.html?UID=" & sUserID & "&SES=" & SessionID & "&OTY=PSP&PP=" & DOMAIN
                    End If
                    If REDIRECT_REC <> "" Then mReDirect = mReDirect & "&CRS=" & REDIRECT_REC
                    If NO_HOME = "Y" Then mReDirect = mReDirect & "&NH=Y"
                    Mobile = True
                    GoTo ExitmCM
                End If
                If LCase(LOGIN_REDIRECT) = "listreg" Then
                    If LANG_CD <> "ENU" Then
                        mReDirect = "https://www.gettips.com/mobile/" & LANG_CD & "/index.html?UID=" & sUserID & "&SES=" & SessionID & "&PP=" & DOMAIN
                    Else
                        mReDirect = "https://www.gettips.com/mobile/index.html?UID=" & sUserID & "&SES=" & SessionID & "&PP=" & DOMAIN
                    End If
                    If REDIRECT_REC <> "" Then mReDirect = mReDirect & "&RG=" & REDIRECT_REC
                    If NO_HOME = "Y" Then mReDirect = mReDirect & "&NH=Y"
                    mReDirect = mReDirect & "#reg"
                    Mobile = True
                    GoTo ExitmCM
                End If
                If LCase(LOGIN_REDIRECT) = "mydocs" Then
                    If LANG_CD <> "ENU" Then
                        mReDirect = "https://www.gettips.com/mobile/" & LANG_CD & "/index.html?UID=" & sUserID & "&SES=" & SessionID & "&PP=" & DOMAIN
                    Else
                        mReDirect = "https://www.gettips.com/mobile/index.html?UID=" & sUserID & "&SES=" & SessionID & "&PP=" & DOMAIN
                    End If
                    If NO_HOME = "Y" Then mReDirect = mReDirect & "&NH=Y"
                    mReDirect = mReDirect & "#docs"
                    Mobile = True
                    GoTo ExitmCM
                End If
                If LCase(LOGIN_REDIRECT) = "accessclass" Then
                    If LANG_CD <> "ENU" Then
                        mReDirect = "https://hciscorm.certegrity.com/ls/" & LANG_CD & "/ClassAccess.html?UID=" & sUserID & "&SES=" & SessionID & "&CID=" & REDIRECT_REC & "&PP=" & DOMAIN
                    Else
                        mReDirect = "https://hciscorm.certegrity.com/ls/ClassAccess.html?UID=" & sUserID & "&SES=" & SessionID & "&CID=" & REDIRECT_REC & "&PP=" & DOMAIN
                    End If
                    If NO_HOME = "Y" Then mReDirect = mReDirect & "&NH=Y"
                    Mobile = True
                    GoTo ExitmCM
                End If
                If LCase(LOGIN_REDIRECT) = "newsurvey" Then
                    If LANG_CD <> "ENU" Then
                        mReDirect = "https://www.gettips.com/mobile/" & LANG_CD & "/OpenSurvey.html?UID=" & sUserID & "&SES=" & SessionID & "&CON=" & CON_ID & "&ID=" & REDIRECT_REC & "&PP=" & DOMAIN & "&PROT=https:"
                        If NO_HOME = "Y" Then mReDirect = mReDirect & "&NH=Y"
                        mReDirect = mReDirect & "&CUR=https://www.gettips.com/mobile/ENU/index.html"
                    Else
                        mReDirect = "https://www.gettips.com/mobile/OpenSurvey.html?UID=" & sUserID & "&SES=" & SessionID & "&CON=" & CON_ID & "&ID=" & REDIRECT_REC & "&PP=" & DOMAIN & "&PROT=https:"
                        If NO_HOME = "Y" Then mReDirect = mReDirect & "&NH=Y"
                        mReDirect = mReDirect & "&CUR=https://www.gettips.com/mobile/index.html"
                    End If
                    Mobile = True
                    GoTo ExitmCM
                End If
                If LCase(LOGIN_REDIRECT) = "allsurveylist" Then
                    If LANG_CD <> "ENU" Then
                        mReDirect = "https://www.gettips.com/mobile/" & LANG_CD & "/index.html?UID=" & sUserID & "&SES=" & SessionID & "&PP=" & DOMAIN
                    Else
                        mReDirect = "https://www.gettips.com/mobile/index.html?UID=" & sUserID & "&SES=" & SessionID & "&PP=" & DOMAIN
                    End If
                    If NO_HOME = "Y" Then mReDirect = mReDirect & "&NH=Y"
                    mReDirect = mReDirect & "#cert"
                    Mobile = True
                    GoTo ExitmCM
                End If
                If LCase(LOGIN_REDIRECT) = "surveylist" Then
                    If LANG_CD <> "ENU" Then
                        mReDirect = "https://www.gettips.com/mobile/" & LANG_CD & "/index.html?SVL=Y&UID=" & sUserID & "&SES=" & SessionID & "&PP=" & DOMAIN
                    Else
                        mReDirect = "https://www.gettips.com/mobile/index.html?SVL=Y&UID=" & sUserID & "&SES=" & SessionID & "&PP=" & DOMAIN
                    End If
                    If NO_HOME = "Y" Then mReDirect = mReDirect & "&NH=Y"
                    mReDirect = mReDirect & "#cert"
                    Mobile = True
                    GoTo ExitmCM
                End If
                If LCase(LOGIN_REDIRECT) = "questions" Then
                    If LANG_CD <> "ENU" Then
                        mReDirect = "https://www.gettips.com/mobile/" & LANG_CD & "/index.html?SVL=Y&UID=" & sUserID & "&SES=" & SessionID & "&PP=" & DOMAIN
                    Else
                        mReDirect = "https://www.gettips.com/mobile/index.html?SVL=Y&UID=" & sUserID & "&SES=" & SessionID & "&PP=" & DOMAIN
                    End If
                    If NO_HOME = "Y" Then mReDirect = mReDirect & "&NH=Y"
                    mReDirect = mReDirect & "#cert"
                    Mobile = True
                    GoTo ExitmCM
                End If
                If LCase(LOGIN_REDIRECT) = "assessment" Then
                    If LANG_CD <> "ENU" Then
                        mReDirect = "https://www.gettips.com/mobile/" & LANG_CD & "/index.html?UID=" & sUserID & "&SES=" & SessionID & "&PP=" & DOMAIN
                    Else
                        mReDirect = "https://www.gettips.com/mobile/index.html?UID=" & sUserID & "&SES=" & SessionID & "&PP=" & DOMAIN
                    End If
                    If REDIRECT_REC <> "" Then mReDirect = mReDirect & "&RG=" & REDIRECT_REC
                    If NO_HOME = "Y" Then mReDirect = mReDirect & "&NH=Y"
                    mReDirect = mReDirect & "#reg"
                    Mobile = True
                    GoTo ExitmCM
                End If
            End If

ExitmCM:
            If Debug = "Y" Then
                mydebuglog.Debug("Final ReDirect: " & ReDirect)
                mydebuglog.Debug("Final mReDirect: " & mReDirect)
            End If

            ' ==============================
            ' LOG IF APPLICABLE
            If Path <> "" Then
                If SUB_ID <> "" Then
                    mydebuglog.Debug(vbCrLf & "Logging Queries- ")
                    ' Log the user's activities in their personal record
                    SqlS = "UPDATE siebeldb.dbo.CX_SUB_CON " & _
                           "SET LAST_INST='" & Path & "', LAST_LOGIN=GETDATE(), LAST_SESS_ID='" & SessionID & "' " & _
                           "FROM (SELECT ROW_ID FROM siebeldb.dbo.S_CONTACT WHERE X_REGISTRATION_NUM='" & sUserID & "') U " & _
                           "WHERE siebeldb.dbo.CX_SUB_CON.CON_ID=U.ROW_ID "
                    temp = ExecQuery("Update", "CX_SUB_CON", cmd, SqlS, mydebuglog, Debug)

                    ' Log the user's activities
                    If sUserID <> "" Then
                        SqlS = "INSERT INTO reports.dbo.CM_LOG(REG_ID, SESSION_ID, ACTION, REMOTE_ADDR, BROWSER) " & _
                               "VALUES('" & sUserID & "','" & SessionID & "','PWSLOGIN.ASHX LOGIN', '" & REMOTE_ADDR & "','" & BROWSER & "')"
                        temp = ExecQuery("Insert", "CM_LOG", cmd, SqlS, mydebuglog, Debug)

                        SqlS = "INSERT siebeldb.dbo.CX_SUB_CON_HIST(CONFLICT_ID,CREATED_BY,LAST_UPD_BY,ROW_ID," & _
                               "SUB_CON_ID,USER_ID,SESSION_ID,REMOTE_ADDR) " & _
                               "SELECT 0,'1-3HIZ7','1-3HIZ7','" & SessionID & "', " & _
                               "SC.ROW_ID,'" & sUserID & "','" & SessionID & "','" & REMOTE_ADDR & "' " & _
                               "FROM siebeldb.dbo.S_CONTACT C " & _
                               "LEFT OUTER JOIN siebeldb.dbo.CX_SUB_CON SC ON SC.CON_ID=C.ROW_ID " & _
                               "WHERE C.X_REGISTRATION_NUM='" & sUserID & "' AND NOT EXISTS " & _
                               "(SELECT ROW_ID FROM siebeldb.dbo.CX_SUB_CON_HIST WHERE SESSION_ID='" & SessionID & "' AND USER_ID='" & sUserID & "')"
                        temp = ExecQuery("Insert", "CX_SUB_CON_HIST", cmd, SqlS, mydebuglog, Debug)
                    End If
                End If
            End If

            ' ==============================
            ' GO TO COMPUTED NEXT PAGE
            Dim EOL As String
            EOL = Chr(13) & Chr(10)
            output = output & "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 3.2//EN"">"
            output = output & "<html>"
            output = output & "<head>"

            ' Set cookies and send the user to the appropriate instance
            output = output & "<title>One moment please...</title>"
            output = output & "<link rel=""stylesheet"" href=""https://www.gettips.com/mobile/jquery.mobile-1.3.2.css"">"
            output = output & "<link rel=""stylesheet"" href=""https://www.gettips.com/mobile/css/customizations.css"">"
            output = output & "<script src=""https://www.gettips.com/mobile/jquery-1.9.1.js"" type=""text/javascript""></script>"
            output = output & "<script src=""https://www.gettips.com/mobile/jquery.mobile-1.3.2.min.js""></script>"
            output = output & "<style type=""text/css"">"""
            output = output & "a#navbutton1 a#navbutton2 {"
            output = output & "   border: 0px;"
            output = output & "   background-image: none;"
            output = output & "   background-color: #5EA5E1;"
            output = output & "   width: 55%;"
            output = output & "   color: #fff;"
            output = output & "   padding: 16px 19px;"
            output = output & "   text-transform: uppercase;"
            output = output & "   margin: 0px auto;"
            output = output & "   -webkit-border-radius: 3px;"
            output = output & "   -moz-border-radius: 3px;"
            output = output & "   border-radius: 3px;"
            output = output & "   font-family: 'Source Sans Pro', Helvetica, sans-serif;"
            output = output & "   opacity: 1;"
            output = output & "   text-indent: 0px;"
            output = output & "}"
            output = output & "body {"
            output = output & "   font-family: 'Source Sans Pro', Helvetica, sans-serif;"
            output = output & "   font-size: 17pt;"
            output = output & "   background-color: #3e4553;"
            output = output & "   background-image: none;"
            output = output & "}"
            output = output & "</style>"
            output = output & "<script src=""https://www.gettips.com/js/sessvars.js"" language=""JavaScript"" type=""text/javascript""></script> " & EOL & _
               "<!--Flash Check-->" & EOL & "<script src=""https://www.gettips.com/js/swfobject.js"" language=""JavaScript"" type=""text/javascript""></script>" & EOL & _
               "<script type=""text/javascript"">"
            output = output & "if (!swfobject.hasFlashPlayerVersion(""10.0.0"")) {" & EOL & _
               "   var expdate2 = new Date();" & EOL & _
               "   expdate2.setTime (expdate2.getTime() + 86400000);" & EOL & _
               "   SetCookie(""FLASH"",""NO"", expdate2, ""/"", "".certegrity.com"");" & EOL & _
                "}"

            output = output & "function mobileAndTabletcheck() {" & EOL & _
                "   var check = false;" & EOL & _
                "   var expdate2 = new Date();" & EOL & _
                "   expdate2.setTime (expdate2.getTime() + 86400000);" & EOL & _
                "   (function(a){if(/(android|bb\d+|meego).+mobile|avantgo|bada\/|blackberry|blazer|compal|elaine|fennec|hiptop|iemobile|ip(hone|od)|iris|kindle|lge |maemo|midp|mmp|mobile.+firefox|netfront|opera m(ob|in)i|palm( os)?|phone|p(ixi|re)\/|plucker|pocket|psp|series(4|6)0|symbian|treo|up\.(browser|link)|vodafone|wap|windows ce|xda|xiino|android|ipad|playbook|silk/i.test(a)||/1207|6310|6590|3gso|4thp|50[1-6]i|770s|802s|a wa|abac|ac(er|oo|s\-)|ai(ko|rn)|al(av|ca|co)|amoi|an(ex|ny|yw)|aptu|ar(ch|go)|as(te|us)|attw|au(di|\-m|r |s )|avan|be(ck|ll|nq)|bi(lb|rd)|bl(ac|az)|br(e|v)w|bumb|bw\-(n|u)|c55\/|capi|ccwa|cdm\-|cell|chtm|cldc|cmd\-|co(mp|nd)|craw|da(it|ll|ng)|dbte|dc\-s|devi|dica|dmob|do(c|p)o|ds(12|\-d)|el(49|ai)|em(l2|ul)|er(ic|k0)|esl8|ez([4-7]0|os|wa|ze)|fetc|fly(\-|_)|g1 u|g560|gene|gf\-5|g\-mo|go(\.w|od)|gr(ad|un)|haie|hcit|hd\-(m|p|t)|hei\-|hi(pt|ta)|hp( i|ip)|hs\-c|ht(c(\-| |_|a|g|p|s|t)|tp)|hu(aw|tc)|i\-(20|go|ma)|i230|iac( |\-|\/)|ibro|idea|ig01|ikom|im1k|inno|ipaq|iris|ja(t|v)a|jbro|jemu|jigs|kddi|keji|kgt( |\/)|klon|kpt |kwc\-|kyo(c|k)|le(no|xi)|lg( g|\/(k|l|u)|50|54|\-[a-w])|libw|lynx|m1\-w|m3ga|m50\/|ma(te|ui|xo)|mc(01|21|ca)|m\-cr|me(rc|ri)|mi(o8|oa|ts)|mmef|mo(01|02|bi|de|do|t(\-| |o|v)|zz)|mt(50|p1|v )|mwbp|mywa|n10[0-2]|n20[2-3]|n30(0|2)|n50(0|2|5)|n7(0(0|1)|10)|ne((c|m)\-|on|tf|wf|wg|wt)|nok(6|i)|nzph|o2im|op(ti|wv)|oran|owg1|p800|pan(a|d|t)|pdxg|pg(13|\-([1-8]|c))|phil|pire|pl(ay|uc)|pn\-2|po(ck|rt|se)|prox|psio|pt\-g|qa\-a|qc(07|12|21|32|60|\-[2-7]|i\-)|qtek|r380|r600|raks|rim9|ro(ve|zo)|s55\/|sa(ge|ma|mm|ms|ny|va)|sc(01|h\-|oo|p\-)|sdk\/|se(c(\-|0|1)|47|mc|nd|ri)|sgh\-|shar|sie(\-|m)|sk\-0|sl(45|id)|sm(al|ar|b3|it|t5)|so(ft|ny)|sp(01|h\-|v\-|v )|sy(01|mb)|t2(18|50)|t6(00|10|18)|ta(gt|lk)|tcl\-|tdg\-|tel(i|m)|tim\-|t\-mo|to(pl|sh)|ts(70|m\-|m3|m5)|tx\-9|up(\.b|g1|si)|utst|v400|v750|veri|vi(rg|te)|vk(40|5[0-3]|\-v)|vm40|voda|vulc|vx(52|53|60|61|70|80|81|83|85|98)|w3c(\-| )|webc|whit|wi(g |nc|nw)|wmlb|wonu|x700|yas\-|your|zeto|zte\-/i.test(a.substr(0,4)))check = true})(navigator.userAgent||navigator.vendor||window.opera);" & EOL & _
                "   if (check) { SetCookie(""MOBILE"",""YES"", expdate2, ""/"", "".certegrity.com""); }" & EOL & _
                "   return check;" & EOL & _
                "}"

            output = output & "//sessvars.$.debug();" & EOL & _
                "sessvars.$.prefs.crossDomain = true; " & EOL & _
                "function openNewWindow(fileName,theWidth,theHeight) {" & EOL & _
                "   window.open(fileName,""Details"",""toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=1,resizable=1,width=""+theWidth+"",height=""+theHeight)" & EOL & _
                "}"

            output = output & "function SetCookie (name, value) {" & EOL & _
                "   var argv = SetCookie.arguments;" & EOL & _
                "   var argc = SetCookie.arguments.length;" & EOL & _
                "   var expires = (argc > 2) ? argv[2] : null;" & EOL & _
                "   var path = (argc > 3) ? argv[3] : null;" & EOL & _
                "   var domain = (argc > 4) ? argv[4] : null;" & EOL & _
                "   var secure = (argc > 5) ? argv[5] : false;" & EOL & _
                "   document.cookie = name + ""="" + escape (value) +" & EOL & _
                "      ((path == null) ? """" : (""; path="" + path)) +" & EOL & _
                "      ((domain == null) ? """" : (""; domain="" + domain)) +" & EOL & _
                "      ((secure == true) ? ""; secure"" : """");" & EOL & _
                "}" & EOL & _
                "function logout() { " & EOL & _
                "   var lexpdate = new Date(); " & EOL & _
                "   lexpdate.setTime (lexpdate.getTime() -  86400000); " & EOL & _
                "   SetCookie(""ID"","""", lexpdate, ""/"", "".certegrity.com""); " & EOL & _
                "   SetCookie(""Sess"","""", lexpdate, ""/"", "".certegrity.com""); " & EOL & _
                "} "

            If sUserID = "" And SessionID <> "" Then SessionID = ""

            If SessID = "" Or UserID = "" Or logout_flag Then
                output = output & "function doit() {" & EOL & _
                    "   uid = '" & sUserID & "';" & EOL & _
                    "   sessid = '" & SessionID & "';" & EOL & _
                    "   var check = false;" & EOL & _
                    "   var home = '" & sReDirect & "';" & EOL & _
                    "   var student = '" & Student & "';" & EOL & _
                    "   var lredirect = '" & LOGIN_REDIRECT & "';" & EOL & _
                    "   var mredirect = '" & mReDirect & "';" & EOL & _
                    "   var redirect = '" & ReDirect & "';" & EOL & _
                    "   var mobile = " & LCase(Mobile) & ";" & EOL & _
                    "   var html5 = (typeof document.createElement('canvas').getContext === ""function"")" & EOL & _
                    "   if (sessvars.CIS_uid != null) { var uid = sessvars.CIS_uid; } " & EOL & _
                    "   if (sessvars.CIS_sessid != null) { var sessid = sessvars.CIS_sessid; } " & EOL & _
                    "   if (sessvars.CIS_uid == null) { sessvars.CIS_uid = uid; } " & EOL & _
                    "   if (sessvars.CIS_sessid == null) { sessvars.CIS_sessid = sessid; } " & EOL & _
                    "   if (sessvars.CIS_hosting == null) { sessvars.CIS_hosting = 'certegrity.com'; } " & EOL & _
                    "   if (uid == '' && sessid != '') { sessid = ''; } " & EOL & _
                    "   sessvars.CIS_cmd = ''; " & EOL & _
                    "   if (uid!='' && sessid!='') { " & EOL & _
                    "       var expdate = new Date();" & EOL & _
                    "      expdate.setTime (expdate.getTime() +  86400000);" & EOL & _
                    "      SetCookie(""ID"",uid, expdate , ""/"" , "".certegrity.com""); " & EOL & _
                    "      SetCookie(""Sess"",sessid, ""/"" , ""/"" , "".certegrity.com""); " & EOL & _
                    "      check = mobileAndTabletcheck();" & EOL & _
                    "      if (check && home=='home' && lredirect=='' && student!='True') {" & EOL & _
                    "         document.getElementById('menu').innerHTML='You are logging in from a tablet or phone.<br><br><a href=""http://getti.ps/2asFPcC"" data-role=""button"" rel=""external"" data-theme=""b"">Go to Mobile Certification Manager</a><br><br><a href=""" & ReDirect & """ data-role=""button"" rel=""external"" data-theme=""b"">Go to Desktop Certification Manager</a>';" & EOL & _
                    "         setTimeout(function() {window.location.replace('https://bit.ly/2asFPcC'); }, 15000); return; " & EOL & _
                    "         }" & EOL & _
                    "      else { " & EOL & _
                    "         document.getElementById('menu').innerHTML='<table Width=""100%"" height=""100%""><tr valign=""middle""><td align=""center"" Class=""header"">One moment please...</td></tr></table>';" & EOL & _
                    "         if (mobile && html5 && mredirect!='') {" & EOL & _
                    "            window.location.replace(mredirect); return; " & EOL & _
                    "            } else {" & EOL & _
                    "            window.location.replace(redirect); return; " & EOL & _
                    "            }" & EOL & _
                    "         }" & EOL & _
                    "      }" & EOL & _
                    "   window.location.replace('" & ReDirect & "'); " & EOL & _
                    "}"
            Else
                If Student Then
                    output = output & "function doit() {" & EOL & _
                        "   var check = false;" & EOL & _
                        "   var home = '" & sReDirect & "';" & EOL & _
                        "   var mredirect = '" & mReDirect & "';" & EOL & _
                        "   document.getElementById('menu').innerHTML='<table Width=""100%"" height=""100%""><tr valign=""middle""><td align=""center"" Class=""header"">One moment please...</td></tr></table>';" & EOL & _
                        "   window.location.replace(mredirect);" & EOL & _
                        "}"
                Else
                    output = output & "function doit() {" & EOL & _
                        "   var check = false;" & EOL & _
                        "   var home = '" & sReDirect & "';" & EOL & _
                        "   var lredirect = '" & LOGIN_REDIRECT & "';" & EOL & _
                        "   var mredirect = '" & mReDirect & "';" & EOL & _
                        "   var redirect = '" & ReDirect & "';" & EOL & _
                        "   var mobile = " & LCase(Mobile) & ";" & EOL & _
                        "   var html5 = (typeof document.createElement('canvas').getContext === ""function"")" & EOL & _
                        "   check = mobileAndTabletcheck();" & EOL & _
                        "   if (check && home=='home' && lredirect=='' && !mobile) {" & EOL & _
                        "      document.getElementById('menu').innerHTML='You are logging in from a tablet or phone.<br><br><a href=""http://getti.ps/2asFPcC"" data-role=""button"" rel=""external"" data-theme=""b"">Go to Mobile Certification Manager</a><br><br><a href=""" & ReDirect & """ data-role=""button"" rel=""external"" data-theme=""b"">Go to Desktop Certification Manager</a>';" & EOL & _
                        "      setTimeout(function() {window.location.replace('https://bit.ly/2asFPcC'); }, 15000); " & EOL & _
                        "      }" & EOL & _
                        "   else { " & EOL & _
                        "      document.getElementById('menu').innerHTML='<table Width=""100%"" height=""100%""><tr valign=""middle""><td align=""center"" Class=""header"">One moment please...</td></tr></table>';" & EOL & _
                        "      if (mobile && html5 && mredirect!='') {"
                    output = output & "         window.location.replace(mredirect);"
                    output = output & "      } else {"
                    output = output & "         window.location.replace(redirect); "
                    output = output & "      }"
                    output = output & "   }" & EOL & _
                    "}"
                End If

            End If
            output = output & "// -- End Hiding Here -->"
            output = output & " </script>"
            output = output & "<!--[if lt IE 9]>" & EOL & _
               "<script type=""text/javascript"">var expdate = new Date(); expdate.setTime (expdate.getTime() + 86400000); SetCookie(""HTML5"",""NO"", expdate, ""/"", "".certegrity.com"")</script>" & EOL & _
                "<!--<![endif]-->"
            output = output & "</head>"

            ' Standard Header
            output = output & "<link href=""https://www.gettips.com/css/" & Trim(LCase(DOMAIN)) & ".css"" rel=""stylesheet"" type=""text/css"">"
            output = output & "<style type=""text/css""><!-- body {  margin-top: 0px; margin-right: 0px; margin-bottom: 0px; margin-left: 0px} --></style>"
            output = output & "<body>"
            output = output & "<div data-role=""page"" id=""reg"" data-theme=""d""><div class=""c-interior"">"
            output = output & "<div data-role=""content"">"
            output = output & "<center>"
            output = output & "<div id=""menu""></div>"
            output = output & "</center>"
            output = output & "</div>"
            output = output & "</div></div>"
            output = output & "<script language=""JavaScript"">"
            If logout_flag Then output = output & "logout();"
            output = output & "doit();"
            output = output & "</script>"
            output = output & "</html>"
            GoTo CloseOut
        Else
            GoTo SystemUnavailable
        End If
        GoTo CloseOut

SystemUnavailable:
        If Debug = "Y" Then mydebuglog.Debug(">>SystemUnavailable")
        errmsg = "System Unavailable"
        output = output & "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 3.2//EN"">"
        output = output & "<HTML>"
        output = output & "<HEAD>"
        output = output & "<META HTTP-EQUIV=""Refresh"" CONTENT=""0; URL=https://www.gettips.com/unavailable.shtml"">"
        output = output & "<BODY BGCOLOR='White' leftmargin=0 text='#000040' link='Purple' vlink='Navy'>"
        GoTo CloseOut

SubscriptionError:
        If Debug = "Y" Then mydebuglog.Debug(">>SubscriptionError")
        errmsg = "Subscription Error"
        ErrLvl = "Warning"
        output = output & "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 3.2//EN"">"
        output = output & "<HTML>"
        output = output & "<HEAD>"
        output = output & "<META HTTP-EQUIV=""Refresh"" CONTENT=""0; URL=https://www.gettips.com/accesserror.shtml"">"
        output = output & "<script language=""JavaScript"">"
        output = output & "function dynamicLogout() {"
        output = output & "   var hosting = baseDomainString();"
        output = output & "   DeleteCookie(""ID"",""/"",hosting);"
        output = output & "   DeleteCookie(""Sess"",""/"",hosting);"
        output = output & "   sessvars.$.clearMem()"
        output = output & "   sessvars.$.flush()"
        output = output & "   sessvars.CIS_cmd = 'logout';"
        output = output & "}"
        output = output & "function DeleteCookie( name, path, domain ) {"
        output = output & "    document.cookie = name + ""="" + ( ( path ) ? "";path="" + path : """") + ( ( domain ) ? "";domain="" + domain : """" ) + "";expires=Thu, 01-Jan-1970 00:00:01 GMT"";"
        output = output & "}"
        output = output & "function baseDomainString(){"
        output = output & "     e = document.domain.split(/\./);"
        output = output & "     if(e.length > 1) {"
        output = output & "       return(e[e.length-2] + ""."" +  e[e.length-1]);"
        output = output & "     }else{"
        output = output & "       return("""");"
        output = output & "     }"
        output = output & "}"
        output = output & "</script>"
        output = output & "<body onload=""dynamicLogout()"">"
        GoTo CloseOut

AccessError:
        ErrLvl = "Warning"
        If Debug = "Y" Then mydebuglog.Debug(">>AccessError")
        output = output & "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 3.2//EN"">"
        output = output & "<HTML>"
        output = output & "<HEAD>"
        output = output & "<META HTTP-EQUIV=""Refresh"" CONTENT=""0; URL=https://www.gettips.com/accesserror.html"">"
        output = output & "<BODY BGCOLOR='White' leftmargin=0 text='#000040' link='Purple' vlink='Navy'>"

CloseOut:
        ' ============================================
        ' Close database connections and objects
        If DBisOpen Then
            Try
                dr = Nothing
                con.Dispose()
                con = Nothing
                cmd.Dispose()
                cmd = Nothing
            Catch ex As Exception
                errmsg = errmsg & "Unable to close the database connection. " & vbCrLf
            End Try
        End If

CloseOut2:
        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("WsLogin.ashx : " & ErrLvl & ": " & Trim(errmsg))
        myeventlog.Info("WsLogin.ashx : sUserID: " & sUserID & ", SessionID: " & SessionID & ", ReDirect: " & ReDirect)
        If Debug = "Y" Or (Logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug(vbCrLf & "Results:  sUserID: " & sUserID & ", SessionID: " & SessionID & ", ReDirect: " & ReDirect)
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
                LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, "WsLogin", LogStartTime, VersionNum, Debug)
            Catch ex As Exception
            End Try
        End If

        ' Send results        
        context.Response.ContentType = "text/html"
        context.Response.Write(output)
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
    Public Function SimpleString(sRawURL As String) As String
        ' Removes all but alphanumeric from a string
        Dim iLoop As Integer
        Dim sRtn As String = ""
        Dim sTmp As String = ""
        Const sValidChars = "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        If Len(sRawURL) > 0 Then
            ' Loop through each char
            For iLoop = 1 To Len(sRawURL)
                sTmp = Mid(sRawURL, iLoop, 1)

                If InStr(1, sValidChars, sTmp, vbBinaryCompare) = 0 Then
                    ' If not ValidChar, then remove
                Else
                    sRtn = sRtn & sTmp
                End If

            Next iLoop
        End If
        SimpleString = sRtn
    End Function

    Public Function SimpleString2(sRawURL As String) As String
        ' Removes all but alphanumeric from a string
        Dim iLoop As Integer
        Dim sRtn As String = ""
        Dim sTmp As String = ""
        Const sValidChars = "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz:/&?=%<>+-._"
        If Len(sRawURL) > 0 Then
            ' Loop through each char
            For iLoop = 1 To Len(sRawURL)
                sTmp = Mid(sRawURL, iLoop, 1)

                If InStr(1, sValidChars, sTmp, vbBinaryCompare) = 0 Then
                    ' If not ValidChar, then remove
                Else
                    sRtn = sRtn & sTmp
                End If

            Next iLoop
        End If
        SimpleString2 = sRtn
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