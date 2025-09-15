<%@ WebHandler Language="VB" Class="WSCLogin" %>

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

Public Class WSCLogin : Implements IHttpHandler
    
    Enum enumObjectType
        StrType = 0
        IntType = 1
        DblType = 2
        DteType = 3
    End Enum
    
    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        
        ' Parameter Declarations
        Dim Debug, CRSE_TSTRUN_ID, RETURN_PAGE, DOMAIN, RELOGIN As String
        Dim REG_NUM, SessionID, CURRENT_PAGE, HOME_PAGE, LANG_CD, callback, myprotocol As String
        
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
        mydebuglog = log4net.LogManager.GetLogger("CLoginDebugLog")
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
        
        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service
        Dim Processing As New com.certegrity.cloudsvc.processing.Service
        Dim Cmp As New com.certegrity.hciscormsvc.cmp.CMProfiles
        
        ' Variable declarations
        Dim errmsg, temp As String                    ' Error message (if any)
        Dim EOL, LaunchProtocol, ReturnDestination, NextLink, outdata, ErrLvl As String
        Dim REDIRECT, PORTALLOGIN, LOGOUT, USER, PART_ID, TRAINER_FLG, REMOTE_ADDR_PARAM As String
        Dim LOGGED_IN, CON_ID, SUB_ID, CONTACT_OU_ID, EMAIL_ADDR, oform As String
        Dim FACEBOOK_ID, FACEBOOK_EMAIL, GOOGLE_ID, STORED_FB_ID, basepath As String
        Dim TERM_FLG, SUB_CON_ID, SVC_TYPE, HOME_URL, UNSUB_URL, LOGOUT_URL, ETIPS_DOMAIN, SRC_URL As String
        Dim RecacheProfile As Boolean
        Dim Recheck, LastLogin, LastLogout As String
        
        ' ============================================
        ' Variable setup
        Debug = "Y"
        Logging = "Y"
        REG_NUM = ""
        RELOGIN = ""
        SessionID = ""
        LOGGED_IN = "N"
        CON_ID = ""
        SUB_ID = ""
        CONTACT_OU_ID = ""
        DOMAIN = ""
        LANG_CD = "ENU"
        EMAIL_ADDR = ""
        REDIRECT = ""
        PORTALLOGIN = ""
        LOGOUT = ""
        USER = ""
        PART_ID = ""
        TRAINER_FLG = ""
        callback = ""
        myprotocol = ""
        HOME_PAGE = ""
        RETURN_PAGE = ""
        CURRENT_PAGE = ""
        CRSE_TSTRUN_ID = ""
        errmsg = ""
        EOL = Chr(10)
        LaunchProtocol = "http:"
        ReturnDestination = ""
        NextLink = ""
        outdata = ""
        ErrLvl = "Error"
        oform = "JSON"
        STORED_FB_ID = ""        
        temp = ""
        RecacheProfile = False
        Recheck = ""
        LastLogin = ""
        LastLogout = ""
        FACEBOOK_EMAIL = ""
        FACEBOOK_ID = ""
        GOOGLE_ID = ""
        TERM_FLG = ""
        HOME_URL = ""
        SRC_URL = ""
        SUB_CON_ID = ""
        ETIPS_DOMAIN = ""
        SVC_TYPE = "PUBLIC ACCESS"
        LOGOUT_URL = ""
        REMOTE_ADDR_PARAM = ""
                
        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
            tempdebug = System.Configuration.ConfigurationManager.AppSettings.Get("CLogin_debug")
            If tempdebug = "Y" And Debug <> "T" Then Debug = "Y"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("LaunchProtocol")
            If temp <> "" Then LaunchProtocol = temp
            basepath = LaunchProtocol & "//w2.certegrity.com/cp0.nsf"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            GoTo CloseOut2
        End Try
        
        ' ============================================
        ' Open log file if applicable
        If Logging = "Y" Then
            logfile = "C:\Logs\WSCLogin.log"
            Try
                log4net.GlobalContext.Properties("CLoginLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                GoTo CloseOut2
            End Try
        End If
        
        ' ============================================
        ' Get parameters    
        If Not context.Request.QueryString("UID") Is Nothing Then
            REG_NUM = context.Request.QueryString("UID")
        End If
        
        If Not context.Request.QueryString("SES") Is Nothing Then
            SessionID = context.Request.QueryString("SES")
        End If
 
        If Not context.Request.QueryString("HP") Is Nothing Then
            HOME_PAGE = context.Request.QueryString("HP")
        End If

        If Not context.Request.QueryString("RL") Is Nothing Then
            RELOGIN = context.Request.QueryString("RL")
        End If
        
        If Not context.Request.QueryString("callback") Is Nothing Then
            callback = context.Request.QueryString("callback")
        End If
        
        If Not context.Request.QueryString("FM") Is Nothing Then
            If context.Request.QueryString("FM") = "X" Then oform = "XML"
            If context.Request.QueryString("FM") = "M" Then oform = "MOBILE"
        End If
        
        If Not context.Request.QueryString("LANG") Is Nothing Then
            LANG_CD = UCase(context.Request.QueryString("LANG"))
        End If
        
        If Not context.Request.QueryString("FB") Is Nothing Then
            FACEBOOK_ID = context.Request.QueryString("FB")
        End If
 
        If Not context.Request.QueryString("FE") Is Nothing Then
            FACEBOOK_EMAIL = context.Request.QueryString("FE")
        End If
 
        If Not context.Request.QueryString("GG") Is Nothing Then
            GOOGLE_ID = context.Request.QueryString("GG")
        End If
        
        If Not context.Request.QueryString("RA") Is Nothing Then
            GOOGLE_ID = context.Request.QueryString("RA")
        End If
        
        ' Validate parameters
        If REG_NUM = "null" Or REG_NUM = "undefined" Then REG_NUM = ""
        If SessionID = "null" Or SessionID = "undefined" Then SessionID = ""
        If LANG_CD <> "ESN" And LANG_CD <> "ENU" Then LANG_CD = "ENU"
        If callback = "" Then callback = "?"
        If myprotocol = "" Then myprotocol = "http:"
        If REMOTE_ADDR_PARAM <> "" Then callip = REMOTE_ADDR_PARAM
        
        If Debug = "Y" Then
            mydebuglog.Debug("----------------------------------")
            mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
            mydebuglog.Debug("Parameters-")
            mydebuglog.Debug("  callip: " & callip)
            mydebuglog.Debug("  LaunchProtocol: " & LaunchProtocol)
            mydebuglog.Debug("  SessionID: " & SessionID)
            mydebuglog.Debug("  REG_NUM: " & REG_NUM)
            mydebuglog.Debug("  RELOGIN: " & RELOGIN)
            mydebuglog.Debug("  FACEBOOK_ID: " & FACEBOOK_ID)
            mydebuglog.Debug("  FACEBOOK_EMAIL: " & FACEBOOK_EMAIL)
            mydebuglog.Debug("  GOOGLE_ID: " & GOOGLE_ID)
            mydebuglog.Debug("  REMOTE_ADDR_PARAM: " & REMOTE_ADDR_PARAM)            
            mydebuglog.Debug("  HOME_PAGE: " & HOME_PAGE)
            mydebuglog.Debug("  LANG_CD: " & LANG_CD)
            mydebuglog.Debug("  DOMAIN: " & DOMAIN)
            mydebuglog.Debug("  BROWSER: " & BROWSER)
            mydebuglog.Debug("  oform: " & oform)
            mydebuglog.Debug("  callback: " & callback)
        End If
        
        ' ============================================
        ' Open database connection 
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            errmsg = errmsg & "Unable to open the database connection. " & vbCrLf
            GoTo CloseOut
        End If

        ' ============================================
        ' Process Request
        If Not cmd Is Nothing Then
            
            ' ================================================
            ' IF FACEBOOK ID PROVIDED THEN CHECK TO SEE IF THIS EXISTS AS A USER   
            If FACEBOOK_ID <> "" And REG_NUM = "" Then
                SqlS = "SELECT TOP 1 C.X_REGISTRATION_NUM, C.ROW_ID, C.X_PR_LANG_CD, X.ATTRIB_35 " & _
                    "FROM siebeldb.dbo.S_CONTACT C " & _
                    "LEFT OUTER JOIN siebeldb.dbo.CX_SESS_REG R ON R.CONTACT_ID=C.ROW_ID " & _
                    "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT_X X ON X.PAR_ROW_ID=C.ROW_ID " & _
                    "WHERE X.ATTRIB_35='" & SqlString(FACEBOOK_ID) & "' " & _
                    "ORDER BY C.X_TRAINER_NUM DESC, C.X_PART_ID DESC, R.ROW_ID DESC"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Verify user: " & vbCrLf & "  " & SqlS)
                Try
                    cmd.CommandText = SqlS
                    dr = cmd.ExecuteReader()
                    If Not dr Is Nothing Then
                        While dr.Read()
                            REG_NUM = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            CON_ID = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            LANG_CD = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                            STORED_FB_ID = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                        End While
                    End If
                Catch ex As Exception
                    GoTo AccessError
                End Try
                dr.Close()
                
                If Debug = "Y" Then
                    mydebuglog.Debug("  .. REG_NUM: " & REG_NUM)
                    mydebuglog.Debug("  .. CON_ID: " & CON_ID)
                    mydebuglog.Debug("  .. LANG_CD: " & LANG_CD)
                    mydebuglog.Debug("  .. STORED_FB_ID: " & STORED_FB_ID & vbCrLf)
                End If
            End If
            If REG_NUM = "" Then GoTo CloseOut
            
            ' ================================================
            ' LOOKUP SUBSCRIPTION
            If REG_NUM <> "" Then
                SqlS = "SELECT TOP 1 C.SUB_ID, (SELECT CASE WHEN (S.SVC_TERM_DT<GETDATE() OR (C.USER_EXP_DATE<GETDATE() AND C.USER_EXP_DATE IS NOT NULL)) AND S.SVC_TYPE<>'PUBLIC ACCESS' THEN 'Y' ELSE 'N' END) AS TERM_FLG, " & _
                    "C.ROW_ID AS SUB_CON_ID, S.SVC_TYPE, S.DOMAIN, P.FST_NAME+' '+P.LAST_NAME, P.EMAIL_ADDR, P.X_TRAINER_FLG, P.ROW_ID, " & _
                    "(SELECT CASE WHEN P.X_PR_LANG_CD IS NULL OR P.X_PR_LANG_CD='' THEN 'ENU' ELSE P.X_PR_LANG_CD END), P.X_PART_ID, X.ATTRIB_35, " & _
                    "D.HOME_URL, D.DEF_SUB_ID, D.UNSUB_URL, D.LOGOUT_URL, D.ETIPS_FLG, D.SRC_URL " & _
                    "FROM siebeldb.dbo.S_CONTACT P " & _
                    "LEFT OUTER JOIN siebeldb.dbo.CX_SUB_CON C ON C.CON_ID=P.ROW_ID " & _
                    "LEFT OUTER JOIN siebeldb.dbo.CX_SUBSCRIPTION S ON S.ROW_ID=C.SUB_ID " & _
                    "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT_X X ON X.PAR_ROW_ID=P.ROW_ID " & _
                    "LEFT OUTER JOIN siebeldb.dbo.CX_SUB_DOMAIN D ON D.DOMAIN=S.DOMAIN " & _
                    "WHERE P.X_REGISTRATION_NUM='" & REG_NUM & "'"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Subscription Query: " & vbCrLf & "  " & SqlS)
                Try
                    cmd.CommandText = SqlS
                    dr = cmd.ExecuteReader()
                    If Not dr Is Nothing Then
                        While dr.Read()
                            SUB_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            TERM_FLG = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            SUB_CON_ID = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                            SVC_TYPE = UCase(Trim(CheckDBNull(dr(3), enumObjectType.StrType)))
                            If SVC_TYPE = "PUBLIC ACCESS" Then TERM_FLG = "N"
                            If DOMAIN = "" Then DOMAIN = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                            If DOMAIN = "" Then DOMAIN = "TIPS"
                            USER = Trim(CheckDBNull(dr(5), enumObjectType.StrType))
                            EMAIL_ADDR = Trim(CheckDBNull(dr(6), enumObjectType.StrType))
                            TRAINER_FLG = Trim(CheckDBNull(dr(7), enumObjectType.StrType))
                            CON_ID = Trim(CheckDBNull(dr(8), enumObjectType.StrType))
                            LANG_CD = Trim(CheckDBNull(dr(9), enumObjectType.StrType))
                            PART_ID = Trim(CheckDBNull(dr(10), enumObjectType.StrType))
                            STORED_FB_ID = Trim(CheckDBNull(dr(11), enumObjectType.StrType))
                            HOME_URL = Trim(CheckDBNull(dr(12), enumObjectType.StrType))
                            SUB_ID = Trim(CheckDBNull(dr(13), enumObjectType.StrType))
                            UNSUB_URL = Trim(CheckDBNull(dr(14), enumObjectType.StrType))
                            LOGOUT_URL = Trim(CheckDBNull(dr(15), enumObjectType.StrType))
                            ETIPS_DOMAIN = Trim(CheckDBNull(dr(16), enumObjectType.StrType))
                            SRC_URL = Trim(CheckDBNull(dr(17), enumObjectType.StrType))
                        End While
                    End If
                Catch ex As Exception
                    GoTo AccessError
                End Try
                dr.Close()
                If DOMAIN = "" Then DOMAIN = "CSI"
                If Debug = "Y" Then
                    mydebuglog.Debug("  .. SUB_ID: " & SUB_ID)
                    mydebuglog.Debug("  .. CON_ID: " & CON_ID)
                    mydebuglog.Debug("  .. PART_ID: " & PART_ID)
                    mydebuglog.Debug("  .. DOMAIN: " & DOMAIN)
                    mydebuglog.Debug("  .. TERM_FLG: " & TERM_FLG)
                    mydebuglog.Debug("  .. USER: " & USER)
                    mydebuglog.Debug("  .. EMAIL_ADDR: " & EMAIL_ADDR)
                    mydebuglog.Debug("  .. LANG_CD: " & LANG_CD)
                    mydebuglog.Debug("  .. STORED_FB_ID: " & STORED_FB_ID)
                    mydebuglog.Debug("  .. HOME_URL: " & HOME_URL)
                    mydebuglog.Debug("  .. SRC_URL: " & SRC_URL)
                    mydebuglog.Debug("  .. SUB_CON_ID: " & SUB_CON_ID)
                    mydebuglog.Debug("  .. ETIPS_DOMAIN: " & ETIPS_DOMAIN & vbCrLf)
                End If
                If SUB_CON_ID = "" Or SUB_ID = "" Then GoTo AccessError
                
                If CON_ID = "" Then
                    Dim pdoc As XmlDocument
                    pdoc = Cmp.GetUserProfile(REG_NUM, SessionID, Debug)
                    If pdoc Is Nothing Then
                    Else
                        Dim oNodeList As XmlNodeList = pdoc.SelectNodes("//profile")
                        For i = 0 To oNodeList.Count - 1
                            CON_ID = GetNodeValue("CONTACT_ID ", oNodeList.Item(i))
                            CONTACT_OU_ID = GetNodeValue("CONTACT_OU_ID", oNodeList.Item(i))
                            SUB_ID = GetNodeValue("SUB_ID", oNodeList.Item(i))
                            SUB_CON_ID = GetNodeValue("SUB_CON_ID", oNodeList.Item(i))
                            DOMAIN = GetNodeValue("DOMAIN", oNodeList.Item(i))
                            LOGGED_IN = GetNodeValue("LOGGED_IN", oNodeList.Item(i))
                        Next
                        If Debug = "Y" Then
                            mydebuglog.Debug("  GetUserProfile-")
                            mydebuglog.Debug("  .. LOGGED_IN: " & LOGGED_IN)
                            mydebuglog.Debug("  .. CON_ID: " & CON_ID)
                            mydebuglog.Debug("  .. SUB_CON_ID: " & SUB_CON_ID)
                            mydebuglog.Debug("  .. CONTACT_OU_ID: " & CONTACT_OU_ID)
                            mydebuglog.Debug("  .. DOMAIN: " & DOMAIN & vbCrLf)
                        End If
                    End If
                End If
            End If
            
            ' ================================================
            ' IF FACEBOOK LOGGED IN AND NOT OTHERWISE LOGGED IN, THEN LOGIN USER   
            If REG_NUM <> "" And SUB_CON_ID <> "" Then
                If SessionID = "" Then
                    Randomize()
                    SessionID = "S" & Trim(Str(Day(Now))) & Trim(Str(Hour(Now))) & Trim(Str(Minute(Now))) & Trim(Str(Second(Now))) & Chr(Str(Int(Rnd() * 26)) + 65) & Chr(Str(Int(Rnd() * 26)) + 65) & Chr(Str(Int(Rnd() * 26)) + 65) & Chr(Str(Int(Rnd() * 26)) + 65)
                    If Debug = "Y" Then mydebuglog.Debug("  New SessionID: " & SessionID)
                    
                    ' Log the user's activities   
                    SqlS = "UPDATE siebeldb.dbo.CX_SUB_CON " & _
                        "SET LAST_INST='" & basepath & "', LAST_LOGIN=GETDATE(), LAST_SESS_ID='" & SessionID & "' " & _
                        "WHERE ROW_ID='" & SUB_CON_ID & ""
                    temp = ExecQuery("Update", "CX_SUB_CON", cmd, SqlS, mydebuglog, Debug)
                    
                    SqlS = "INSERT siebeldb.dbo.CX_SUB_CON_HIST(CONFLICT_ID,CREATED_BY,LAST_UPD_BY,ROW_ID," & _
                        "SUB_CON_ID,USER_ID,SESSION_ID,REMOTE_ADDR) " & _
                        "VALUES(0,'1-3HIZ7','1-3HIZ7','" & SessionID & "', " & _
                        "'" & SUB_CON_ID & "','" & REG_NUM & "','" & SessionID & "','" & callip & "')"
                    temp = ExecQuery("Insert", "CX_SUB_CON_HIST", cmd, SqlS, mydebuglog, Debug)
                    
                    SqlS = "INSERT INTO reports.dbo.CM_LOG(REG_ID, SESSION_ID, ACTION, BROWSER, REMOTE_ADDR) " & _
                        "VALUES('" & REG_NUM & "','" & SessionID & "','WsCLogin.ashx LOGIN','" & BROWSER & "','" & callip & "')"
                    temp = ExecQuery("Insert", "CM_LOG", cmd, SqlS, mydebuglog, Debug)
                    
                    ' Create the profile
                    If SVC_TYPE <> "PUBLIC ACCESS" Then
                        Dim sdoc As XmlDocument
                        sdoc = Cmp.SetUserProfile(REG_NUM, SessionID, "N", Debug)
                    End If
                Else
                    ' -----
                    ' LOOKUP LOGIN
                    ' See if a login has been logged
                    Dim LOG_COUNT As Integer
                    SqlS = "SELECT COUNT(*) " & _
                        "FROM reports.dbo.CM_LOG " & _
                        "WHERE REG_ID='" & REG_NUM & "' AND SESSION_ID='" & SessionID & "'"
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Check access query: " & vbCrLf & "  " & SqlS)
                    Try
                        cmd.CommandText = SqlS
                        dr = cmd.ExecuteReader()
                        If Not dr Is Nothing Then
                            While dr.Read()
                                LOG_COUNT = CheckDBNull(dr(0), enumObjectType.IntType)
                            End While
                        Else
                            GoTo AccessError
                        End If
                    Catch ex As Exception
                    End Try
                    If Debug = "Y" Then mydebuglog.Debug("  .. LOG_COUNT: " & Str(LOG_COUNT) & vbCrLf)
                    dr.Close()
                    If LOG_COUNT = 0 Then GoTo CannotValidate
                End If
            End If

            ' ================================================
            ' RECACHE INFORMATION
            If oform = "MOBILE" Then
                SqlS = "SELECT TOP 1 (SELECT CASE WHEN SC.LAST_LOGOUT IS NULL AND DATEDIFF(DY,LAST_LOGIN,GETDATE())<=1 THEN 'N' ELSE 'Y' END), LAST_LOGIN, LAST_LOGOUT " & _
                    "FROM reports.dbo.CM_LOG L " & _
                    "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.X_REGISTRATION_NUM=L.REG_ID " & _
                    "LEFT OUTER JOIN siebeldb.dbo.CX_SUB_CON SC ON SC.CON_ID=C.ROW_ID " & _
                    "WHERE L.REG_ID='" & REG_NUM & "' AND L.SESSION_ID='" & SessionID & "' " & _
                    "ORDER BY LAST_LOGIN DESC"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Verify user: " & vbCrLf & "  " & SqlS)
                Try
                    cmd.CommandText = SqlS
                    dr = cmd.ExecuteReader()
                    If Not dr Is Nothing Then
                        While dr.Read()
                            Recheck = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            LastLogin = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            LastLogout = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                        End While
                    Else
                        GoTo AccessError
                    End If
                Catch ex As Exception
                    GoTo AccessError
                End Try
                dr.Close()
                If Debug = "Y" Then
                    mydebuglog.Debug("  .. Recheck: " & Recheck)
                    mydebuglog.Debug("  .. LastLogin: " & LastLogin)
                    mydebuglog.Debug("  .. LastLogout: " & LastLogout & vbCrLf)
                End If
                If LastLogin = "" Then GoTo CannotValidate
                If Recheck = "Y" Then RecacheProfile = True Else RecacheProfile = False
            
                If RecacheProfile Then
                    SqlS = "UPDATE siebeldb.dbo.CX_SUB_CON SET LAST_LOGOUT=NULL WHERE ROW_ID='" & SUB_CON_ID & "'"
                    temp = ExecQuery("Update", "CX_SUB_CON", cmd, SqlS, mydebuglog, Debug)
                End If
            End If
            
            ' ================================================
            ' UPDATE FACEBOOK ID IF APPLICABLE
            If CON_ID <> "" And STORED_FB_ID = "" And FACEBOOK_ID <> "" Then
                SqlS = "UPDATE siebeldb.dbo.S_CONTACT_X " & _
                    "SET ATTRIB_35='" & SqlString(FACEBOOK_ID) & "', ATTRIB_37='" & SqlString(FACEBOOK_EMAIL) & "' " & _
                    "WHERE PAR_ROW_ID='" & SqlString(CON_ID) & "'"
                temp = ExecQuery("Update", "S_CONTACT_X", cmd, SqlS, mydebuglog, Debug)
            End If
            
            ' ================================================
            ' GENERATE REDIRECT
            ' The required support pages are as follows:
            '   expired.html   -   The user's subscription is expired, they must contact our sales staff
            If TERM_FLG = "Y" Then
                If oform = "MOBILE" Then
                    REDIRECT = "#expired"
                Else
                    If LANG_CD = "ESN" Then
                        REDIRECT = "/ESN/expired.shtml"
                    Else
                        REDIRECT = "/expired.shtml"
                    End If
                End If
                PORTALLOGIN = ""
                Select Case LANG_CD
                    Case "ESN"
                        errmsg = "Su suscripci&oacute;n ha expirado. Por favor, p&oacute;ngase en contacto con nuestro departamento de ventas para m&aacute;s informaci&oacute;n.."
                    Case Else
                        errmsg = "Your subscription has expired.  Please contact our Sales department for more information."
                End Select
                GoTo CloseOut
            Else
                If oform = "MOBILE" Then
                    REDIRECT = "#mainp"
                    PORTALLOGIN = "https://hciscorm.certegrity.com/ls/WsLogin.ashx?ID=" & REG_NUM & "&SESS=" & SessionID & "&REF=" & HOME_PAGE & "&RD=home&DOM=" & DOMAIN
                Else
                    If UCase(DOMAIN) = "TIPS" Then
                        Select Case HOME_PAGE
                            Case "gettips.com"
                                REDIRECT = "https://hciscorm.certegrity.com/ls/WsLogin.ashx?ID=" & REG_NUM & "&SESS=" & SessionID & "&REF=" & HOME_PAGE & "&RD=home&LANG=" & LANG_CD & "&DOM=" & DOMAIN
                            Case Else
                                REDIRECT = "https://hciscorm.certegrity.com/ls/WsLogin.ashx?ID=" & REG_NUM & "&SESS=" & SessionID & "&REF=" & HOME_PAGE & "&RD=home&LANG=" & LANG_CD & "&DOM=" & DOMAIN
                        End Select
                    Else
                        REDIRECT = "https://hciscorm.certegrity.com/ls/WsLogin.ashx?ID=" & REG_NUM & "&SESS=" & SessionID & "&REF=" & HOME_PAGE & "&RD=home&LANG=" & LANG_CD & "&DOM=" & DOMAIN
                    End If
                End If
            End If
            If Debug = "Y" Then mydebuglog.Debug("  REDIRECT: " & REDIRECT & vbCrLf)
            
            ' ================================================
            ' COMPUTE LOGOUT
            If oform = "MOBILE" Then
                LOGOUT = "#logout"
            Else
                LOGOUT = "https://hciscorm.certegrity.com/ls/bcplogout.ashx?LANG=" & LANG_CD & "&RD=" & LOGOUT_URL
            End If
            
        Else
            errmsg = errmsg & "Unable to open the database connection. " & vbCrLf
        End If
        
        ' ================================================
        ' RETURN TO USER
ReturnControl:
        GoTo CloseOut
        
CannotValidate:
        Select Case LANG_CD
            Case "ESN"
                errmsg = "Las credenciales de acceso no se pueden validar. Por favor, iniciar sesi\u00D3n otra vez."
            Case Else
                errmsg = "The access credentials cannot be validated.  Please logout and login again."
        End Select
        If oform = "MOBILE" Then
            LOGOUT = "#logout"
        Else
            LOGOUT = GetLogout(DOMAIN, LANG_CD, "https://w1.certegrity.com/plogini.nsf/")
        End If
        ErrLvl = "Warning"
        REDIRECT = ""
        GoTo CloseOut
   
DBError:
        If Debug = "Y" Then mydebuglog.Debug(">>DBError")
        Select Case LANG_CD
            Case "ESN"
                errmsg = "El sistema puede no estar disponible ahora. Por favor, int&eacute;ntelo de nuevo m&aacute;s tarde"
            Case Else
                errmsg = "The system may be unavailable now.  Please try again later"
        End Select
        GoTo CloseOut
   
AccessError:
        If Debug = "Y" Then mydebuglog.Debug(">>AccessError")
        ErrLvl = "Warning"
        Select Case LANG_CD
            Case "ESN"
                errmsg = "Las credenciales de acceso proporcionadas son incorrectas. Por favor, int\u00E9ntelo de nuevo."
            Case Else
                errmsg = "The access credentials provided are incorrect.  Please try again."
        End Select
          
CloseOut:
        If Debug = "Y" Then
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & ">>Final Values")
            mydebuglog.Debug("  .. Id: " & REG_NUM)
            mydebuglog.Debug("  .. Domain: " & DOMAIN)
            mydebuglog.Debug("  .. Redirect: " & REDIRECT)
            mydebuglog.Debug("  .. Portal: " & PORTALLOGIN)
            mydebuglog.Debug("  .. ErrMsg: " & errmsg)
            mydebuglog.Debug("  .. SessId: " & SessionID)
            mydebuglog.Debug("  .. LogOut: " & LOGOUT)
            mydebuglog.Debug("  .. Username: " & USER)
            mydebuglog.Debug("  .. ContactId: " & CON_ID)
            mydebuglog.Debug("  .. LangCd: " & LANG_CD)
            mydebuglog.Debug("  .. EmailAddr: " & EMAIL_ADDR)
            mydebuglog.Debug("  .. ParticipantId: " & PART_ID)
            mydebuglog.Debug("  .. TrainerFlag: " & TRAINER_FLG)
        End If

        ' ============================================
        ' Finalize output      
        outdata = ""
        outdata = outdata & """Id"":""" & REG_NUM & ""","
        outdata = outdata & """Domain"":""" & EscapeJSON(DOMAIN) & ""","
        outdata = outdata & """Redirect"":""" & EscapeJSON(REDIRECT) & ""","
        outdata = outdata & """Portal"":""" & EscapeJSON(PORTALLOGIN) & ""","
        outdata = outdata & """ErrMsg"":""" & EscapeJSON(errmsg) & ""","
        outdata = outdata & """SessId"":""" & SessionID & ""","
        outdata = outdata & """LogOut"":""" & EscapeJSON(LOGOUT) & ""","
        outdata = outdata & """Username"":""" & EscapeJSON(USER) & ""","
        outdata = outdata & """ContactId"":""" & EscapeJSON(CON_ID) & ""","
        outdata = outdata & """LangCd"":""" & LANG_CD & ""","
        outdata = outdata & """EmailAddr"":""" & EscapeJSON(EMAIL_ADDR) & ""","
        outdata = outdata & """ParticipantId"":""" & PART_ID & ""","
        outdata = outdata & """TrainerFlag"":""" & TRAINER_FLG & """ "
        outdata = callback & "({""ResultSet"": {" & outdata & "} })"
        
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
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("WSCLogin.ashx: " & ErrLvl & ": " & Trim(errmsg))
        myeventlog.Info("WSCLogin.ashx : Id: " & REG_NUM & ", SessId: " & SessionID & ", Domain: " & DOMAIN & ", ConId: " & CON_ID)
        If Debug = "Y" Or (Logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("  outdata: " & outdata & vbCrLf)
                mydebuglog.Debug("Results:  Id: " & REG_NUM & ", SessId: " & SessionID & ", Domain: " & DOMAIN & ", ConId: " & CON_ID)
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
                LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, "WSCLOGIN", LogStartTime, VersionNum, Debug)
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
            If outdata = "" Then jdoc = errmsg
            context.Response.ContentType = "application/json"
            context.Response.Write(outdata)
        End If
    End Sub
 
    ' =================================================d
    ' SUPPORT FUNCTIONS
    Function GetLogout(ByVal DOMAIN As String, ByVal LANG_CD As String, ByVal authpath As String) As String
        Dim LOGOUT As String
        Select Case DOMAIN
            Case "TIPS"
                LOGOUT = "https://hciscorm.certegrity.com/ls/bcplogout.ashx?LANG=" & LANG_CD & "&RD=https://www.gettips.com/logout.shtml"
            Case "PBSA"
                LOGOUT = "https://hciscorm.certegrity.com/ls/bcplogout.ashx?LANG=" & LANG_CD & "&RD=http://www.tipsalcohol.com/"
            Case "IMIRDB"
                LOGOUT = "https://hciscorm.certegrity.com/ls/bcplogout.ashx?LANG=" & LANG_CD & "&RD=http://www.ebevlaw.com/"
            Case "HACCP"
                LOGOUT = "https://hciscorm.certegrity.com/ls/bcplogout.ashx?LANG=" & LANG_CD & "&RD=https://www.gettips.com/logout.shtml"
            Case "CSI"
                LOGOUT = "https://hciscorm.certegrity.com/ls/bcplogout.ashx?LANG=" & LANG_CD & "&RD=http://www.compliancetracking.com/index.html?logout"
            Case "BARLAP"
                LOGOUT = "https://hciscorm.certegrity.com/ls/bcplogout.ashx?LANG=" & LANG_CD & "&RD=https://www.barlap.us/glsuite/home/homeframe.aspx"
            Case Else
                LOGOUT = "https://hciscorm.certegrity.com/ls/bcplogout.ashx?LANG=" & LANG_CD & "&RD=http://www.compliancetracking.com/index.html?logout"
        End Select
        GetLogout = LOGOUT
    End Function
 
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