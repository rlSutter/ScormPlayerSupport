<%@ WebHandler Language="VB" Class="RemoteLogin" %>

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

Public Class RemoteLogin : Implements IHttpHandler
    
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

        ' Logging declarations
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("RemoteLoginDebugLog")
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
        Dim REF_URL As String = Trim(context.Request.ServerVariables("HTTP_REFERER"))
        Dim REMOTE_ADDR As String = Trim(context.Request.ServerVariables("REMOTE_ADDR"))
        Dim HTTP_HOST As String = Trim(context.Request.ServerVariables("HTTP_HOST"))
        Dim BROWSER As String = Trim(context.Request.ServerVariables("HTTP_USER_AGENT"))
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
        Dim REG_NUM, HOME_PAGE, LANG_CD, UID, callback, DOMAIN, SessionID As String
        Dim LOGGED_IN, SUB_ID, CONTACT_OU_ID, LaunchProtocol As String
        Dim T_UserName, T_Pwd, Remember, NXT, Path, ReDirect, output, EOL As String
        Dim CON_ID, TERM_FLG, SUB_CON_ID, SVC_TYPE, NextMsg, loginsuccess, ErrURL, CallURL As String
        Dim Start, EndS As Integer
        
        ' ============================================
        ' Variable setup
        ErrURL = ""
        CallURL = ""
        Debug = "N"
        Logging = "Y"
        errmsg = ""
        LOGGED_IN = "N"
        UID = ""
        NextMsg = ""
        loginsuccess = ""
        CON_ID = ""
        SUB_ID = ""
        CONTACT_OU_ID = ""
        LANG_CD = "ENU"
        callback = ""
        DOMAIN = "TIPS"
        HOME_PAGE = ""
        REG_NUM = ""
        UID = ""
        SessID = ""
        TERM_FLG = ""
        SUB_CON_ID = ""
        SVC_TYPE = ""
        ErrLvl = "Error"
        T_UserName = ""
        LaunchProtocol = "http:"
        T_Pwd = ""
        Remember = ""
        NXT = ""
        Path = ""
        ReDirect = ""
        output = ""
        Start = 0
        EndS = 0
        SessionID = ""
        EOL = Chr(13) & Chr(10)
        
        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
            tempdebug = System.Configuration.ConfigurationManager.AppSettings.Get("RemoteLogin_debug")
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
            logfile = "C:\Logs\RemoteLogin.log"
            Try
                log4net.GlobalContext.Properties("RemoteLoginLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                GoTo CloseOut2
            End Try
        End If
        
        ' ============================================
        ' Get parameters    
        If Not context.Request.Form("T_UserName") Is Nothing Then
            T_UserName = context.Request.Form("T_UserName")
        End If
        
        If Not context.Request.Form("T_Pwd") Is Nothing Then
            T_Pwd = context.Request.Form("T_Pwd")
        End If

        If Not context.Request.Form("Remember") Is Nothing Then
            Remember = context.Request.Form("Remember")
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
            mydebuglog.Debug("  Cookie User Id: " & UserID)
            mydebuglog.Debug("  Cookie Session Id: " & SessID)
            mydebuglog.Debug("  T_UserName: " & T_UserName)
            'mydebuglog.Debug("  T_Pwd: " & T_Pwd)
            mydebuglog.Debug("  Remember: " & Remember & vbCrLf)
        End If
        
        If T_Pwd = "" Or T_UserName = "" Then GoTo AccessError
        
        ' Get the domain from the referrer
        If REF_URL <> "" Then
            Start = InStr(1, LCase(REF_URL), "&dom=")
            If Start > 0 Then
                Start = Start + 5
                EndS = InStr(LCase(REF_URL), "&rec=")
                If EndS = 0 Then
                    EndS = Len(REF_URL) + 1
                End If
                DOMAIN = UCase(Trim(Mid(REF_URL, Start, EndS - Start)))
            End If
        End If
        If Debug = "Y" Then mydebuglog.Debug("DOMAIN: " & DOMAIN & vbCrLf)
        
        ' ================================================
        ' DETERMINE THE INSTANCE TO USE
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
        ReDirect = Path & "/main?OpenForm"
                
        ' ============================================
        ' Open database connection 
OpenDB:
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            errmsg = errmsg & "Unable to open the database connection. " & vbCrLf
            GoTo SystemUnavailable
        End If

        ' ============================================
        ' Prepare results
        If Not cmd Is Nothing Then
            
            ' ================================================   
            ' LOOKUP SUBSCRIPTION            
            ' If unable to locate, then subscription not setup, error out
            SqlS = "SELECT TOP 1 C.X_REGISTRATION_NUM, C.ROW_ID, S.DOMAIN, " & _
                   "(SELECT CASE WHEN S.SVC_TERM_DT<GETDATE() AND S.SVC_TYPE<>'PUBLIC ACCESS' THEN 'Y' ELSE 'N' END) AS TERM_FLG, SC.ROW_ID, S.SVC_TYPE, S.ROW_ID " & _
                   "FROM siebeldb.dbo.S_CONTACT C " & _
                   "LEFT OUTER JOIN siebeldb.dbo.CX_SUB_CON SC ON SC.CON_ID=C.ROW_ID " & _
                   "LEFT OUTER JOIN siebeldb.dbo.CX_SUBSCRIPTION S ON S.ROW_ID=SC.SUB_ID "
            If InStr(1, T_UserName, "@") Then
                SqlS = SqlS & "WHERE upper(EMAIL_ADDR)='" & UCase(T_UserName) & "' AND C.PWD_IDX=HASHBYTES('SHA2_256','" & UCase(T_Pwd) & "')"
            Else
                SqlS = SqlS & "WHERE upper(LOGIN)='" & UCase(T_UserName) & "' AND C.PWD_IDX=HASHBYTES('SHA2_256','" & UCase(T_Pwd) & "')"
            End If
            SqlS = SqlS & " ORDER BY C.X_TRAINER_NUM DESC, C.X_PART_ID DESC"
            'If Debug = "Y" Then mydebuglog.Debug("Subscription Query: " & vbCrLf & "  " & SqlS)
            Try
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        REG_NUM = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                        CON_ID = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                        DOMAIN = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                        If DOMAIN <> "" Then ReDirect$ = ReDirect$ & "&DOM=" & DOMAIN
                        TERM_FLG = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                        SUB_CON_ID = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                        SVC_TYPE = Trim(CheckDBNull(dr(5), enumObjectType.StrType))
                        If SVC_TYPE = "PUBLIC ACCESS" Then TERM_FLG = "N"
                        SUB_ID = Trim(CheckDBNull(dr(6), enumObjectType.StrType))
                    End While
                    If Len(Trim(REG_NUM)) > 0 Then
                        UID = REG_NUM
                        NextMsg = "<BR><B><FONT COLOR=""800080"" FACE=""Verdana, Arial"">One moment please...</FONT></B><BR>"
                        loginsuccess = "Y"
                    Else
                        UID = ""
                        NextMsg = "<font size=1><font color=""Red"">This information is incorrect..</font></span><BR>Please <A HREF=""JavaScript:openNewWindow('https://w1.certegrity.com/plogini.nsf/Help/13?OpenDocument',350,250)"">verify your username and password</A>. </font>"
                        loginsuccess = "N"
                    End If
                Else
                    UID = ""
                    NextMsg = "<font size=1><font color=""Red"">This information is incorrect..</font></span><BR>Please <A HREF=""JavaScript:openNewWindow('https://w1.certegrity.com/plogini.nsf/Help/13?OpenDocument',350,250)"">verify your username and password</A>. </font>"
                    loginsuccess = "N"
                    GoTo AccessError
                End If
            Catch ex As Exception
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Unable to locate credentials. Error: " & vbCrLf & ex.ToString & vbCrLf)
                GoTo SystemUnavailable
            End Try
            dr.Close()
            If Debug = "Y" Then
                mydebuglog.Debug("  .. REG_NUM: " & REG_NUM)
                mydebuglog.Debug("  .. CON_ID: " & CON_ID)
                mydebuglog.Debug("  .. DOMAIN: " & DOMAIN)
                mydebuglog.Debug("  .. TERM_FLG: " & TERM_FLG)
                mydebuglog.Debug("  .. SUB_CON_ID: " & SUB_CON_ID)
                mydebuglog.Debug("  .. SUB_ID: " & SUB_ID)
                mydebuglog.Debug("  .. SVC_TYPE: " & SVC_TYPE & vbCrLf)
            End If
            If SUB_CON_ID = "" Then GoTo AccessError
            If SUB_ID = "" Then GoTo SubscriptionError            

            ' Translate source for destination
            Select Case Trim(LCase(REF_URL))
                Case "http://www.haccp-training.com/"
                    ErrURL = "http://www.haccp-training.com/indexerr.html"
                    CallURL = REF_URL
                Case "http://www.ebevlaw.com/"
                    ErrURL = "http://www.ebevlaw.com/indexerr.html"
                    CallURL = REF_URL
                Case "http://www.ebevlaw.com/index.html"
                    ErrURL = "http://www.ebevlaw.com/indexerr.html"
                    CallURL = REF_URL
                Case "http://www.ebevlaw.com/indexerr.html"
                    ErrURL = "http://www.ebevlaw.com/indexerr.html"
                    CallURL = REF_URL
                Case "https://web.gettips.com"
                    ErrURL = "https://www.gettips.com/indexerr.shtml"
                    CallURL = REF_URL
                Case "https://web.gettips.com/index.shtml"
                    ErrURL = "https://www.gettips.com/indexerr.shtml"
                    CallURL = REF_URL
                Case "https://web.gettips.com/indexerr.shtml"
                    ErrURL = "https://www.gettips.com/indexerr.shtml"
                    CallURL = REF_URL
                Case "http://web.gettips.com"
                    ErrURL = "http://www.gettips.com/indexerr.shtml"
                    CallURL = REF_URL
                Case "http://web.gettips.com/index.shtml"
                    ErrURL = "http://www.gettips.com/indexerr.shtml"
                    CallURL = REF_URL
                Case "http://web.gettips.com/indexerr.shtml"
                    ErrURL = "http://web.gettips.com/indexerr.shtml"
                    CallURL = REF_URL
                Case Else
                    If InStr(1, LCase(REF_URL), "index") > 0 Then
                        CallURL = REF_URL
                        If InStr(1, LCase(REF_URL), "indexerr") = 0 Then
                            ErrURL = Left(REF_URL, InStr(1, LCase(REF_URL), ".html") - 1) & "err.html"
                        Else
                            ErrURL = CallURL
                        End If
                    Else
                        Start = 0
                        EndS = 0
                        Start = InStr(1, REF_URL, "http:")
                        If Start = 0 Then Start = InStr(1, REF_URL, "https:")
                        If Start > 0 Then
                            EndS = InStr(REF_URL, ".nsf") + 5
                            If EndS = 0 Then
                                EndS = Len(REF_URL) + 1
                            End If
                            CallURL = Trim(Mid(REF_URL, Start, EndS - Start))
                        End If
                        ErrURL = CallURL & "main?OpenForm&PAG=RELOGIN&DOM=" & DOMAIN
                    End If
            End Select
            If Debug = "Y" Then
                mydebuglog.Debug("ErrURL: " & ErrURL)
                mydebuglog.Debug("CallURL: " & CallURL & vbCrLf)
            End If
            
            ' Reset service type to PUBLIC ACCESS if expired and accessing via public portal
            If TERM_FLG = "Y" And SVC_TYPE <> "PUBLIC ACCESS" Then
                SqlS = "UPDATE siebeldb.dbo.CX_SUBSCRIPTION SET SVC_TYPE='PUBLIC ACCESS' WHERE ROW_ID='" & SUB_ID & "'"
                temp = ExecQuery("Update", "CX_SUBSCRIPTION", cmd, SqlS, mydebuglog, Debug)
                GoTo ExpiredSubscription
            End If
            
            ' ==============================   
            ' GENERATE A SESSION ID
            If loginsuccess = "Y" Then
                Randomize()
                SessionID = UCase(LoggingService.GeneratePassword(Debug)) & NXT & Chr(Str(Int(Rnd() * 26)) + 65) & Chr(Str(Int(Rnd() * 26)) + 65) & Chr(Str(Int(Rnd() * 26)) + 65)
            End If
            If Debug = "Y" Then mydebuglog.Debug("SessionID: " & SessionID & vbCrLf)
            
            ' ==============================
            ' LOG IF APPLICABLE
            If loginsuccess = "Y" And SUB_ID <> "" And REG_NUM <> "" Then

                ' Log the user's activities in their personal record
                mydebuglog.Debug(vbCrLf & "Logging Queries- ")
                SqlS = "UPDATE siebeldb.dbo.CX_SUB_CON " & _
                       "SET LAST_INST='" & Path & "', LAST_LOGIN=GETDATE(), LAST_SESS_ID='" & SessionID & "' " & _
                       "FROM (SELECT ROW_ID FROM siebeldb.dbo.S_CONTACT WHERE X_REGISTRATION_NUM='" & REG_NUM & "') U " & _
                       "WHERE siebeldb.dbo.CX_SUB_CON.CON_ID=U.ROW_ID "
                temp = ExecQuery("Update", "CX_SUB_CON", cmd, SqlS, mydebuglog, Debug)
                    
                SqlS = "INSERT INTO reports.dbo.CM_LOG(REG_ID, SESSION_ID, ACTION, REMOTE_ADDR, BROWSER) " & _
                           "VALUES('" & REG_NUM & "','" & SessionID & "','REMOTELOGIN.ashx LOGIN', '" & REMOTE_ADDR & "','" & BROWSER & "')"
                temp = ExecQuery("Insert", "CM_LOG", cmd, SqlS, mydebuglog, Debug)
                        
                SqlS = "INSERT siebeldb.dbo.CX_SUB_CON_HIST(CONFLICT_ID,CREATED_BY,LAST_UPD_BY,ROW_ID," & _
                       "SUB_CON_ID,USER_ID,SESSION_ID,REMOTE_ADDR) " & _
                       "SELECT 0,'1-3HIZ7','1-3HIZ7','" & SessionID & "', " & _
                       "SC.ROW_ID,'" & REG_NUM & "','" & SessionID & "','" & REMOTE_ADDR & "' " & _
                       "FROM siebeldb.dbo.S_CONTACT C " & _
                       "LEFT OUTER JOIN siebeldb.dbo.CX_SUB_CON SC ON SC.CON_ID=C.ROW_ID " & _
                       "WHERE C.X_REGISTRATION_NUM='" & REG_NUM & "' AND NOT EXISTS " & _
                       "(SELECT ROW_ID FROM siebeldb.dbo.CX_SUB_CON_HIST WHERE SESSION_ID='" & SessionID & "' AND USER_ID='" & REG_NUM & "')"
                temp = ExecQuery("Insert", "CX_SUB_CON_HIST", cmd, SqlS, mydebuglog, Debug)
                
                ' Redirect the user
                output = output & "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 3.2//EN"">"
                output = output & "<html>"
                output = output & "<head>"
                output = output & "<title>One moment please...</title>"
                output = output & "<link rel=""stylesheet"" href=""https://www.gettips.com/css/stylesheet.css"" type=""text/css"">"
                output = output & "<script language=""JavaScript""> " & EOL & _
                "function checkLogin() { " & EOL & _
                "   uid = '" & Trim(REG_NUM) & "';" & EOL & _
                "   sessid = '" & SessionID & "';" & EOL & _
                "   if (uid!='' && sessid!='') { " & EOL & _
                "       var expdate = new Date();" & EOL & _
                "      expdate.setTime (expdate.getTime() +  86400000);" & EOL & _
                "      SetCookie(""ID"",uid, expdate, ""/"" , "".certegrity.com""); " & EOL & _
                "      SetCookie(""Sess"",sessid, ""/"" , ""/"" , "".certegrity.com""); " & EOL & _
                "      window.location.href=""" & ReDirect & """;" & EOL & _
                "   }" & EOL & _
                "   else {" & EOL & _
                "      window.location.href=""" & ErrURL & """;" & EOL & _
                "   }" & EOL & _
                "} " & EOL & _
                "function SetCookie (name, value) { " & EOL & _
                "     var argv = SetCookie.arguments; " & EOL & _
                "     var argc = SetCookie.arguments.length; " & EOL & _
                "     var expires = (argc > 2) ? argv[2] : null; " & EOL & _
                "     var path = (argc > 3) ? argv[3] : null; " & EOL & _
                "     var domain = (argc > 4) ? argv[4] : null; " & EOL & _
                "     var secure = (argc > 5) ? argv[5] : false; " & EOL & _
                "     document.cookie = name + ""="" + escape (value) + " & EOL & _
                "       ((path == null) ? """" : (""; path="" + path)) + " & EOL & _
                "       ((domain == null) ? """" : (""; domain="" + domain)) + " & EOL & _
                "       ((secure == true) ? ""; secure"" : """"); " & EOL & _
                "} " & EOL & _
                "function baseDomainString(){ " & EOL & _
                "     e = document.domain.split(/\./); " & EOL & _
                "     if(e.length > 1) { " & EOL & _
                "         return(e[e.length-2] + ""."" +  e[e.length-1]); " & EOL & _
                "     } else { " & EOL & _
                "         return(""""); " & EOL & _
                "     } " & EOL & _
                "} " & EOL & _
                "</script> " & EOL & _
                "<style type=""text/css""><!-- body {  margin-top: 2px; margin-right: 2px; margin-bottom: 2px; margin-left: 2px} --></style>" & EOL & _
                "</head>" & EOL & _
                "<link href=""https://www.gettips.com/css/" & Trim(LCase(DOMAIN)) & ".css"" rel=""stylesheet"">" & EOL & _
                "<body onLoad=""checkLogin()"">" & EOL & _
                "<TABLE WIDTH=""100%"" HEIGHT=""100%"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0""><TR ALIGN=""Center"" VALIGN=""Top""><TD HEIGHT=""100%"">" & EOL & _
                "<br /><TABLE BORDER=0 CELLSPACING=0 width=""700"" BGCOLOR=""FFFFFF""><TR VALIGN=top><TD CLASS=""Heading"">You are now logged in. <br>One moment please...<BR></TD></TR></TABLE>" & EOL & _
                "</body></html>"
            Else
                ' Unsuccessful login     
                output = output & "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">"
                output = output & "<html>"
                output = output & "<head>"
                output = output & "<meta http-equiv=""Pragma"" content=""no-cache"">"
                output = output & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
                output = output & "<META HTTP-EQUIV=""Refresh"" CONTENT=""0; URL=" & ErrURL & """>"
                output = output & "<meta name=""GOOGLEBOT"" CONTENT=""NOARCHIVE"">"
                output = output & "<title>Compliance Solutions, Inc.</title><style type=""text/css"">"
                output = output & "<link rel=""stylesheet"" href=""https://www.gettips.com/css/stylesheet.css"" type=""text/css"">"
                output = output & "</head>"
                output = output & "<body bgcolor=""#FFFFFF"" text=""#000000"" leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
                output = output & "</body>"
                output = output & "</html>"
            End If
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
        output = output & "<title>Login Error</title></head>"
        output = output & "</HEAD>"
        output = output & "<BODY BGCOLOR='White' leftmargin=0 text='#000040' link='Purple' vlink='Navy'>"
        output = output & "<br><font size=2 face='Arial, Sans-Serif' color='red'>There was an error reported</font><br>"
        output = output & "<B><FONT SIZE=2 COLOR='800080' FACE='Arial'>The system may be unavailable now.  Please try again later</FONT></B><BR>"
        output = output & "</BODY>"
        output = output & "</HTML>"
        GoTo CloseOut
         
ExpiredSubscription:
        If Debug = "Y" Then mydebuglog.Debug(">>ExpiredSubscription")
        errmsg = "Expired Subscription"
        ErrLvl = "Warning"
        output = output & "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 3.2//EN"">"
        output = output & "<HTML>"
        output = output & "<HEAD>"
        output = output & "<META HTTP-EQUIV=""Refresh"" CONTENT=""1; URL=http://www.compliancetracking.com/expired.html"">"
        output = output & "</HEAD>"
        output = output & "<BODY BGCOLOR='White' leftmargin=0 text='#000040' link='Purple' vlink='Navy'>"
        output = output & "</BODY></HTML>"
        GoTo CloseOut
        
SubscriptionError:
        If Debug = "Y" Then mydebuglog.Debug(">>SubscriptionError")
        errmsg = "Subscription Error"
        ErrLvl = "Warning"
        output = output & "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 3.2//EN"">"
        output = output & "<HTML>"
        output = output & "<HEAD>"
        output = output & "<META HTTP-EQUIV=""Refresh"" CONTENT=""1; URL=http://www.compliancetracking.com/accesserror.html"">"
        output = output & "<script language=""JavaScript"">"
        output = output & "function dynamicLogout() {"
        output = output & "   var hosting = baseDomainString();"
        output = output & "   DeleteCookie(""ID"",""/"",hosting);"
        output = output & "   DeleteCookie(""Sess"",""/"",hosting);"
        output = output & "}"
        output = output & "function DeleteCookie( name, path, domain ) {"
        output = output & "    document.cookie = name + ""="" + ( ( path ) ? "";path="" + path : """") + ( ( domain ) ? "";domain="" + domain : """" ) + "";expires=Thu, 01-Jan-1970 00:00:01 GMT"";"
        output = output & "}"
        output = output & "function baseDomainString(){"
        output = output & "     e = document.domain.split(/\./);"
        output = output & "     if(e.length > 1) {"
        output = output & "       return(e[e.length-2] + ""."" +  e[e.length-1]);"
        output = output & "     } else {"
        output = output & "       return("""");"
        output = output & "     }"
        output = output & "}"
        output = output & "</script>"
        output = output & "<body onload=""dynamicLogout()"">"
        GoTo CloseOut
        
AccessError:
        ErrLvl = "Warning"
        If Debug = "Y" Then mydebuglog.Debug(">>AccessError")
        errmsg = "Access Error"
        output = output & "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 3.2//EN"">"
        output = output & "<HTML>"
        output = output & "<HEAD>"
        output = output & "<title>Login Error</title></head>"
        output = output & "</HEAD>"
        output = output & "<BODY BGCOLOR='White' leftmargin=0 text='#000040' link='Purple' vlink='Navy'><center>"
        If NextMsg <> "" Then
            output = output & "<br /><font size=1 face='Arial, Sans-Serif' color='red'>" & NextMsg & "</font><br /><br />"
            output = output & "<B><FONT SIZE=2 COLOR='800080' FACE='Arial'><a href=""javascript:history.go(-1)"">Please try again</a></FONT></B><BR>"
        Else
            output = output & "<br /><font size=1 face='Arial, Sans-Serif' color='red'>Please provide credentials</font><br /><br />"
            output = output & "<B><FONT SIZE=2 COLOR='800080' FACE='Arial'>There was a problem with your username and login. <a href=""javascript:history.go(-1)"">Please try again</a></FONT></B><BR>"
        End If
        output = output & "</center></BODY>"
        output = output & "</HTML>"
        GoTo CloseOut
        
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
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("RemoteLogin.ashx : " & ErrLvl & ": " & Trim(errmsg))
        myeventlog.Info("RemoteLogin.ashx : REG_NUM: " & REG_NUM & ", SessionID: " & SessionID & ", ReDirect: " & ReDirect)
        If Debug = "Y" Or (Logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug(vbCrLf & "Results:  REG_NUM: " & REG_NUM & ", SessionID: " & SessionID & ", ReDirect: " & ReDirect)
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
                LoggingService.LogPerformanceDataAsync(System.Environment.MachineName.ToString, "RemoteLogin", LogStartTime, VersionNum, Debug)
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