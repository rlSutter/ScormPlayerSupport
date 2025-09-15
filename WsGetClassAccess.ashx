<%@ WebHandler Language="VB" Class="WsGetClassAccess" %>

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

Public Class WsGetClassAccess : Implements IHttpHandler
    
    Enum enumObjectType
        StrType = 0
        IntType = 1
        DblType = 2
        DteType = 3
    End Enum
    
    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        
        ' Description
        '
        
        ' Parameter Declarations
        Dim Debug, CRSE_TSTRUN_ID, RETURN_PAGE, DOMAIN As String
        Dim REG_ID, UID, SessID, CURRENT_PAGE, HOME_PAGE, LANG_CD, callback, myprotocol As String
        
        ' Result Declarations
        Dim jdoc, outdata As String
        
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
        mydebuglog = log4net.LogManager.GetLogger("GetClassAccessDebugLog")
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
        Dim errmsg As String                    ' Error message (if any)
        Dim EOL, LaunchProtocol, ReturnDestination, NextLink, ErrLvl As String
        Dim NewParticipation, RegFound, PortalLink As Boolean
        Dim MSG1, CRSE_NAME As String
        Dim CERT_LINK, SCORM_FLG, PASSED_FLG, TEST_STATUS, REG_STATUS_CD, TRAIN_TYPE As String
        Dim LOGGED_IN, CONTACT_ID, SUB_ID, CONTACT_OU_ID, OFFR_ID As String
        Dim RegLink, RegList, PaymentLink, ClassLink, ReviewReg, Refresh, LAST_INST As String
        Dim CRSE_CONTENT_URL, CRSE_ID, COURSE, RESOLUTION, JURIS_ID, KBA_REQD, temp As String
        Dim ORDER_ITEM_ID, PAYMENT_REQD, BL_PROD_INT_ID, ORDER_ID, EMAIL_ADDR, ORDER_STATUS, ORDER_CONTACT_ID, REG_CONTACT_ID, ltemp As String
        Dim KBA_QUESTIONS, TotalPgm, KBA_COUNT, NUM_ANSRD As Integer
        
        ' ============================================
        ' Variable setup
        Debug = "Y"
        Logging = "Y"
        UID = ""
        SessID = ""
        LOGGED_IN = "N"
        CONTACT_ID = ""
        SUB_ID = ""
        CONTACT_OU_ID = ""
        DOMAIN = "TIPS"
        LANG_CD = "ENU"
        callback = ""
        myprotocol = ""
        HOME_PAGE = ""
        RETURN_PAGE = ""
        CURRENT_PAGE = ""
        CRSE_TSTRUN_ID = ""
        REG_ID = ""
        errmsg = ""
        EOL = Chr(10)
        LaunchProtocol = "http:"
        NewParticipation = False
        ReturnDestination = ""
        NextLink = ""
        outdata = ""
        jdoc = ""
        ErrLvl = "Error"
        MSG1 = ""
        CRSE_NAME = ""
        REG_STATUS_CD = ""
        TEST_STATUS = ""
        PASSED_FLG = "N"
        SCORM_FLG = "Y"
        CERT_LINK = ""
        TRAIN_TYPE = ""
        CONTACT_ID = ""
        OFFR_ID = ""
        RegLink = ""
        RegList = ""
        PaymentLink = ""
        ClassLink = ""
        ReviewReg = ""
        Refresh = ""
        RegFound = False
        TotalPgm = 0
        ORDER_ITEM_ID = ""
        PAYMENT_REQD = ""
        BL_PROD_INT_ID = ""
        JURIS_ID = ""
        KBA_REQD = ""
        CRSE_ID = ""
        EMAIL_ADDR = ""
        ORDER_ID = ""
        ORDER_STATUS = ""
        COURSE = ""
        KBA_COUNT = 0
        NUM_ANSRD = 0
        PortalLink = False
        LAST_INST = ""
        CRSE_CONTENT_URL = ""
        
        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
            tempdebug = System.Configuration.ConfigurationManager.AppSettings.Get("GetClassAccess_debug")
            If tempdebug = "Y" And Debug <> "T" Then Debug = "Y"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            GoTo CloseOut2
        End Try
        
        ' ============================================
        ' Open log file if applicable
        If Logging = "Y" Then
            logfile = "C:\Logs\WsGetClassAccess.log"
            Try
                log4net.GlobalContext.Properties("GetClassAccessLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                GoTo CloseOut2
            End Try
        End If
        
        ' ============================================
        ' Get parameters    
        If Not context.Request.QueryString("UID") Is Nothing Then
            UID = context.Request.QueryString("UID")
        End If
        
        If Not context.Request.QueryString("SES") Is Nothing Then
            SessID = context.Request.QueryString("SES")
        End If

        If Not context.Request.QueryString("PP") Is Nothing Then
            DOMAIN = context.Request.QueryString("PP")
        End If

        If Not context.Request.QueryString("CID") Is Nothing Then
            OFFR_ID = context.Request.QueryString("CID")
        End If
        
        If OFFR_ID = "" Then
            If Not context.Request.QueryString("OFR") Is Nothing Then
                OFFR_ID = context.Request.QueryString("OFR")
            End If            
        End If
          
        If Not context.Request.QueryString("RG") Is Nothing Then
            REG_ID = context.Request.QueryString("RG")
        End If
        
        If Not context.Request.QueryString("LANG") Is Nothing Then
            LANG_CD = UCase(context.Request.QueryString("LANG"))
        End If
        
        If Not context.Request.QueryString("HP") Is Nothing Then
            HOME_PAGE = context.Request.QueryString("HP")
        End If
 
        If Not context.Request.QueryString("callback") Is Nothing Then
            callback = context.Request.QueryString("callback")
        End If

        If Not context.Request.QueryString("PROT") Is Nothing Then
            myprotocol = LCase(context.Request.QueryString("PROT"))
        End If
        
        
        ' Validate parameters
        If OFFR_ID = "" Then OFFR_ID = "1-IGDAB"
        If LANG_CD <> "ESN" And LANG_CD <> "ENU" Then LANG_CD = "ENU"
        If callback = "" Then callback = "?"
        If myprotocol = "" Then myprotocol = "http:"
        If InStr(1, PrevLink, "?UID") = 0 Then PrevLink = PrevLink & "?UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN
        PrevLink = Replace(PrevLink, "#reg", "")
        If Left(HOME_PAGE, 4) <> "web." And Left(HOME_PAGE, 4) <> "www." Then
            If InStr(1, PrevLink, "web.") > 0 Then HOME_PAGE = "web." & HOME_PAGE Else HOME_PAGE = "www." & HOME_PAGE
        End If

        If Debug = "Y" Then
            mydebuglog.Debug("----------------------------------")
            mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
            mydebuglog.Debug("Parameters-")
            mydebuglog.Debug("  Debug: " & Debug)
            mydebuglog.Debug("  UID: " & UID)
            mydebuglog.Debug("  cookieid: " & cookieid)
            mydebuglog.Debug("  SessID: " & SessID)
            mydebuglog.Debug("  OFFR_ID : " & OFFR_ID)
            mydebuglog.Debug("  REG_ID : " & REG_ID)
            mydebuglog.Debug("  PrevLink: " & PrevLink)
            mydebuglog.Debug("  myprotocol: " & myprotocol)
            mydebuglog.Debug("  LANG_CD: " & LANG_CD)
            mydebuglog.Debug("  DOMAIN: " & DOMAIN)
            mydebuglog.Debug("  BROWSER: " & BROWSER)
            mydebuglog.Debug("  HOME_PAGE: " & HOME_PAGE)
            mydebuglog.Debug("  callback: " & callback)
        End If
        
        If UID = "" Then GoTo NotLoggedIn
        If OFFR_ID = "" And REG_ID = "" Then GoTo DataError
        If cookieid <> UID Then GoTo AccessError
        
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
                errmsg = ex.ToString & vbCrLf
                GoTo DBError
            End Try
            dr.Close()
            If Debug = "Y" Then
                mydebuglog.Debug("  .. LOGGED_IN: " & LOGGED_IN)
                mydebuglog.Debug("  .. CONTACT_ID: " & CONTACT_ID)
                mydebuglog.Debug("  .. CONTACT_OU_ID: " & CONTACT_OU_ID)
                mydebuglog.Debug("  .. DOMAIN: " & DOMAIN)
            End If
            If LOGGED_IN = "N" Then GoTo AccessError
        
            ' ================================================
            ' GET REGISTRATION INFORMATION
            SqlS = "SELECT R.STATUS_CD, T.CRSE_CONTENT_URL, R.ROW_ID, T.TRAIN_TYPE, R.ORDER_ITEM_ID, " & _
            "T.PAYMENT_REQD, T.BL_PROD_INT_ID, T.DOMAIN, CR.X_SCORM_FLG, CR.X_RESOLUTION, " & _
            "R.JURIS_ID, JC.KBA_REQD, JC.KBA_QUESTIONS, CR.ROW_ID, C.EMAIL_ADDR, O.ROW_ID, O.STATUS_CD, CR.NAME, " & _
            "O.CONTACT_ID, R.CONTACT_ID, OI.ROW_ID, C.X_PR_LANG_CD " & _
            "FROM siebeldb.dbo.CX_SESS_REG R " & _
            "LEFT OUTER JOIN siebeldb.dbo.CX_TRAIN_OFFR T ON T.ROW_ID=R.TRAIN_OFFR_ID " & _
            "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=R.CONTACT_ID " & _
            "LEFT OUTER JOIN siebeldb.dbo.S_CRSE CR ON CR.ROW_ID=R.CRSE_ID " & _
            "LEFT OUTER JOIN siebeldb.dbo.CX_JURIS_CRSE JC ON JC.JURIS_ID=R.JURIS_ID AND JC.CRSE_ID=R.CRSE_ID " & _
            "LEFT OUTER JOIN siebeldb.dbo.S_ORDER_ITEM OI ON OI.ROW_ID=R.ORDER_ITEM_ID " & _
            "LEFT OUTER JOIN siebeldb.dbo.S_ORDER O ON O.ROW_ID = OI.ORDER_ID OR O.ROW_ID=R.ROW_ID " & _
            "WHERE C.X_REGISTRATION_NUM='" & UID & "'"
            If OFFR_ID <> "" Then SqlS = SqlS & " AND R.TRAIN_OFFR_ID='" & OFFR_ID & "'"
            If REG_ID <> "" Then SqlS = SqlS & " AND R.ROW_ID='" & REG_ID & "'"
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  GET SPECIFIC REGISTRATION INFORMATION: " & vbCrLf & "  " & SqlS)
            Try
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        RegFound = True
                        TotalPgm = 1
                        REG_STATUS_CD = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                        CRSE_CONTENT_URL = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                        REG_ID = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                        TRAIN_TYPE = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                        ORDER_ITEM_ID = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                        PAYMENT_REQD = Trim(CheckDBNull(dr(5), enumObjectType.StrType))
                        BL_PROD_INT_ID = Trim(CheckDBNull(dr(6), enumObjectType.StrType))
                        If DOMAIN = "" Then DOMAIN = Trim(CheckDBNull(dr(7), enumObjectType.StrType))
                        SCORM_FLG = Trim(CheckDBNull(dr(8), enumObjectType.StrType))
                        COURSE = Trim(CheckDBNull(dr(9), enumObjectType.StrType))
                        JURIS_ID = Trim(CheckDBNull(dr(10), enumObjectType.StrType))
                        KBA_REQD = Trim(CheckDBNull(dr(11), enumObjectType.StrType))
                        temp = Trim(CheckDBNull(dr(12), enumObjectType.StrType))
                        If temp = "" Then KBA_QUESTIONS = 0 Else KBA_QUESTIONS = Val(temp)
                        CRSE_ID = Trim(CheckDBNull(dr(13), enumObjectType.StrType))
                        EMAIL_ADDR = Trim(CheckDBNull(dr(14), enumObjectType.StrType))
                        ORDER_ID = Trim(CheckDBNull(dr(15), enumObjectType.StrType))
                        ORDER_STATUS = Trim(CheckDBNull(dr(16), enumObjectType.StrType))
                        COURSE = Trim(CheckDBNull(dr(17), enumObjectType.StrType))
                        ORDER_CONTACT_ID = Trim(CheckDBNull(dr(18), enumObjectType.StrType))
                        REG_CONTACT_ID = Trim(CheckDBNull(dr(19), enumObjectType.StrType))
                        ORDER_ITEM_ID = Trim(CheckDBNull(dr(20), enumObjectType.StrType))
                        ltemp = Trim(CheckDBNull(dr(21), enumObjectType.StrType))
                        If LANG_CD = "" And LANG_CD <> "" Then LANG_CD = ltemp
                        If LANG_CD = "" Then LANG_CD = "ENU"
                        If ORDER_CONTACT_ID <> REG_CONTACT_ID And ORDER_ID <> "" And ORDER_ITEM_ID = "" Then
                            ORDER_ID = ""
                        End If
                        
                        If Debug = "Y" Then mydebuglog.Debug("  .. REG_ID/REG_STATUS_CD: " & REG_ID & " / " & REG_STATUS_CD)
                        Select Case REG_STATUS_CD
                            Case "Tentative"
                                GoTo ProcStat
                            Case "Incomplete"
                                GoTo ProcStat
                            Case "In Progress"
                                GoTo ProcStat
                            Case "Accepted"
                                GoTo ProcStat
                            Case "Exam Reqd"
                                GoTo ProcStat
                            Case "On Hold"
                                GoTo ProcStat
                        End Select
                    End While
                End If
                dr.Close()
            Catch ex As Exception
                GoTo DBError
            End Try

            ' If not found, let's see if there is a registration for another class
            SqlS = "SELECT R.STATUS_CD, T.CRSE_CONTENT_URL, R.ROW_ID, T.TRAIN_TYPE, R.ORDER_ITEM_ID, " & _
            "T.PAYMENT_REQD, T.BL_PROD_INT_ID, T.DOMAIN, T.ROW_ID, CR.X_SCORM_FLG, CR.X_RESOLUTION, " & _
            "R.JURIS_ID, JC.KBA_REQD, JC.KBA_QUESTIONS, CR.ROW_ID, C.EMAIL_ADDR, O.ROW_ID, O.STATUS_CD, CR.NAME " & _
            "FROM siebeldb.dbo.CX_SESS_REG R " & _
            "LEFT OUTER JOIN siebeldb.dbo.CX_TRAIN_OFFR T ON T.ROW_ID=R.TRAIN_OFFR_ID " & _
            "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=R.CONTACT_ID " & _
            "LEFT OUTER JOIN siebeldb.dbo.S_CRSE CR ON CR.ROW_ID=R.CRSE_ID " & _
            "LEFT OUTER JOIN siebeldb.dbo.CX_JURIS_CRSE JC ON JC.JURIS_ID=R.JURIS_ID AND JC.CRSE_ID=R.CRSE_ID " & _
            "LEFT OUTER JOIN siebeldb.dbo.S_ORDER_ITEM OI ON OI.ROW_ID=R.ORDER_ITEM_ID " & _
            "LEFT OUTER JOIN siebeldb.dbo.S_ORDER O ON O.ROW_ID = OI.ORDER_ID " & _
            "WHERE C.X_REGISTRATION_NUM='" & UID & "' AND T.CRSE_CONTENT_URL<>'' AND R.STATUS_CD in ('Accepted','In Progress','Exam Reqd','Tentative','Incomplete','On Hold')"
            If OFFR_ID <> "" Then SqlS = SqlS & " AND R.TRAIN_OFFR_ID='" & OFFR_ID & "'"
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  GET ALL REGISTRATION INFORMATION: " & vbCrLf & "  " & SqlS)
            Try
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        REG_STATUS_CD = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                        If REG_STATUS_CD <> "" Then
                            TotalPgm = TotalPgm + 1
                            RegFound = True
                            CRSE_CONTENT_URL = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            REG_ID = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                            TRAIN_TYPE = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                            ORDER_ITEM_ID = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                            PAYMENT_REQD = Trim(CheckDBNull(dr(5), enumObjectType.StrType))
                            BL_PROD_INT_ID = Trim(CheckDBNull(dr(6), enumObjectType.StrType))
                            If DOMAIN = "" Then DOMAIN = Trim(CheckDBNull(dr(7), enumObjectType.StrType))
                            OFFR_ID = Trim(CheckDBNull(dr(8), enumObjectType.StrType))
                            SCORM_FLG = Trim(CheckDBNull(dr(9), enumObjectType.StrType))
                            JURIS_ID = Trim(CheckDBNull(dr(10), enumObjectType.StrType))
                            RESOLUTION = Trim(CheckDBNull(dr(11), enumObjectType.StrType))
                            KBA_REQD = Trim(CheckDBNull(dr(12), enumObjectType.StrType))
                            temp = Trim(CheckDBNull(dr(13), enumObjectType.StrType))
                            If temp = "" Then KBA_QUESTIONS = 0 Else KBA_QUESTIONS = Val(temp)
                            CRSE_ID = Trim(CheckDBNull(dr(14), enumObjectType.StrType))
                            EMAIL_ADDR = Trim(CheckDBNull(dr(15), enumObjectType.StrType))
                            ORDER_ID = Trim(CheckDBNull(dr(16), enumObjectType.StrType))
                            ORDER_STATUS = Trim(CheckDBNull(dr(17), enumObjectType.StrType))
                            COURSE = Trim(CheckDBNull(dr(18), enumObjectType.StrType))
                            REG_CONTACT_ID = Trim(CheckDBNull(dr(19), enumObjectType.StrType))
                        
                            If Debug = "Y" Then mydebuglog.Debug("  .. REG_ID/REG_STATUS_CD: " & REG_ID & " / " & REG_STATUS_CD)
                            Select Case REG_STATUS_CD
                                Case "Tentative"
                                    GoTo ProcStat
                                Case "Incomplete"
                                    GoTo ProcStat
                                Case "In Progress"
                                    GoTo ProcStat
                                Case "Accepted"
                                    GoTo ProcStat
                                Case "Exam Reqd"
                                    GoTo ProcStat
                                Case "On Hold"
                                    GoTo ProcStat
                            End Select
                            
                        End If
                    End While
                End If
                dr.Close()
            Catch ex As Exception
                GoTo DBError
            End Try
            
            ' Set SSLFlag based on course string
            If InStr(CRSE_CONTENT_URL, "https:") > 0 Then
                LaunchProtocol = "https:"
            End If

            ' If we made it to this point, check to see if we have multiple, and if so, bring up a list of them
            MSG1 = ""
            If TotalPgm > 1 Then
                If LANG_CD <> "ENU" Then
                    RegList = "https://www.gettips.com/mobile/" & LANG_CD & "/index.html?UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "#reg"
                Else
                    RegList = "https://www.gettips.com/mobile/index.html?UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "#reg"
                End If
                Select Case LANG_CD
                    Case "ESN"
                        MSG1 = "<br /><span class=""BigHeader""><font color=""Red"">Encontramos m&aacute;s de un registro para usted.<br /><br /><a href=""" & RegList & """ data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton5"" style=""width: 90%;"">Haga clic aqu&iacute; para ver una lista de sus registros</a></font></span>"
                    Case Else
                        MSG1 = "<br /><span class=""BigHeader""><font color=""Red"">We found more than one registration for you.<br /><br /><a href=""" & RegList & """ data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton5"" style=""width: 90%;"">Click here to view a list of your registrations</a></font></span>"
                End Select
            End If
   
ProcStat:
            If Debug = "Y" Then
                mydebuglog.Debug("  .. CRSE_CONTENT_URL: " & CRSE_CONTENT_URL)
                mydebuglog.Debug("  .. LaunchProtocol: " & LaunchProtocol)
                mydebuglog.Debug("  .. TRAIN_TYPE: " & TRAIN_TYPE)
                mydebuglog.Debug("  .. ORDER_ITEM_ID: " & ORDER_ITEM_ID)
                mydebuglog.Debug("  .. PAYMENT_REQD: " & PAYMENT_REQD)
                mydebuglog.Debug("  .. BL_PROD_INT_ID: " & BL_PROD_INT_ID)
                mydebuglog.Debug("  .. DOMAIN: " & DOMAIN)
                mydebuglog.Debug("  .. OFFR_ID: " & OFFR_ID)
                mydebuglog.Debug("  .. SCORM_FLG: " & SCORM_FLG)
                mydebuglog.Debug("  .. JURIS_ID: " & JURIS_ID)
                mydebuglog.Debug("  .. KBA_REQD: " & KBA_REQD)
                mydebuglog.Debug("  .. CRSE_ID: " & CRSE_ID)
                mydebuglog.Debug("  .. EMAIL_ADDR: " & EMAIL_ADDR)
                mydebuglog.Debug("  .. ORDER_ID: " & ORDER_ID)
                mydebuglog.Debug("  .. ORDER_STATUS: " & ORDER_STATUS)
                mydebuglog.Debug("  .. COURSE: " & COURSE)
            End If
            
            ' ================================================
            ' PROCESS BASED ON STATUS
            ' Registration on hold
            If REG_STATUS_CD = "On Hold" Then
                If LANG_CD <> "ENU" Then
                    ReviewReg = "https://www.gettips.com/mobile/" & LANG_CD & "/index.html?UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "#reg"
                    NextLink = "https://www.gettips.com/mobile/" & LANG_CD & "/index.html?UID=" & UID & "&SES=" & SessID & "&RG=" & REG_ID & "&PP=" & DOMAIN
                Else
                    ReviewReg = "https://www.gettips.com/mobile/index.html?UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "#reg"
                    NextLink = "https://www.gettips.com/mobile/index.html?UID=" & UID & "&SES=" & SessID & "&RG=" & REG_ID & "&PP=" & DOMAIN
                End If
                Select Case LANG_CD
                    Case "ESN"
                        MSG1 = MSG1 & "<br /><span class=""BigHeader"">Est&aacute; registrado en <i>" & COURSE & "</i> pero su certificaci&oacute;n no se puede completar debido a razones reglamentarias o de pago.<br /><br /><a href=""" & ReviewReg & """  data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton2"" style=""width: 90%;"">Haga clic aqu&iacute; para revisar sus registros</a>."
                        MSG1 = MSG1 & "<br /><br /><a href=""" & NextLink & """  data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton1"" style=""width: 90%;"">Haga clic para ir a su portal de entrenamiento personal</a>"
                    Case Else
                        MSG1 = MSG1 & "<br /><span class=""BigHeader"">You are registered for <i>" & COURSE & "</i> but your certification cannot be completed because of regulatory or payment reasons.<br /><br /><a href=""" & ReviewReg & """  data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton2"" style=""width: 90%;"">Click here to review your registrations</a>."
                        MSG1 = MSG1 & "<br /><br /><a href=""" & NextLink & """  data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton1"" style=""width: 90%;"">Click to go to your Personal Training Portal</a>"
                End Select
                MSG1 = MSG1 & "</span>"
                GoTo ReturnUser
            End If
            
            ' One or more unpaid registrations
            If REG_STATUS_CD = "Tentative" Then
                If PAYMENT_REQD = "Y" Then
                    If LANG_CD <> "ENU" Then
                        NextLink = "https://www.gettips.com/mobile/" & LANG_CD & "/index.html?UID=" & UID & "&SES=" & SessID & "&RG=" & REG_ID & "&PP=" & DOMAIN
                        ReviewReg = "https://www.gettips.com/mobile/" & LANG_CD & "/index.html?UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "#reg"
                    Else
                        NextLink = "https://www.gettips.com/mobile/index.html?UID=" & UID & "&SES=" & SessID & "&RG=" & REG_ID & "&PP=" & DOMAIN
                        ReviewReg = "https://www.gettips.com/mobile/index.html?UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "#reg"
                    End If
                    If ORDER_ITEM_ID = "" Then
                        If LANG_CD <> "ENU" Then
                            PaymentLink = "http://www.gettips.com/mobile/" & LANG_CD & "/orders.html?UID=" & UID & "&SES=" & SessID & "&OTY=REG&PP=TIPS&TRN=" & REG_ID
                        Else
                            PaymentLink = "http://www.gettips.com/mobile/orders.html?UID=" & UID & "&SES=" & SessID & "&OTY=REG&PP=TIPS&TRN=" & REG_ID
                        End If
                        Select Case LANG_CD
                            Case "ESN"
                                MSG1 = MSG1 & "<br /><span class=""BigHeader"">Encontramos un registro para <i>" & COURSE & "</i>, pero no lo ha pagado. " & _
                                "<br /><br /><a href=" & PaymentLink & " data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton4"" style=""width: 90%;"">Haga clic aqu&iacute; para realizar un pedido y pagar su inscripci&oacute;n</a><br /><br /><a href=""" & ReviewReg & """ data-role=""button"" rel=""external"" id=""navbutton2"" data-theme=""b"" style=""width: 90%;"">Haga clic aqu&iacute; para ir a su registro(s)</a>"
                            Case Else
                                MSG1 = MSG1 & "<br /><span class=""BigHeader"">We found a registration for <i>" & COURSE & "</i>, but you have not paid for it. " & _
                                "<br /><br /><a href=" & PaymentLink & " data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton4"" style=""width: 90%;"">Click here to place an order and pay for your registration</a><br /><br /><a href=""" & ReviewReg & """ data-role=""button"" rel=""external"" id=""navbutton2"" data-theme=""b"" style=""width: 90%;"">Click here to go to your registration(s)</a>"
                        End Select
                    Else
                        If ORDER_STATUS = "Paymt Reqd" Then
                            If LANG_CD <> "ENU" Then
                                PaymentLink = "https://www.gettips.com/mobile/" & LANG_CD & "/orders.html?UID=" & UID & "&SES=" & SessID & "&OID=" & ORDER_ID & "PP=" & DOMAIN
                            Else
                                PaymentLink = "https://www.gettips.com/mobile/orders.html?UID=" & UID & "&SES=" & SessID & "&OID=" & ORDER_ID & "PP=" & DOMAIN
                            End If
                            Select Case LANG_CD
                                Case "ESN"
                                    MSG1 = MSG1 & "<br /><span class=""BigHeader"">Est&aacute;s registrado en <i>" & COURSE & "</i> pero no has pagado." & _
                                    "<br /><br /><a href=" & PaymentLink & " data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton4"" style=""width: 90%;"">Haga clic aqu&iacute; para proporcionar informaci&oacute;n de pago</a><br /><br /><a href=""" & ReviewReg & """ data-role=""button"" rel=""external"" id=""navbutton2"" data-theme=""b"" style=""width: 90%;"">Haga clic aqu&iacute; para ir a su registro(s)</a>"
                                Case Else
                                    MSG1 = MSG1 & "<br /><span class=""BigHeader"">You are registered for <i>" & COURSE & "</i> but have not paid." & _
                                    "<br /><br /><a href=" & PaymentLink & " data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton4"" style=""width: 90%;"">Click here to provide payment information</a><br /><br /><a href=""" & ReviewReg & """ data-role=""button"" rel=""external"" id=""navbutton2"" data-theme=""b"" style=""width: 90%;"">Click here to go to your registration(s)</a>"
                            End Select
                        Else
                            If LANG_CD <> "ENU" Then
                                PaymentLink = "http://www.gettips.com/mobile/" & LANG_CD & "/orders.html?UID=" & UID & "&SES=" & SessID & "&OTY=REG&PP=TIPS&TRN=" & REG_ID
                            Else
                                PaymentLink = "http://www.gettips.com/mobile/orders.html?UID=" & UID & "&SES=" & SessID & "&OTY=REG&PP=TIPS&TRN=" & REG_ID
                            End If
                            Select Case LANG_CD
                                Case "ESN"
                                    MSG1 = MSG1 & "<br /><span class=""BigHeader"">Est&aacute;s registrado en <i>" & COURSE & "</i> pero debes proporcionar el pago para que este curso contin&uacute;e." & _
                                    "<br /><br /><a href=" & PaymentLink & " data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton4"" style=""width: 90%;"">Haga clic aqu&iacute; para proporcionar informaci&oacute;n de pago</a><br /><br /><a href=""" & ReviewReg & """ data-role=""button"" rel=""external"" id=""navbutton2"" data-theme=""b"" style=""width: 90%;"">Haga clic aqu&iacute; para ir a su registro(s)</a>"
                                Case Else
                                    MSG1 = MSG1 & "<br /><span class=""BigHeader"">You are registered for <i>" & COURSE & "</i> but must provide payment for this course to proceed." & _
                                    "<br /><br /><a href=" & PaymentLink & " data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton4"" style=""width: 90%;"">Click here to provide payment information</a><br /><br /><a href=""" & ReviewReg & """ data-role=""button"" rel=""external"" id=""navbutton2"" data-theme=""b"" style=""width: 90%;"">Click here to go to your registration(s)</a>"
                            End Select
               
                        End If
                    End If
                Else
                    ClassLink = ""
                    If LANG_CD <> "ENU" Then
                        ReviewReg = "https://www.gettips.com/mobile/" & LANG_CD & "/index.html?UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "#reg"
                        NextLink = "https://www.gettips.com/mobile/" & LANG_CD & "/index.html?UID=" & UID & "&SES=" & SessID & "&RG=" & REG_ID & "&PP=" & DOMAIN
                    Else
                        ReviewReg = "https://www.gettips.com/mobile/index.html?UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "#reg"
                        NextLink = "https://www.gettips.com/mobile/index.html?UID=" & UID & "&SES=" & SessID & "&RG=" & REG_ID & "&PP=" & DOMAIN
                    End If
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = MSG1 & "<br /><span class=""BigHeader"">Est&aacute;s registrado para <i>" & COURSE & "</i> pero hay un problema con este registro.<br /><br /><a href=""" & ReviewReg & """  data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton2"" style=""width: 90%;"">Haga clic aqu&iacute; para revisar sus registros</a>."
                        Case Else
                            MSG1 = MSG1 & "<br /><span class=""BigHeader"">You are registered for <i>" & COURSE & "</i> but there is an issue with this registration.<br /><br /><a href=""" & ReviewReg & """  data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton2"" style=""width: 90%;"">Click here to review your registrations</a>."
                    End Select
                End If
                Select Case LANG_CD
                    Case "ESN"
                        MSG1 = MSG1 & "<br /><br /><a href=""" & NextLink & """  data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton1"" style=""width: 90%;"">Haga clic para ir a su portal de entrenamiento personal</a>"
                    Case Else
                        MSG1 = MSG1 & "<br /><br /><a href=""" & NextLink & """  data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton1"" style=""width: 90%;"">Click to go to your Personal Training Portal</a>"
                End Select
                MSG1 = MSG1 & "</span>"
                GoTo ReturnUser
            End If
   
            ' User did not complete their registration
            If REG_STATUS_CD = "Incomplete" Then
                If LANG_CD <> "ENU" Then
                    NextLink = "https://www.gettips.com/mobile/" & LANG_CD & "/index.html?UID=" & UID & "&SES=" & SessID & "&RG=" & REG_ID & "&PP=" & DOMAIN
                    ReviewReg = "https://www.gettips.com/mobile/" & LANG_CD & "/index.html?UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "#reg"
                Else
                    NextLink = "https://www.gettips.com/mobile/index.html?UID=" & UID & "&SES=" & SessID & "&RG=" & REG_ID & "&PP=" & DOMAIN
                    ReviewReg = "https://www.gettips.com/mobile/index.html?UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "#reg"
                End If
                Select Case LANG_CD
                    Case "ESN"
                        MSG1 = "<span class=""BigHeader"">Est&aacute; en el proceso de tomar <i>" & COURSE & "</i> y se marc&oacute; como 'Incomplete'. <br /><br />Continuar el curso<br /><a href=""" & ReviewReg & """  data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton2"" style=""width: 90%;"">Haga clic aqu&iacute; para ir a su lista de clases</a>."
                        MSG1 = MSG1 & "<br /><br /><a href=""" & NextLink & """  data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton1"" style=""width: 90%;"">Haga clic para ir a su portal de entrenamiento personal</a>"
                    Case Else
                        MSG1 = "<span class=""BigHeader"">You are in the process of taking <i>" & COURSE & "</i> and it was marked 'Incomplete'. <br /><br />To continue the course<br /><a href=""" & ReviewReg & """  data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton2"" style=""width: 90%;"">Click here to go to your list of classes</a>."
                        MSG1 = MSG1 & "<br /><br /><a href=""" & NextLink & """  data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton1"" style=""width: 90%;"">Click to go to your Personal Training Portal</a>"
                End Select
      
                MSG1 = MSG1 & "</span>"
                GoTo ReturnUser
            End If
            
            ' Found registration ready to be taken   
            If (REG_STATUS_CD = "Accepted" Or REG_STATUS_CD = "In Progress") Then
                
                ' ----------
                ' CHECK KBA STATUS
                If KBA_REQD = "Y" Then
                    ' Check to see how many KBA questions exist for this jurisdiction
                    SqlS = "SELECT COUNT(*) " & _
                    "FROM elearning.dbo.KBA_QUES Q  " & _
                    "LEFT OUTER JOIN elearning.dbo.KBA_JURIS J ON J.QUES_ID=Q.ROW_ID  " & _
                    "LEFT OUTER JOIN elearning.dbo.KBA_CRSE C ON C.QUES_ID=Q.ROW_ID  " & _
                    "WHERE Q.ROW_ID IS NOT NULL AND C.CRSE_ID='" & CRSE_ID & "' AND J.JURIS_ID='" & JURIS_ID & "'"
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  COUNT NUMBER OF KBA QUESTIONS FOR A JURISDICTION: " & vbCrLf & "  " & SqlS)
                    cmd.CommandText = SqlS
                    Try
                        KBA_COUNT = CheckDBNull(cmd.ExecuteScalar(), enumObjectType.IntType)
                    Catch ex As Exception
                    End Try
                    mydebuglog.Debug("  .. KBA_COUNT : " & Str(KBA_COUNT))
                    
                    ' Check to see how many answered questions we have - if less than the number required then ask them
                    SqlS = "SELECT COUNT(*) " & _
                    "FROM elearning.dbo.KBA_ANSR A " & _
                    "WHERE A.REG_ID='" & REG_ID & "'"
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  CHECK TO SEE HOW MANY QUESTIONS ANSWERED: " & vbCrLf & "  " & SqlS)
                    cmd.CommandText = SqlS
                    Try
                        NUM_ANSRD = CheckDBNull(cmd.ExecuteScalar(), enumObjectType.IntType)
                    Catch ex As Exception
                    End Try
                    mydebuglog.Debug("  .. NUM_ANSRD : " & Str(NUM_ANSRD))
                    
                    ' Determine if we need to ask KBA questions
                    '      KBA_QUESTIONS = Number of questions required by the jurisdiction - "0" means all of the questions available
                    '      KBA_COUNT = Number of questions available for a jurisdiction
                    '      NUM_ANSRD = Number of questions answered by the student
                    '   Redirect to the KBA question screen if the number asked is less than the number required, and the number required is greater than zero
                    If KBA_QUESTIONS = 0 Then KBA_QUESTIONS = KBA_COUNT ' Set the number of questions to the jurisdiction count if zero
                    If NUM_ANSRD < KBA_QUESTIONS Then
                        If LANG_CD <> "ENU" Then
                            NextLink = "https://www.gettips.com/mobile/" & LANG_CD & "/index.html?UID=" & UID & "&SES=" & SessID & "&RG=" & REG_ID & "&PP=" & DOMAIN
                            ClassLink = LaunchProtocol & "//hciscorm.certegrity.com/ls/" & LANG_CD & "/OpenClass.html?RID=" & REG_ID & "&UID=" & UID & "&SES=" & SessID & "&LANG=" & LANG_CD & "&HP=gettips.com&PROT=http:&CUR=https://www.gettips.com/mobile/index.html"
                            RegLink = "https://www.gettips.com/mobile/" & LANG_CD & "/index.html?UID=" & UID & "&SES=" & SessID & "&RG=" & REG_ID & "&PP=" & DOMAIN & "#reg"
                        Else
                            NextLink = "https://www.gettips.com/mobile/index.html?UID=" & UID & "&SES=" & SessID & "&RG=" & REG_ID & "&PP=" & DOMAIN
                            ClassLink = LaunchProtocol & "//hciscorm.certegrity.com/ls/OpenClass.html?RID=" & REG_ID & "&UID=" & UID & "&SES=" & SessID & "&LANG=" & LANG_CD & "&HP=gettips.com&PROT=http:&CUR=https://www.gettips.com/mobile/index.html"
                            RegLink = "https://www.gettips.com/mobile/index.html?UID=" & UID & "&SES=" & SessID & "&RG=" & REG_ID & "&PP=" & DOMAIN & "#reg"
                        End If
                        Select Case LANG_CD
                            Case "ESN"
                                MSG1 = "<span class=""BigHeader""><br />Encontramos su registro " & TRAIN_TYPE & "." & _
                                "<br /><br />Si no se abre una ventana del navegador con su clase, puede tomar su clase <i>" & COURSE & "</i> haciendo clic en el bot&oacute;n ""Tomar Clase"" " & _
                                "on your class registration record.<br /><a href=""" & ClassLink & """ data-role=""button"" id=""navbutton3"" rel=""external"" data-theme=""b"" style=""width: 90%;"">Haga clic aqu&iacute; para abrir la clase</a><br /><br />" & _
                                "<a href=""" & RegLink & """ data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton6"" style=""width: 90%;"">Click here to go to your registration</a>"
                                If Not PortalLink Then
                                    MSG1 = MSG1 & "<br /><br /><a href=""" & NextLink & """ data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton1"" style=""width: 90%;"">Ve a tu Portal de Entrenamiento Personal</a></span>"
                                End If
                            Case Else
                                MSG1 = "<span class=""BigHeader""><br />We found your " & TRAIN_TYPE & " registration." & _
                                "<br /><br />If a browser window with your class fails to open, you may take your class <i>" & COURSE & "</i> by clicking on the ""Take Class"" button " & _
                                "on your class registration record.<br /><a href=""" & ClassLink & """ data-role=""button"" id=""navbutton3"" rel=""external"" data-theme=""b"" style=""width: 90%;"">Click here to open the class</a><br /><br />" & _
                                "<a href=""" & RegLink & """ data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton6"" style=""width: 90%;"">Click here to go to your registration</a>"
                                If Not PortalLink Then
                                    MSG1 = MSG1 & "<br /><br /><a href=""" & NextLink & """ data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton1"" style=""width: 90%;"">Go to your Personal Training Portal</a></span>"
                                End If
                        End Select
                        GoTo ReturnUser
                    End If
                End If
                
                ' ----------
                ' GENERATE OTHER LINKS
                If LANG_CD <> "ENU" Then
                    RegLink = "https://www.gettips.com/mobile/" & LANG_CD & "/index.html?UID=" & UID & "&SES=" & SessID & "&RG=" & REG_ID & "&PP=" & DOMAIN & "#reg"
                    NextLink = "https://www.gettips.com/mobile/" & LANG_CD & "/index.html?UID=" & UID & "&SES=" & SessID & "&RG=" & REG_ID & "&PP=" & DOMAIN
                    RegList = "https://www.gettips.com/mobile/" & LANG_CD & "/index.html?UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "#reg"
                    PaymentLink = "http://www.gettips.com/mobile/" & LANG_CD & "/orders.html?OTY=REG&PP=TIPS&TRN=" & REG_ID
                Else
                    RegLink = "https://www.gettips.com/mobile/index.html?UID=" & UID & "&SES=" & SessID & "&RG=" & REG_ID & "&PP=" & DOMAIN & "#reg"
                    NextLink = "https://www.gettips.com/mobile/index.html?UID=" & UID & "&SES=" & SessID & "&RG=" & REG_ID & "&PP=" & DOMAIN
                    RegList = "https://www.gettips.com/mobile/index.html?UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "#reg"
                    PaymentLink = "http://www.gettips.com/mobile/orders.html?OTY=REG&PP=TIPS&TRN=" & REG_ID
                End If
                If TotalPgm = 0 Then
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = "<span class=""BigHeader"">No se encontraron clases.<br /><br />"
                        Case Else
                            MSG1 = "<span class=""BigHeader"">No classes were found.<br /><br />"
                    End Select
                End If
                If TotalPgm = 1 Then
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = "<span class=""BigHeader"">Tu clase <i>" & COURSE & "</i> fue localizada.<br /><br />"
                        Case Else
                            MSG1 = "<span class=""BigHeader"">Your class <i>" & COURSE & "</i> was located.<br /><br />"
                    End Select
                End If
                If TotalPgm > 1 Then
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = "<span class=""BigHeader"">Tus clases fueron ubicadas.<br /><br />"
                        Case Else
                            MSG1 = "<span class=""BigHeader"">Your classes were located.<br /><br />"
                    End Select
                End If
                If TotalPgm > 0 Then
                    If RegFound Then
                        Select Case LANG_CD
                            Case "ESN"
                                MSG1 = MSG1 & "<a href=""" & RegList & """ data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton5"" style=""width: 90%;"">Haga clic aqu&iacute; para ir a su registro</a>"
                            Case Else
                                MSG1 = MSG1 & "<a href=""" & RegList & """ data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton5"" style=""width: 90%;"">Click here to go to your registration</a>"
                        End Select
                    Else
                        Select Case LANG_CD
                            Case "ESN"
                                MSG1 = MSG1 & "<a href=""" & RegList & """ data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton5"" style=""width: 90%;"">Haga clic aqu&iacute; para ir a una lista de todos sus registros</a>."
                            Case Else
                                MSG1 = MSG1 & "<a href=""" & RegList & """ data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton5"" style=""width: 90%;"">Click here to go to a list of all your registrations</a>."
                        End Select
                    End If
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = MSG1 & "<br /><br />Para volver a su clase despu&eacute;s de salir de ella, pero antes de completarla, consulte las instrucciones sobre c&oacute;mo hacerlo para enviarlas a <i>" & EMAIL_ADDR & "</i>."
                        Case Else
                            MSG1 = MSG1 & "<br /><br />To return to your class after exiting it, but before you complete it, please check for instructions on how to do so that will be sent to <i>" & EMAIL_ADDR & "</i>."
                    End Select
                Else
                    If LANG_CD <> "ENU" Then
                        NextLink = "https://www.gettips.com/mobile/" & LANG_CD & "/index.html?UID=" & UID & "&SES=" & SessID & "&RG=" & REG_ID & "&PP=" & DOMAIN & "#reg"
                        PaymentLink = "http://www.gettips.com/mobile/" & LANG_CD & "/orders.html?UID=&SES=&OTY=REG&PP=" & DOMAIN & "&TRN=" & REG_ID
                    Else
                        NextLink = "https://www.gettips.com/mobile/index.html?UID=" & UID & "&SES=" & SessID & "&RG=" & REG_ID & "&PP=" & DOMAIN & "#reg"
                        PaymentLink = "http://www.gettips.com/mobile/orders.html?UID=&SES=&OTY=REG&PP=" & DOMAIN & "&TRN=" & REG_ID
                    End If
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = MSG1 & "<br /><br /><a href=""https://www.gettips.com/mobile/" & LANG_CD & "/index.html?PP=" & DOMAIN & """ data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton1"" style=""width: 90%;"">Haga clic para ir a su portal de entrenamiento personal</a></span>"
                        Case Else
                            MSG1 = MSG1 & "<br /><br /><a href=""https://www.gettips.com/mobile/index.html?PP=" & DOMAIN & """ data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton1"" style=""width: 90%;"">Go to your Personal Training Portal</a></span>"
                    End Select
                End If
                
                ' ----------
                ' UPDATE REGISTRATION STATUS IF NECESSARY
                If REG_STATUS_CD = "Accepted" Then
                    SqlS = "UPDATE siebeldb.dbo.CX_SESS_REG SET STATUS_CD='In Progress', LAST_UPD=GETDATE() WHERE ROW_ID='" & REG_ID & "'"
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  UPDATE STATUS QUERY: " & vbCrLf & "  " & SqlS)
                    temp = ExecQuery("Update", "CX_SESS_REG", cmd, SqlS, mydebuglog, "N")
                End If
                
                ' ----------
                ' VERIFY LAST_INST IN SUB_CON_ID - NEEDED FOR REDIRECT BACK
                If CONTACT_ID <> "" Then
                    SqlS = "SELECT LAST_INST FROM siebeldb.dbo.CX_SUB_CON WHERE CON_ID='" & CONTACT_ID & "'"
                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  GET LAST_INST INFORMATION: " & vbCrLf & "  " & SqlS)
                    cmd.CommandText = SqlS
                    Try
                        LAST_INST = CheckDBNull(cmd.ExecuteScalar(), enumObjectType.StrType)
                    Catch ex As Exception
                    End Try
                    mydebuglog.Debug("  .. LAST_INST : " & LAST_INST)
                    
                    If LAST_INST = "" Then
                        If LANG_CD <> "ENU" Then
                            LAST_INST = "https://www.gettips.com/mobile/" & LANG_CD & "/index.html?UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "#reg"
                        Else
                            LAST_INST = "https://www.gettips.com/mobile/index.html?UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "#reg"
                        End If
                        SqlS = "UPDATE siebeldb.dbo.CX_SUB_CON SET LAST_INST='" & LAST_INST & "' WHERE CON_ID='" & CONTACT_ID & "'"
                        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  UPDATING LAST_INST: " & vbCrLf & "  " & SqlS)
                        temp = ExecQuery("Update", "CX_SUB_CON", cmd, SqlS, mydebuglog, "N")
                    End If
                End If
                
            Else
                If LANG_CD <> "ENU" Then
                    ClassLink = ""
                    NextLink = "https://www.gettips.com/mobile/" & LANG_CD & "/index.html?UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN
                    RegLink = "https://www.gettips.com/mobile/" & LANG_CD & "/register.html?UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN
                    ReviewReg = "https://www.gettips.com/mobile/" & LANG_CD & "/index.html?UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "#reg"
                Else
                    ClassLink = ""
                    NextLink = "https://www.gettips.com/mobile/index.html?UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN
                    RegLink = "https://www.gettips.com/mobile/register.html?UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN
                    ReviewReg = "https://www.gettips.com/mobile/index.html?UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "#reg"
                End If
                Select Case LANG_CD
                    Case "ESN"
                        MSG1 = MSG1 & "<span class=""BigHeader"">Actualmente no est&aacute;s registrado para tomar una clase. <br /> <br /> Recibir&aacute; este mensaje si a&uacute;n no se ha registrado o si ya ha completado el curso.<br /><a href=""" & RegLink & """  data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton6"" style=""width: 90%;"">Haga clic aqu&iacute; si desea registrarse </a> <br /> si cree que ya ha tomado el curso.<br /><a href=""" & ReviewReg & """  data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton2"" style=""width: 90%;"">Haga clic aqu&iacute; para revisar sus registros</a>."
                        If Not PortalLink Then
                            MSG1 = MSG1 & "<br /><a href=""" & NextLink & """  data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton1"" style=""width: 90%;"">Haga clic para ir a su portal de entrenamiento personal</a>"
                        End If
                    Case Else
                        MSG1 = MSG1 & "<span class=""BigHeader"">You are not currently registered to take a class. <br /><br />You will receive this message if you have not registered yet, or if you have already completed the course.<br /><a href=""" & RegLink & """  data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton6"" style=""width: 90%;"">Click here if you would like to register</a><br />If you believe you have already taken the course..<br /><a href=""" & ReviewReg & """  data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton2"" style=""width: 90%;"">Click here to review your registrations</a>."
                        If Not PortalLink Then
                            MSG1 = MSG1 & "<br /><a href=""" & NextLink & """  data-role=""button"" rel=""external"" data-theme=""b"" id=""navbutton1"" style=""width: 90%;"">Click to go to your Personal Training Portal</a>"
                        End If
                End Select
                MSG1 = MSG1 & "</span>"
            End If
            Refresh = "60"
   
        Else
            GoTo DBError
        End If
        
        ' ================================================
        ' RETURN TO USER
        ' This creates a frame which loads the content.  If the content is an exam, it also loads a Javascript library
        ' that handles the unload event
ReturnUser:
        outdata = ""
        outdata = outdata & """RegLink"":""" & RegLink & ""","
        outdata = outdata & """RegList"":""" & EscapeJSON(RegList) & ""","
        outdata = outdata & """PaymentLink"":""" & EscapeJSON(PaymentLink) & ""","
        outdata = outdata & """ClassLink"":""" & EscapeJSON(ClassLink) & ""","
        outdata = outdata & """ReviewReg"":""" & EscapeJSON(ReviewReg) & ""","
        outdata = outdata & """MSG1"":""" & EscapeJSON(MSG1) & ""","
        outdata = outdata & """Refresh"":""" & Refresh & ""","
        GoTo CloseOut
        
NotLoggedIn:
        If Debug = "Y" Then mydebuglog.Debug(">>NotLoggedIn")
        ErrLvl = "Warning"
        Select Case LANG_CD
            Case "ESN"
                errmsg = "No has iniciado sesi&oacute;n y debes hacerlo antes de continuar"
            Case Else
                errmsg = "You are not logged in and need to do so before proceeding"
        End Select
        GoTo DisplayError
        
DBError:
        If Debug = "Y" Then mydebuglog.Debug(">>DBError")
        Select Case LANG_CD
            Case "ESN"
                errmsg = "Se ha producido un error al acceder a los datos de clase."
            Case Else
                errmsg = "There was an error accessing class data."
        End Select
        GoTo DisplayError
   
DataError:
        If Debug = "Y" Then mydebuglog.Debug(">>DataError")
        ErrLvl = "Warning"
        Select Case LANG_CD
            Case "ESN"
                errmsg = "Se ha producido un error al acceder a los datos de tu clase."
            Case Else
                errmsg = "There was an error accessing  your class data."
        End Select
        GoTo DisplayError
   
AccessError:
        If Debug = "Y" Then mydebuglog.Debug(">>AccessError")
        ErrLvl = "Warning"
        Select Case LANG_CD
            Case "ESN"
                errmsg = "Hubo un error al intentar abrir tus clases."
            Case Else
                errmsg = "There was an error attempting to open your classes."
        End Select
   
DisplayError:
        If CURRENT_PAGE <> "" Then
            NextLink = CURRENT_PAGE
            If InStr(NextLink, "UID=") = 0 Then NextLink = NextLink & "?UID=" & UID
            If InStr(NextLink, "SES=") = 0 Then NextLink = NextLink & "&SES=" & SessID
            If InStr(NextLink, "#reg") = 0 Then NextLink = NextLink & "#reg"
        Else
            If LANG_CD <> "ENU" Then
                NextLink = "https://www.gettips.com/mobile/" & LANG_CD & "/index.html#reg"
            Else
                NextLink = "https://www.gettips.com/mobile/index.html" & "#reg"
            End If
        End If
        MSG1 = MSG1 & "<br /><br /><a href=" & NextLink & " data-role=""button"" id=""navbutton1"" rel=""external"" data-theme=""b"">Click to Continue</a>"
        
CloseOut:
        If Debug = "Y" Then
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & ">>Final Values")
            mydebuglog.Debug("  .. Refresh: " & Refresh)
            mydebuglog.Debug("  .. ReviewReg: " & ReviewReg)
            mydebuglog.Debug("  .. ClassLink: " & ClassLink)
            mydebuglog.Debug("  .. NextLink: " & NextLink)
            mydebuglog.Debug("  .. PaymentLink: " & PaymentLink)
            mydebuglog.Debug("  .. RegList: " & RegList)
            mydebuglog.Debug("  .. RegLink: " & RegLink)
            mydebuglog.Debug("  .. MSG1: " & MSG1)
            mydebuglog.Debug("  .. ErrMsg: " & errmsg)
        End If

        ' ============================================
        ' Finalize output      
        outdata = outdata & """NextLink"":""" & EscapeJSON(NextLink) & ""","
        outdata = outdata & """ErrMsg"":""" & errmsg & """ "
        jdoc = callback & "({""ResultSet"": {" & outdata & "} })"
        
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
        If Trim(errmsg) <> "" Then myeventlog.Error("WsGetClassAccess.ashx : " & ErrLvl & ": " & Trim(errmsg))
        myeventlog.Info("WsGetClassAccess.ashx : Contact Id: " & CONTACT_ID & ", Reg Id: " & REG_ID & ", Class Id: " & OFFR_ID & " - NextLink: " & NextLink)
        If Debug = "Y" Or (Logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("  outdata: " & outdata & vbCrLf)
                mydebuglog.Debug("Results:  Contact Id: " & CONTACT_ID & ", Reg Id: " & REG_ID & ", Class Id: " & OFFR_ID & " - NextLink: " & NextLink)
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
                LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, "WSGETCLASSACCESS", LogStartTime, VersionNum, Debug)
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