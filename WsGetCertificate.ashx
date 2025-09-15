<%@ WebHandler Language="VB" Class="WsGetCertificate" %>

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

Public Class WsGetCertificate : Implements IHttpHandler
    
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
        mydebuglog = log4net.LogManager.GetLogger("GetCertificateDebugLog")
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
        Dim EOL, LaunchProtocol, ReturnDestination, NextLink, outdata, ErrLvl As String
        Dim NewParticipation As Boolean
        Dim MSG1, CRSE_NAME, temp, HLINK, RefreshID As String
        Dim MSG2, CERT_LINK, SCORM_FLG, PASSED_FLG, TEST_STATUS, REG_STATUS_CD, TRAIN_TYPE As String
        Dim LOGGED_IN, CONTACT_ID, SUB_ID, CONTACT_OU_ID As String
        Dim EXAM_REQD, imgURL, FST_NAME, LAST_NAME, COMPLETE_DATE, ADDR, CITY, STATE, ZIP As String
        Dim CRSE_CONTENT_URL, CON_ID, SESS_PART_ID, EMAIL_ADDR, USERNAME, PASSWORD, CRSE_ID, SP_EXP_DATE, ERROR_MSG, COUNTRY As String
        Dim ContentLink, TYPE_CD, RESOLUTION, EDOMAIN, COURSE_TYPE_CD As String
        Dim TR_EXP_DATE, CURRCLM_PER_ID, DEF_THEME_ID, EXAM, RES_X, RES_Y As String
        
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
        ErrLvl = "Error"
        MSG1 = ""
        CRSE_NAME = ""
        REG_STATUS_CD = ""
        TEST_STATUS = ""
        PASSED_FLG = "N"
        SCORM_FLG = "Y"
        CERT_LINK = ""
        MSG2 = ""
        TRAIN_TYPE = ""
        CONTACT_ID = ""
        temp = ""
        HLINK = ""
        EXAM_REQD = ""
        imgURL = ""
        FST_NAME = ""
        LAST_NAME = ""
        COMPLETE_DATE = ""
        COURSE_TYPE_CD = ""
        ADDR = ""
        CITY = ""
        STATE = ""
        ZIP = ""
        ContentLink = ""
        CRSE_CONTENT_URL = ""
        CON_ID = ""
        SESS_PART_ID = ""
        EMAIL_ADDR = ""
        USERNAME = ""
        PASSWORD = ""
        CRSE_ID = ""
        SP_EXP_DATE = ""
        ERROR_MSG = ""
        COUNTRY = ""
        TYPE_CD = ""
        RESOLUTION = ""
        EDOMAIN = ""
        TR_EXP_DATE = ""
        CURRCLM_PER_ID = ""
        DEF_THEME_ID = ""
        EXAM = ""
        RES_X = ""
        RES_Y = ""
        RefreshID = ""

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
            tempdebug = System.Configuration.ConfigurationManager.AppSettings.Get("GetCertificate_debug")
            If tempdebug = "Y" And Debug <> "T" Then Debug = "Y"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            GoTo CloseOut2
        End Try
        
        ' ============================================
        ' Open log file if applicable
        If Logging = "Y" Then
            logfile = "C:\Logs\WsGetCertificate.log"
            Try
                log4net.GlobalContext.Properties("GetCertificateLogFileName") = logfile
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
            CRSE_TSTRUN_ID = context.Request.QueryString("TID")
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
 
        If Not context.Request.QueryString("RTN") Is Nothing Then
            RETURN_PAGE = context.Request.QueryString("RTN")
        End If
        
        If Not context.Request.QueryString("RFR") Is Nothing Then
            RefreshID = context.Request.QueryString("RFR")
        End If
 
        If Not context.Request.QueryString("PP") Is Nothing Then
            DOMAIN = context.Request.QueryString("PP")
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
        If DOMAIN = "" Then DOMAIN = "TIPS"
        If LANG_CD <> "ESN" And LANG_CD <> "ENU" Then LANG_CD = "ENU"
        If callback = "" Then callback = "?"
        If myprotocol = "" Then myprotocol = "http:"
        If InStr(1, PrevLink, "?UID") = 0 Then PrevLink = PrevLink & "?UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN
        PrevLink = Replace(PrevLink, "#reg", "")
        If Left(HOME_PAGE, 4) <> "web." And Left(HOME_PAGE, 4) <> "www." And HOME_PAGE <> "certegrity.com" Then
            If InStr(1, PrevLink, "web.") > 0 Then HOME_PAGE = "web." & HOME_PAGE Else HOME_PAGE = "www." & HOME_PAGE
        End If
        
        ReturnDestination = "#reg"
        If InStr(1, PrevLink, "FinishInstrument") > 0 Then ReturnDestination = "#cert"
        If InStr(1, PrevLink, "FinishAssessment") > 0 Then ReturnDestination = "#cert"
        If RETURN_PAGE <> "" Then ReturnDestination = "#" & LCase(RETURN_PAGE)
        
        If CURRENT_PAGE = "" Then
            If LANG_CD = "ENU" Then
                NextLink = myprotocol & "//" & HOME_PAGE & "/mobile/index.html" & "?UID=" & UID & "&SES=" & SessID & ReturnDestination
            Else
                NextLink = myprotocol & "//" & HOME_PAGE & "/mobile/" & LANG_CD & "/index.html" & "?UID=" & UID & "&SES=" & SessID & ReturnDestination
            End If
        Else
            NextLink = CURRENT_PAGE
            If InStr(NextLink, "UID=") = 0 Then NextLink = NextLink & "?UID=" & UID
            If InStr(NextLink, "SES=") = 0 Then NextLink = NextLink & "&SES=" & SessID
            If InStr(NextLink, "PP=") = 0 Then NextLink = NextLink & "&PP=" & DOMAIN
            If RefreshID <> "" Then
                NextLink = NextLink & "&RFR=" & RefreshID
            Else
                If InStr(NextLink, ReturnDestination) = 0 Then NextLink = NextLink & ReturnDestination
            End If
        End If
        
        If Debug = "Y" Then
            mydebuglog.Debug("----------------------------------")
            mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
            mydebuglog.Debug("Parameters-")
            mydebuglog.Debug("  Debug: " & Debug)
            mydebuglog.Debug("  UID: " & UID)
            mydebuglog.Debug("  cookieid: " & cookieid)
            mydebuglog.Debug("  SessID: " & SessID)
            mydebuglog.Debug("  CRSE_TSTRUN_ID: " & CRSE_TSTRUN_ID)
            mydebuglog.Debug("  REG_ID: " & REG_ID)
            mydebuglog.Debug("  PrevLink: " & PrevLink)
            mydebuglog.Debug("  ReturnDestination: " & ReturnDestination)
            mydebuglog.Debug("  NextLink: " & NextLink)
            mydebuglog.Debug("  myprotocol: " & myprotocol)
            mydebuglog.Debug("  HOME_PAGE: " & HOME_PAGE)
            mydebuglog.Debug("  RETURN_PAGE: " & RETURN_PAGE)
            mydebuglog.Debug("  RefreshID : " & RefreshID)
            mydebuglog.Debug("  CURRENT_PAGE: " & CURRENT_PAGE)
            mydebuglog.Debug("  LANG_CD: " & LANG_CD)
            mydebuglog.Debug("  DOMAIN: " & DOMAIN)
            mydebuglog.Debug("  BROWSER: " & BROWSER)
            mydebuglog.Debug("  callback: " & callback)
        End If
        
        If REG_ID = "" And CRSE_TSTRUN_ID = "" Then GoTo AccessError
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
            
            ' ================================================
            ' GET THE RESULTS OF THE CLASS AS REPORTED IN CX_SESS_REG
            If REG_ID <> "" Then
                SqlS = "SELECT R.STATUS_CD, T.DOMAIN, CR.X_CRSE_CONTENT_URL, SP.PASS_FLG, C.FST_NAME, C.LAST_NAME, " & _
                      "CR.NAME, FORMAT(SP.CREATED,'M/d/yyyy'), C.ROW_ID, SP.ROW_ID, C.EMAIL_ADDR, C.LOGIN, C.X_PASSWORD, CR.X_SCORM_FLG, " & _
                      "CR.ROW_ID, CT.STATUS_CD, CT.ROW_ID, FORMAT(SP.EXP_DATE,'M/d/yyyy'), CT.X_ERROR_MSG, " & _
                      "PA.ADDR, PA.CITY, PA.STATE, PA.ZIPCODE, PA.COUNTRY, " & _
                      "CA.ADDR, CA.CITY, CA.STATE, CA.ZIPCODE, CA.COUNTRY, CR.TYPE_CD, CR.X_EXAM_REQD, CR.X_RESOLUTION, " & _
                      "CR.X_SUMMARY_CD, CR.TYPE_CD " & _
                      "FROM siebeldb.dbo.CX_SESS_REG R " & _
                      "LEFT OUTER JOIN siebeldb.dbo.CX_TRAIN_OFFR T ON T.ROW_ID=R.TRAIN_OFFR_ID " & _
                      "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=R.CONTACT_ID " & _
                      "LEFT OUTER JOIN siebeldb.dbo.CX_SESS_PART_X SP ON SP.ROW_ID=R.SESS_PART_ID " & _
                      "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TSTRUN CT ON CT.ROW_ID=SP.CRSE_TSTRUN_ID " & _
                      "LEFT OUTER JOIN siebeldb.dbo.S_CRSE CR ON CR.ROW_ID=R.CRSE_ID " & _
                      "LEFT OUTER JOIN siebeldb.dbo.S_ADDR_PER PA ON PA.ROW_ID=C.PR_PER_ADDR_ID " & _
                      "LEFT OUTER JOIN siebeldb.dbo.S_ADDR_ORG CA ON CA.ROW_ID=C.PR_OU_ADDR_ID " & _
                      "WHERE R.ROW_ID='" & REG_ID & "' AND C.X_REGISTRATION_NUM='" & UID & "'"
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  REGISTRATION QUERY: " & vbCrLf & "  " & SqlS)
                Try
                    cmd.CommandText = SqlS
                    dr = cmd.ExecuteReader()
                    If Not dr Is Nothing Then
                        While dr.Read()                            
                            REG_STATUS_CD = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            If DOMAIN = "" Then DOMAIN = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            CRSE_CONTENT_URL = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                            If InStr(CRSE_CONTENT_URL, "https:") > 0 Then
                                LaunchProtocol = "https:"
                            End If
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
                            SCORM_FLG = Trim(CheckDBNull(dr(13), enumObjectType.StrType))
                            CRSE_ID = Trim(CheckDBNull(dr(14), enumObjectType.StrType))
                            TEST_STATUS = Trim(CheckDBNull(dr(15), enumObjectType.StrType))
                            CRSE_TSTRUN_ID = Trim(CheckDBNull(dr(16), enumObjectType.StrType))
                            SP_EXP_DATE = Trim(CheckDBNull(dr(17), enumObjectType.StrType))
                            ERROR_MSG = Trim(CheckDBNull(dr(18), enumObjectType.StrType))
                            TRAIN_TYPE = "Participant"
                            ADDR = Trim(CheckDBNull(dr(24), enumObjectType.StrType))
                            If ADDR = "" Then
                                ADDR = Trim(CheckDBNull(dr(19), enumObjectType.StrType))
                                CITY = Trim(CheckDBNull(dr(20), enumObjectType.StrType))
                                STATE = Trim(CheckDBNull(dr(21), enumObjectType.StrType))
                                ZIP = Trim(CheckDBNull(dr(22), enumObjectType.StrType))
                                COUNTRY = Trim(CheckDBNull(dr(23), enumObjectType.StrType))
                            Else
                                CITY = Trim(CheckDBNull(dr(25), enumObjectType.StrType))
                                STATE = Trim(CheckDBNull(dr(26), enumObjectType.StrType))
                                ZIP = Trim(CheckDBNull(dr(27), enumObjectType.StrType))
                                COUNTRY = Trim(CheckDBNull(dr(28), enumObjectType.StrType))
                            End If
                            If COUNTRY = "" Then COUNTRY = "USA"
                            TYPE_CD = Trim(CheckDBNull(dr(29), enumObjectType.StrType))
                            EXAM_REQD = Trim(CheckDBNull(dr(30), enumObjectType.StrType))
                            If EXAM_REQD = "" Then EXAM_REQD = "N"
                            RESOLUTION = Trim(CheckDBNull(dr(31), enumObjectType.StrType))
                            EDOMAIN = Trim(CheckDBNull(dr(32), enumObjectType.StrType))
                            If SCORM_FLG = "" Then SCORM_FLG = "N"
                            COURSE_TYPE_CD = Trim(CheckDBNull(dr(33), enumObjectType.StrType))
                        End While
                    End If
                Catch ex As Exception
                    GoTo AccessError
                End Try
                dr.Close()
                If Debug = "Y" Then
                    mydebuglog.Debug("  .. REG_STATUS_CD: " & REG_STATUS_CD)
                    mydebuglog.Debug("  .. CRSE_CONTENT_URL: " & CRSE_CONTENT_URL)
                    mydebuglog.Debug("  .. PASSED_FLG: " & PASSED_FLG)
                    mydebuglog.Debug("  .. FST_NAME: " & FST_NAME)
                    mydebuglog.Debug("  .. LAST_NAME: " & LAST_NAME)
                    mydebuglog.Debug("  .. CRSE_NAME: " & CRSE_NAME)
                    mydebuglog.Debug("  .. COMPLETE_DATE: " & COMPLETE_DATE)
                    mydebuglog.Debug("  .. CON_ID: " & CON_ID)
                    mydebuglog.Debug("  .. SESS_PART_ID: " & SESS_PART_ID)
                    mydebuglog.Debug("  .. EMAIL_ADDR: " & EMAIL_ADDR)
                    mydebuglog.Debug("  .. USERNAME: " & USERNAME)
                    mydebuglog.Debug("  .. PASSWORD: " & PASSWORD)
                    mydebuglog.Debug("  .. SCORM_FLG: " & SCORM_FLG)
                    mydebuglog.Debug("  .. CRSE_ID: " & CRSE_ID)
                    mydebuglog.Debug("  .. CRSE_TSTRUN_ID: " & CRSE_TSTRUN_ID)
                    mydebuglog.Debug("  .. SP_EXP_DATE: " & SP_EXP_DATE)
                    mydebuglog.Debug("  .. ERROR_MSG: " & ERROR_MSG)
                    mydebuglog.Debug("  .. TRAIN_TYPE: " & TRAIN_TYPE)
                    mydebuglog.Debug("  .. ADDR: " & ADDR)
                    mydebuglog.Debug("  .. CITY: " & CITY)
                    mydebuglog.Debug("  .. STATE: " & STATE)
                    mydebuglog.Debug("  .. ZIP: " & ZIP)
                    mydebuglog.Debug("  .. COUNTRY: " & COUNTRY)
                    mydebuglog.Debug("  .. TYPE_CD: " & TYPE_CD)
                    mydebuglog.Debug("  .. EXAM_REQD: " & EXAM_REQD)
                    mydebuglog.Debug("  .. RESOLUTION: " & RESOLUTION)
                    mydebuglog.Debug("  .. EDOMAIN: " & EDOMAIN)
                    mydebuglog.Debug("  .. COURSE_TYPE_CD: " & COURSE_TYPE_CD)
                End If
            End If

            ' Compute redirect
            If REG_ID <> "" And REG_STATUS_CD <> "" Then
                Select Case LANG_CD
                    Case "ESN"
                        ContentLink = LaunchProtocol & "//hciscorm.certegrity.com/ls/" & LANG_CD & "/OpenClass.html?RID=" & REG_ID & "&UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "&HP=" & HOME_PAGE & "&LANG=" & LANG_CD & "&CUR=" & CURRENT_PAGE
                    Case Else
                        ContentLink = LaunchProtocol & "//hciscorm.certegrity.com/ls/OpenClass.html?RID=" & REG_ID & "&UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "&HP=" & HOME_PAGE & "&CUR=" & CURRENT_PAGE
                End Select
                If Debug = "Y" Then mydebuglog.Debug("  .. ContentLink: " & ContentLink)
            End If
            
            ' ================================================
            ' GET THE RESULTS OF THE ASSESSMENT AS REPORTED IN S_CRSE_TSTRUN
            If CRSE_TSTRUN_ID <> "" And (REG_ID = "" or TEST_STATUS="") Then
                SqlS = "SELECT TR.STATUS_CD, S.DOMAIN, CR.X_CRSE_CONTENT_URL, TR.PASSED_FLG, C.FST_NAME, C.LAST_NAME, T.NAME,  " & _
                      "FORMAT(TR.GRADED_DT,'M/d/yyyy'), C.ROW_ID, SP.ROW_ID, C.EMAIL_ADDR, C.LOGIN, C.X_PASSWORD, CR.X_SCORM_FLG,  " & _
                      "CR.ROW_ID, FORMAT(SP.EXP_DATE,'M/d/yyyy'), FORMAT(CP.EXPIRATION_DT,'M/d/yyyy'), CP.ROW_ID, TR.X_ERROR_MSG, T.SKILL_LEVEL_CD, " & _
                      "T.X_DEF_THEME_ID, TR.STATUS_CD, PA.ADDR, PA.CITY, PA.STATE, PA.ZIPCODE, PA.COUNTRY, " & _
                      "CA.ADDR, CA.CITY, CA.STATE, CA.ZIPCODE, CA.COUNTRY, T.NAME, CR.X_EXAM_REQD, T.X_RESOLUTION, CR.X_SUMMARY_CD, CR.TYPE_CD " & _
                      "FROM siebeldb.dbo.S_CRSE_TSTRUN TR  " & _
                      "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=TR.PERSON_ID " & _
                      "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TST T ON T.ROW_ID=TR.CRSE_TST_ID  " & _
                      "LEFT OUTER JOIN siebeldb.dbo.S_CRSE CR ON CR.ROW_ID=T.CRSE_ID  " & _
                      "LEFT OUTER JOIN siebeldb.dbo.CX_SUB_CON SC ON SC.CON_ID=C.ROW_ID  " & _
                      "LEFT OUTER JOIN siebeldb.dbo.CX_SUBSCRIPTION S ON S.ROW_ID=SC.SUB_ID  " & _
                      "LEFT OUTER JOIN siebeldb.dbo.CX_SESS_PART_X SP ON SP.ROW_ID=TR.X_PART_ID " & _
                      "LEFT OUTER JOIN siebeldb.dbo.S_CURRCLM_PER CP ON CP.X_CRSE_TSTRUN_ID=TR.ROW_ID " & _
                      "LEFT OUTER JOIN siebeldb.dbo.S_ADDR_PER PA ON PA.ROW_ID=C.PR_PER_ADDR_ID " & _
                      "LEFT OUTER JOIN siebeldb.dbo.S_ADDR_ORG CA ON CA.ROW_ID=C.PR_OU_ADDR_ID " & _
                      "WHERE TR.ROW_ID='" & CRSE_TSTRUN_ID & "' AND C.X_REGISTRATION_NUM='" & UID & "' "
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  ASSESSMENT QUERY: " & vbCrLf & "  " & SqlS)
                cmd.CommandText = SqlS
                dr = cmd.ExecuteReader()
                If Not dr Is Nothing Then
                    While dr.Read()
                        Try
                            temp = ""
                            TEST_STATUS = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            If DOMAIN = "" Then DOMAIN = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
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
                            SCORM_FLG = Trim(CheckDBNull(dr(13), enumObjectType.StrType))
                            CRSE_ID = Trim(CheckDBNull(dr(14), enumObjectType.StrType))
                            SP_EXP_DATE = Trim(CheckDBNull(dr(15), enumObjectType.StrType))
                            TR_EXP_DATE = Trim(CheckDBNull(dr(16), enumObjectType.StrType))
                            CURRCLM_PER_ID = Trim(CheckDBNull(dr(17), enumObjectType.StrType))
                            ERROR_MSG = Trim(CheckDBNull(dr(18), enumObjectType.StrType))
                            TRAIN_TYPE = Trim(CheckDBNull(dr(19), enumObjectType.StrType))
                            DEF_THEME_ID = Trim(CheckDBNull(dr(20), enumObjectType.StrType))
                            TEST_STATUS = Trim(CheckDBNull(dr(21), enumObjectType.StrType))
                            ADDR = Trim(CheckDBNull(dr(27), enumObjectType.StrType))
                            If ADDR = "" Then
                                ADDR = Trim(CheckDBNull(dr(22), enumObjectType.StrType))
                                CITY = Trim(CheckDBNull(dr(23), enumObjectType.StrType))
                                STATE = Trim(CheckDBNull(dr(24), enumObjectType.StrType))
                                ZIP = Trim(CheckDBNull(dr(25), enumObjectType.StrType))
                                COUNTRY = Trim(CheckDBNull(dr(26), enumObjectType.StrType))
                            Else
                                CITY = Trim(CheckDBNull(dr(28), enumObjectType.StrType))
                                STATE = Trim(CheckDBNull(dr(29), enumObjectType.StrType))
                                ZIP = Trim(CheckDBNull(dr(30), enumObjectType.StrType))
                                COUNTRY = Trim(CheckDBNull(dr(31), enumObjectType.StrType))
                            End If
                            If COUNTRY = "" Then COUNTRY = "USA"
                            EXAM = Trim(CheckDBNull(dr(32), enumObjectType.StrType))
                            EXAM_REQD = Trim(CheckDBNull(dr(33), enumObjectType.StrType))
                            If EXAM_REQD = "" Then EXAM_REQD = "N"
                            RESOLUTION = Trim(CheckDBNull(dr(34), enumObjectType.StrType))
                            EDOMAIN = Trim(CheckDBNull(dr(35), enumObjectType.StrType))
                            If SCORM_FLG = "" Then SCORM_FLG = "N"
                            COURSE_TYPE_CD = Trim(CheckDBNull(dr(36), enumObjectType.StrType))
                            If COURSE_TYPE_CD = "TIPS Participant" Then REG_ID = ""
                            
                            ' Compute redirect
                            Select Case LANG_CD
                                Case "ESN"
                                    ContentLink = LaunchProtocol & "//hciscorm.certegrity.com/ls/" & LANG_CD & "/OpenClass.html?RID=" & REG_ID & "&UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "&HP=" & HOME_PAGE & "&LANG=" & LANG_CD & "&CUR=" & CURRENT_PAGE
                                Case Else
                                    ContentLink = LaunchProtocol & "//hciscorm.certegrity.com/ls/OpenClass.html?RID=" & REG_ID & "&UID=" & UID & "&SES=" & SessID & "&PP=" & DOMAIN & "&LANG=" & LANG_CD & "&HP=" & HOME_PAGE & "&CUR=" & CURRENT_PAGE
                            End Select
                            
                        Catch ex As Exception
                            MSG1 = ex.Message
                            GoTo DBError
                        End Try
                    End While
                Else
                    GoTo DataError
                End If
                dr.Close()
            End If
            
            ' Process Resolution
            If RESOLUTION = "" Then RESOLUTION = "800x600"
            RESOLUTION = Replace(RESOLUTION, "x", ",")
            RES_X = Trim(Str(Val(Left(RESOLUTION, InStr(1, RESOLUTION, ",") - 1)) - 20))
            RES_Y = Trim(Str(Val(Right(RESOLUTION, Len(RESOLUTION) - InStr(1, RESOLUTION, ","))) - 20))
            
            If Debug = "Y" Then
                mydebuglog.Debug("  .. REG_ID: " & REG_ID)
                mydebuglog.Debug("  .. TEST_STATUS: " & TEST_STATUS)
                mydebuglog.Debug("  .. SESS_PART_ID: " & SESS_PART_ID)
                mydebuglog.Debug("  .. REG_ID: " & REG_ID)
                mydebuglog.Debug("  .. CRSE_NAME: " & CRSE_NAME)
                mydebuglog.Debug("  .. TRAIN_TYPE: " & TRAIN_TYPE)
                mydebuglog.Debug("  .. ERROR_MSG: " & ERROR_MSG)
                mydebuglog.Debug("  .. EXAM_REQD: " & EXAM_REQD)
                mydebuglog.Debug("  .. PASSED_FLG: " & PASSED_FLG)
                mydebuglog.Debug("  .. RESOLUTION: " & RESOLUTION)
                mydebuglog.Debug("  .. RES_X: " & RES_X)
                mydebuglog.Debug("  .. RES_Y: " & RES_Y)
                mydebuglog.Debug("  .. EDOMAIN: " & EDOMAIN)
                mydebuglog.Debug("  .. REG_STATUS_CD: " & REG_STATUS_CD)
                mydebuglog.Debug("  .. TEST_STATUS: " & TEST_STATUS)
                mydebuglog.Debug("  .. CRSE_TSTRUN_ID: " & CRSE_TSTRUN_ID)
                mydebuglog.Debug("  .. SCORM_FLG: " & SCORM_FLG)
                mydebuglog.Debug("  .. COMPLETE_DATE: " & COMPLETE_DATE)
                mydebuglog.Debug("  .. ContentLink: " & ContentLink)
                mydebuglog.Debug("  .. COURSE_TYPE_CD: " & COURSE_TYPE_CD)
            End If
            
            ' ================================================
            ' CHECK FOR CERTIFICATION PRODUCT
            If CONTACT_ID <> "" Then
                If TRAIN_TYPE = "Trainer" Then
                    SqlS = "SELECT TOP 1 GENERATED R " & _
                        "FROM siebeldb.dbo.CX_CERT_PROD_RESULTS R " & _
                        "LEFT OUTER JOIN DMS.dbo.Documents D on D.row_id=R.DOC_ID "
                    SqlS = SqlS & "WHERE R.PROD_TYPE='T' AND R.CURRCLM_PER_ID='" & CRSE_TSTRUN_ID & "'"
                    SqlS = SqlS & " AND D.deleted is null AND D.row_id is not null"
                    SqlS = SqlS & " ORDER BY R.CREATED DESC"                    
                Else
                    SqlS = "SELECT TOP 1 GENERATED R " & _
                        "FROM siebeldb.dbo.CX_CERT_PROD_RESULTS R " & _
                        "LEFT OUTER JOIN DMS.dbo.Documents D on D.row_id=R.DOC_ID "
                    If REG_ID <> "" Then
                        SqlS = SqlS & "WHERE R.PROD_TYPE='R' AND R.REG_ID='" & REG_ID & "'"
                    Else
                        SqlS = SqlS & "WHERE R.PROD_TYPE IN ('T','R') AND R.CRSE_TSTRUN_ID='" & CRSE_TSTRUN_ID & "'"
                    End If
                    SqlS = SqlS & " AND D.deleted is null AND D.row_id is not null"
                    SqlS = SqlS & " ORDER BY R.CREATED DESC"
                End If
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  CERTIFICATION PRODUCT QUERY: " & vbCrLf & "  " & SqlS)
                cmd.CommandText = SqlS
                Try
                    CERT_LINK = CheckDBNull(cmd.ExecuteScalar(), enumObjectType.StrType)
                    If InStr(LCase(CERT_LINK), "http:") = 0 And InStr(LCase(CERT_LINK), "https:") = 0 Then CERT_LINK = ""
                    CERT_LINK = Replace(CERT_LINK, "http:", myprotocol)
                Catch ex As Exception
                    GoTo AccessError
                End Try
                If myprotocol = "https:" Then CERT_LINK = Replace(CERT_LINK, "http:", myprotocol)
                If Debug = "Y" Then mydebuglog.Debug("  .. CERT_LINK: " & CERT_LINK)
            End If
            
            ' ================================================
            ' DETERMINE NEXT STEP BASED ON REGISTRATION STATUS
            If Debug = "Y" Then
                mydebuglog.Debug("  DETERMINE NEXT STEP BASED ON REGISTRATION STATUS")
                mydebuglog.Debug("  .. REG_STATUS_CD: " & REG_STATUS_CD)
                mydebuglog.Debug("  .. TEST_STATUS: " & TEST_STATUS)
                mydebuglog.Debug("  .. REG_ID: " & REG_ID)
                mydebuglog.Debug("  .. PASSED_FLG: " & PASSED_FLG)
            End If
            If REG_STATUS_CD = "Completed" Or (TEST_STATUS = "Graded" And REG_ID <> "") Then GoTo GenerateCertificate
   
            ' Determine next step from status
            MSG1 = ""
            If REG_STATUS_CD <> "" Then
                Select Case LANG_CD
                    Case "ESN"
                        Select Case REG_STATUS_CD
                            Case "Completed"
                            Case "Exam Reqd"         ' Go to OpenInstrument agent
                                MSG1 = "<font color=""Gray"">Has completado tu curso y ahora debes tomar el examen de certificaci&oacute;n.<br><br>"
                                MSG1 = MSG1 & "<center><a href=" & LaunchProtocol & "//hciscorm.certegrity.com/ls/ESN/OpenAssessment.html?RID=" & REG_ID & "&UID=" & UID & "&SES=" & SessID & "&HP=" & HOME_PAGE & "&PP=" & DOMAIN & "&PROT=" & myprotocol & "&LANG=" & LANG_CD & "&CUR=" & CURRENT_PAGE & " id=""navbutton1"" data-role=""button"" rel=""external"" data-theme=""b"">Haga clic para pasar al examen de certificaci&oacute;n</a></center>"
                            Case "Retake"            ' Go to OpenSClass
                                MSG1 = "<font color=""Gray"">Puedes retomar el curso en cualquier momento.<br><br>"
                                MSG1 = MSG1 & "<center><a href=" & ContentLink & ">Haz click para tomar la clase</a></center>"
                            Case "Accepted"            ' Set status to "In Progress" - handle as "In Progress"
                                MSG1 = "<font color=""Gray"">No has completado este curso. Puede comenzar o continuar su curso en cualquier momento.<br><br>"
                                MSG1 = MSG1 & "H&aacute;galo ahora <a href=" & ContentLink & ">haciendo clic en este enlace</a>"
                            Case "In Progress"         ' Provide ability to re-enter or close course
                                MSG1 = "<font color=""Gray"">No has completado este curso. Puedes continuar tu curso en cualquier momento.<br><br>"
                                MSG1 = MSG1 & "Puede hacerlo ahora <a href=" & ContentLink & ">haciendo clic en este enlace</a>"
                            Case Else                  ' Error out
                                MSG1 = " <font color=""Gray"">Hubo un problema con tu curso. <br><br>P&oacute;ngase en contacto con nuestro Soporte T&eacute;cnico para obtener ayuda."
                        End Select

                    Case Else
                        Select Case REG_STATUS_CD
                            Case "Completed"
                            Case "Exam Reqd"
                                MSG1 = "<font color=""Gray"">You have completed your course, and must now take the certification exam.<br><br>"
                                MSG1 = MSG1 & "<center><a href=" & LaunchProtocol & "//hciscorm.certegrity.com/ls/OpenAssessment.html?RID=" & REG_ID & "&UID=" & UID & "&SES=" & SessID & "&HP=" & HOME_PAGE & "&PP=" & DOMAIN & "&PROT=" & myprotocol & "&LANG=" & LANG_CD & "&CUR=" & CURRENT_PAGE & " id=""navbutton1"" data-role=""button"" rel=""external"" data-theme=""b"">Click to proceed to the certification exam</a></center>"
                            Case "Retake"
                                MSG1 = "<font color=""Gray"">You may retake the course at any time.<br><br>"
                                MSG1 = MSG1 & "<center><a href=" & ContentLink & ">Click to take the class</a></center>"
                            Case "Accepted"            ' Set status to "In Progress" - handle as "In Progress"
                                MSG1 = "<font color=""Gray"">You have not completed this course.  You may begin or continue your course at any time.<br><br>"
                                MSG1 = MSG1 & "Do so now by <a href=" & ContentLink & ">Clicking on this link</a>"
                            Case "In Progress"         ' Provide ability to re-enter or close course
                                MSG1 = "<font color=""Gray"">You have not completed this course.  You may continue your course at any time.<br><br>"
                                MSG1 = MSG1 & "You may do so now by <a href=" & ContentLink & ">Clicking on this link</a>"
                            Case Else                  ' Error out
                                MSG1 = " <font color=""Gray"">There was a problem with your course. <br><br>Please contact our Technical Support department for assistance."
                        End Select
                End Select
            Else
                If TEST_STATUS = "Graded" And PASSED_FLG = "N" And REG_ID = "" Then
                    Select Case LANG_CD
                        Case "ESN"
                            MSG2 = "<font color=""Gray"">Su examen ha sido completado. Lamentablemente no pasaste.<br><br>"
                            MSG2 = MSG2 & "Por favor, consulte el correo electr&oacute;nico que se le envi&oacute; para obtener instrucciones especiales."
                            MSG1 = "<div class=""ui-grid-a"">" & _
                            "<div class=""ui-block-a""><a href=" & NextLink & " id=""navbutton1"" data-role=""button"" rel=""external"" data-theme=""b"">Volver al portal</a></div>" & _
                            "</div>"
                        Case Else
                            MSG2 = "<font color=""Gray"">Your exam has been completed.  Unfortunately you did not pass.<br><br>"
                            MSG2 = MSG2 & "Please see the email sent to you for any special instructions."
                            MSG1 = "<div class=""ui-grid-a"">" & _
                            "<div class=""ui-block-a""><a href=" & NextLink & " id=""navbutton1"" data-role=""button"" rel=""external"" data-theme=""b"">Return to the portal</a></div>" & _
                            "</div>"
                    End Select
                End If
            End If
            
            If TEST_STATUS <> "" And MSG1 = "" Then
                Select Case LANG_CD
                    Case "ESN"
                        Select Case UCase(TEST_STATUS)
                            Case "PENDING"
                                MSG1 = "<font color=""Gray"">No has completado tu examen. Puedes comenzar en cualquier momento.<br><br>"
                                MSG1 = MSG1 & "<center><a href=" & LaunchProtocol & "//hciscorm.certegrity.com/ls/ESN/OpenAssessment.html?RID=" & REG_ID & "&UID=" & UID & "&SES=" & SessID & "&HP=" & HOME_PAGE & "&PP=" & DOMAIN & "&PROT=" & myprotocol & "&LANG=" & LANG_CD & "&CUR=" & CURRENT_PAGE & " id=""navbutton1"" data-role=""button"" rel=""external"" data-theme=""b"">Haga clic para pasar al examen de certificaci&oacute;n</a></center>"
                            Case "UNSUBMITTED"
                                MSG1 = "<font color=""Gray"">No has completado tu examen. Puedes comenzar en cualquier momento.<br><br>"
                                MSG1 = MSG1 & "<center><a href=" & LaunchProtocol & "//hciscorm.certegrity.com/ls/ESN/OpenAssessment.html?RID=" & REG_ID & "&UID=" & UID & "&SES=" & SessID & "&HP=" & HOME_PAGE & "&PP=" & DOMAIN & "&PROT=" & myprotocol & "&LANG=" & LANG_CD & "&CUR=" & CURRENT_PAGE & " id=""navbutton1"" data-role=""button"" rel=""external"" data-theme=""b"">Haga clic para pasar al examen de certificaci&oacute;n</a></center>"
                            Case "SUBMITTED"
                                MSG1 = "<font color=""Gray"">Tu examen no ha sido calificado.<br><br>"
                                MSG1 = MSG1 & "Hazlo ahora por <a href=" & LaunchProtocol & "//hciscorm.certegrity.com/ls/ESN/OpenAssessment.html?RID=" & REG_ID & "&UID=" & UID & "&SES=" & SessID & "&HP=" & HOME_PAGE & "&PP=" & DOMAIN & "&PROT=" & myprotocol & "&LANG=" & LANG_CD & "&CUR=" & CURRENT_PAGE & " id=""navbutton1"" data-role=""button"" rel=""external"" data-theme=""b"">Haga clic para pasar al examen de certificaci&oacute;n</a>"
                            Case "UNGRADED"
                                MSG1 = "<font color=""Gray"">Your exam has not been score.<br><br>"
                                MSG1 = MSG1 & "Hazlo ahora por <a href=" & LaunchProtocol & "//hciscorm.certegrity.com/ls/ESN/OpenAssessment.html?RID=" & REG_ID & "&UID=" & UID & "&SES=" & SessID & "&HP=" & HOME_PAGE & "&PP=" & DOMAIN & "&PROT=" & myprotocol & "&LANG=" & LANG_CD & "&CUR=" & CURRENT_PAGE & " id=""navbutton1"" data-role=""button"" rel=""external"" data-theme=""b"">Haga clic para pasar al examen de certificaci&oacute;n</a>"
                            Case "GRADED"
                                If PASSED_FLG = "Y" Then GoTo GenerateCertificate
                                MSG1 = "<font color=""Gray"">Su examen ha sido completado. Lamentablemente no pasaste.<br><br>"
                                MSG1 = MSG1 & "Por favor, consulte el correo electr&oacute;nico que se le envi&oacute; para obtener instrucciones especiales."
                            Case "CANCELLED"
                                MSG1 = " <font color=""Gray"">Hubo un problema con su curso o puede haber un problema con nuestro sistema. <br><br>P&oacute;ngase en contacto con nuestro Departamento deSoporte T&eacute;cnico para asistencia."
                            Case "ERROR"
                                MSG1 = " <font color=""Gray"">Hubo un problema con su curso o puede haber un problema con nuestro sistema. <br><br>P&oacute;ngase en contacto con nuestro Departamento deSoporte T&eacute;cnico para asistencia."
                        End Select
         
                    Case Else
                        Select Case UCase(TEST_STATUS)
                            Case "PENDING"
                                MSG1 = "<font color=""Gray"">You have not completed your exam.  You may begin it at any time.<br><br>"
                                MSG1 = MSG1 & "<center><a href=" & LaunchProtocol & "//hciscorm.certegrity.com/ls/OpenAssessment.html?RID=" & REG_ID & "&UID=" & UID & "&SES=" & SessID & "&HP=" & HOME_PAGE & "&PP=" & DOMAIN & "&PROT=" & myprotocol & "&LANG=" & LANG_CD & "&CUR=" & CURRENT_PAGE & " id=""navbutton1"" data-role=""button"" rel=""external"" data-theme=""b"">Click to proceed to the certification exam</a></center>"
                            Case "UNSUBMITTED"
                                MSG1 = "<font color=""Gray"">You have not completed your exam.  You may begin it at any time.<br><br>"
                                MSG1 = MSG1 & "<center><a href=" & LaunchProtocol & "//hciscorm.certegrity.com/ls/OpenAssessment.html?RID=" & REG_ID & "&UID=" & UID & "&SES=" & SessID & "&HP=" & HOME_PAGE & "&PP=" & DOMAIN & "&PROT=" & myprotocol & "&LANG=" & LANG_CD & "&CUR=" & CURRENT_PAGE & " id=""navbutton1"" data-role=""button"" rel=""external"" data-theme=""b"">Click to proceed to the certification exam</a></center>"
                            Case "SUBMITTED"
                                MSG1 = "<font color=""Gray"">Your exam has not been scored.<br><br>"
                                MSG1 = MSG1 & "Do so now by <a href=" & LaunchProtocol & "//hciscorm.certegrity.com/ls/OpenAssessment.html?RID=" & REG_ID & "&UID=" & UID & "&SES=" & SessID & "&HP=" & HOME_PAGE & "&PP=" & DOMAIN & "&PROT=" & myprotocol & "&LANG=" & LANG_CD & "&CUR=" & CURRENT_PAGE & " id=""navbutton1"" data-role=""button"" rel=""external"" data-theme=""b"">Clicking on this link</a>"
                            Case "UNGRADED"
                                MSG1 = "<font color=""Gray"">Your exam has not been scored.<br><br>"
                                MSG1 = MSG1 & "Do so now by <a href=" & LaunchProtocol & "//hciscorm.certegrity.com/ls/OpenAssessment.html?RID=" & REG_ID & "&UID=" & UID & "&SES=" & SessID & "&HP=" & HOME_PAGE & "&PP=" & DOMAIN & "&PROT=" & myprotocol & "&LANG=" & LANG_CD & "&CUR=" & CURRENT_PAGE & " id=""navbutton1"" data-role=""button"" rel=""external"" data-theme=""b"">Clicking on this link</a>"
                            Case "GRADED"
                                If PASSED_FLG = "Y" Then GoTo GenerateCertificate
                                MSG1 = "<font color=""Gray"">Your exam has been completed.  Unfortunately you did not pass.<br><br>"
                                MSG1 = MSG1 & "Please see the email sent to you for any special instructions."
                            Case "CANCELLED"
                                MSG1 = " <font color=""Gray"">There was a problem with your course.  There may be an issue with our system. <br><br>Please contact our Technical Support department for assistance."
                            Case "ERROR"
                                MSG1 = " <font color=""Gray"">There was a problem with your course.  There may be an issue with our system. <br><br>Please contact our Technical Support department for assistance."
                        End Select
                End Select
            End If
            GoTo ReturnControl
            
GenerateCertificate:
            ' ================================================
            ' GENERATE RETURN BUTTON
            Select Case LANG_CD
                Case "ESN"
                    HLINK = "<a href=" & NextLink & " id=""navbutton1"" data-role=""button"" rel=""external"" data-theme=""b"">Haga clic para Volver al Portal</a>"
                    HLINK = HLINK & "&nbsp; &nbsp; <a href=""#"" value=""Print"" onClick=""javascript:window.print()"" id=""navbutton2"" data-role=""button"" rel=""external"" data-theme=""b"">Certificado de Impresi&oacute;n</a>"
                Case Else
                    HLINK = "<a href=" & NextLink & " id=""navbutton1"" data-role=""button"" rel=""external"" data-theme=""b"">Click to return to the portal</a>"
                    HLINK = HLINK & "&nbsp; &nbsp; <a href=""#"" value=""Print"" onClick=""javascript:window.print()"" id=""navbutton2"" data-role=""button"" rel=""external"" data-theme=""b"">Print Certificate</a>"
            End Select
            If Debug = "Y" Then mydebuglog.Debug("  .. HLINK: " & HLINK & vbCrLf)
   
            ' ================================================
            ' THE COURSE HAS BEEN COMPLETED, RETRIEVE ELEMENTS AND GENERATE A CERTIFICATE
            If PASSED_FLG = "Y" Or EXAM_REQD = "N" Then
                NextLink = Replace(NextLink, "#reg", "#cert")
                If Debug = "Y" Then
                    mydebuglog.Debug("  .. NextLink: " & NextLink)
                    mydebuglog.Debug("  .. CERT_LINK: " & CERT_LINK & vbCrLf)
                End If
                If CERT_LINK = "" Then
                    ' --------------------------------------------------
                    ' Generate Certificate of completion
                    If imgURL = "" Then imgURL = "/ls/images/"
                    imgURL = Replace(imgURL, "http:", myprotocol)
                    If DOMAIN = "TIPS" Then
                        If SCORM_FLG = "Y" Then
                            Select Case LANG_CD
                                Case "ESN"
                                    MSG1 = MSG2 & "<span class=""Title""><br/><br/><br/>" & EOL & _
                                    CRSE_NAME & "</span> " & EOL & _
                                    "<br/><br/><br/><br/><br/><br/><br/><span class=""Student"">" & FST_NAME & " " & LAST_NAME & "<br/><br/> " & EOL & _
                                    "Para "
                                    If REG_ID = "" Then
                                        MSG1 = MSG1 & "un examen"
                                    Else
                                        MSG1 = MSG1 & "cursos"
                                    End If
                                    MSG1 = MSG1 & " completado en " & COMPLETE_DATE & "</span><br/><br/> " & EOL & _
                                    "<span class=""CertInfo"">Los documentos de certificaci&oacute;n que se enviar&aacute;n a:<br/>" & EOL & _
                                    ADDR & ", " & CITY & ", " & STATE & " " & ZIP & EOL & _
                                    "</p>" & EOL

                                Case Else
                                    MSG1 = MSG2 & "<span class=""Title""><br/><br/><br/>" & EOL & _
                                    CRSE_NAME & "</span> " & EOL & _
                                    "<br/><br/><br/><br/><br/><br/><br/><span class=""Student"">" & FST_NAME & " " & LAST_NAME & "<br/><br/> " & EOL & _
                                    "For "
                                    If REG_ID = "" Then
                                        MSG1 = MSG1 & "an exam"
                                    Else
                                        MSG1 = MSG1 & "coursework"
                                    End If
                                    MSG1 = MSG1 & " completed on " & COMPLETE_DATE & "</span><br/><br/> " & EOL & _
                                    "<span class=""CertInfo"">Certification documents to be sent to:<br/>" & EOL & _
                                    ADDR & ", " & CITY & ", " & STATE & " " & ZIP & EOL & _
                                    "</p>" & EOL
                            End Select
               
                        Else
                            Select Case LANG_CD
                                Case "ESN"
                                    MSG2 = "<table width=""100%"" height=""100%"" cellpadding=""1"" cellspacing=""0""><tr align=""Center"" valign=""Center""><td colspan=2>" & _
                                    "<table cellpadding=0 cellspacing=0 width=779 border=0 bgcolor=""#f9f9f9"">" & _
                                    "<tr height=141>" & _
                                    "<td colspan=3><img src=" & imgURL & "Tcert1.gif border=0></td>" & _
                                    "</tr>" & _
                                    "<tr height=51 valign=""Center"">" & _
                                    "<td width=13><img src=" & imgURL & "Tcert2.gif border=0></td>" & _
                                    "<td width=752 align=""Center""><H2>" & CRSE_NAME & "</td>" & _
                                    "<td width=14><img src=" & imgURL & "Tcert3.gif border=0></td>" & _
                                    "</tr>" & _
                                    "<tr height=66>" & _
                                    "<td colspan=3 width=779><img src=" & imgURL & "Tcert4.gif border=0></td>" & _
                                    "</tr>" & _
                                    "<tr height=101 valign=""Center"">" & _
                                    "<td width=13><img src=" & imgURL & "Tcert7.gif border=0></td>" & _
                                    "<td width=752 align=""Center"" bgcolor=""#f9f9f9""><H3>" & FST_NAME & " " & LAST_NAME & "<br>Para "
                                    If CRSE_TSTRUN_ID <> "" Then
                                        MSG2 = MSG2 & "un examen"
                                    Else
                                        MSG2 = MSG2 & "cursos"
                                    End If
                                    MSG2 = MSG2 & " completado en " & COMPLETE_DATE & "</h3>" & _
                                    " Certificaci&oacute;n a ser enviada a: " & ADDR & ", " & CITY & ", " & STATE & " " & ZIP & "</td>" & _
                                    "<td width=14><img src=" & imgURL & "Tcert6.gif border=0></td>" & _
                                    "</tr>" & _
                                    "<tr height=209>"
                                Case Else
                                    MSG2 = "<table width=""100%"" height=""100%"" cellpadding=""1"" cellspacing=""0""><tr align=""Center"" valign=""Center""><td colspan=2>" & _
                                    "<table cellpadding=0 cellspacing=0 width=779 border=0 bgcolor=""#f9f9f9"">" & _
                                    "<tr height=141>" & _
                                    "<td colspan=3><img src=" & imgURL & "Tcert1.gif border=0></td>" & _
                                    "</tr>" & _
                                    "<tr height=51 valign=""Center"">" & _
                                    "<td width=13><img src=" & imgURL & "Tcert2.gif border=0></td>" & _
                                    "<td width=752 align=""Center""><H2>" & CRSE_NAME & "</td>" & _
                                    "<td width=14><img src=" & imgURL & "Tcert3.gif border=0></td>" & _
                                    "</tr>" & _
                                    "<tr height=66>" & _
                                    "<td colspan=3 width=779><img src=" & imgURL & "Tcert4.gif border=0></td>" & _
                                    "</tr>" & _
                                    "<tr height=101 valign=""Center"">" & _
                                    "<td width=13><img src=" & imgURL & "Tcert7.gif border=0></td>" & _
                                    "<td width=752 align=""Center"" bgcolor=""#f9f9f9""><H3>" & FST_NAME & " " & LAST_NAME & "<br>For "
                                    If CRSE_TSTRUN_ID <> "" Then
                                        MSG2 = MSG2 & "an exam"
                                    Else
                                        MSG2 = MSG2 & "coursework"
                                    End If
                                    MSG2 = MSG2 & " completed on " & COMPLETE_DATE & "</h3>" & _
                                    " Certification to be sent to: " & ADDR & ", " & CITY & ", " & STATE & " " & ZIP & "</td>" & _
                                    "<td width=14><img src=" & imgURL & "Tcert6.gif border=0></td>" & _
                                    "</tr>" & _
                                    "<tr height=209>"
                            End Select

                            If InStr(CRSE_NAME, "ETIPS") > 0 Then
                                MSG2 = MSG2 & "<td colspan=3><img src=" & imgURL & "Tcert5.gif border=0></td>"
                            Else
                                MSG2 = MSG2 & "<td colspan=3><img src=" & imgURL & "Tcert5g.gif border=0></td>"
                            End If
                            MSG2 = MSG2 & "</tr></table></td></tr></table>"
               
                            Select Case LANG_CD
                                Case "ESN"
                                    MSG1 = "<div class=""ui-grid-a"">" & _
                                    "<div class=""ui-block-a""><a href=" & NextLink & " id=""navbutton1"" data-role=""button"" rel=""external"" data-theme=""b"">Volver al portal</a></div>" & _
                                    "<div class=""ui-block-b""><a href=""#"" value=""Print"" onClick=""javascript:window.print()"" id=""navbutton2"" data-role=""button"" rel=""external"" data-theme=""b"">Certificado de Impresi&oacute;n</a></div>" & _
                                    "</div>"
                                Case Else
                                    MSG1 = "<div class=""ui-grid-a"">" & _
                                    "<div class=""ui-block-a""><a href=" & NextLink & " id=""navbutton1"" data-role=""button"" rel=""external"" data-theme=""b"">Return to the portal</a></div>" & _
                                    "<div class=""ui-block-b""><a href=""#"" value=""Print"" onClick=""javascript:window.print()"" id=""navbutton2"" data-role=""button"" rel=""external"" data-theme=""b"">Print Certificate</a></div>" & _
                                    "</div>"
                            End Select

                        End If
                    Else

                        Select Case LANG_CD
                            Case "ESN"
                                MSG2 = "<table width=""100%"" height=""100%""><tr align=""Center"" valign=""Center""><td>" & _
                                "<table cellpadding=0 cellspacing=0 width=787 border=0 bgcolor=""#f9f9f9"">" & _
                                "<tr height=142>" & _
                                "<td colspan=3><img src=" & imgURL & "cert1.gif border=0></td>" & _
                                "</tr><tr height=298>" & _
                                "<td width=22><img src=" & imgURL & "cert2.gif border=0></td>" & _
                                "<td width=743 align=""Center"" valign=""top"">" & _
                                "<br><h1>" & CRSE_NAME & "</h1><img src=" & imgURL & "cert8.gif border=0><br>" & _
                                "<h1>" & FST_NAME & " " & LAST_NAME & "</h1>" & _
                                "<br><br><h2>Para los cursos completados en " & COMPLETE_DATE & "</h2>" & _
                                "</td>" & _
                                "<td width=22><img src=" & imgURL & "cert3.gif border=0></td>" & _
                                "</tr>" & _
                                "<tr height=115>" & _
                                "<td width=22><img src=" & imgURL & "cert6.gif border=0></td>" & _
                                "<td width=743>&nbsp;</td>" & _
                                "<td width=22><img src=" & imgURL & "cert7.gif border=0></td>" & _
                                "</tr>" & _
                                "<tr height=24>" & _
                                "<td colspan=3><img src=" & imgURL & "cert4.gif border=0></td>" & _
                                "</tr></table></td></tr></table>"

                                MSG1 = "<div class=""ui-grid-a"">" & _
                                "<div class=""ui-block-a""><a href=" & NextLink & " id=""navbutton1"" data-role=""button"" rel=""external"" data-theme=""b"">Volver al portal</a></div>" & _
                                "<div class=""ui-block-b""><a href=""#"" value=""Print"" onClick=""javascript:window.print()"" id=""navbutton2"" data-role=""button"" rel=""external"" data-theme=""b"">Certificado de Impresi&oacute;n</a></div>" & _
                                "</div>"
                            Case Else
                                MSG2 = "<table width=""100%"" height=""100%""><tr align=""Center"" valign=""Center""><td>" & _
                                "<table cellpadding=0 cellspacing=0 width=787 border=0 bgcolor=""#f9f9f9"">" & _
                                "<tr height=142>" & _
                                "<td colspan=3><img src=" & imgURL & "cert1.gif border=0></td>" & _
                                "</tr><tr height=298>" & _
                                "<td width=22><img src=" & imgURL & "cert2.gif border=0></td>" & _
                                "<td width=743 align=""Center"" valign=""top"">" & _
                                "<br><h1>" & CRSE_NAME & "</h1><img src=" & imgURL & "cert8.gif border=0><br>" & _
                                "<h1>" & FST_NAME & " " & LAST_NAME & "</h1>" & _
                                "<br><br><h2>For coursework completed on " & COMPLETE_DATE & "</h2>" & _
                                "</td>" & _
                                "<td width=22><img src=" & imgURL & "cert3.gif border=0></td>" & _
                                "</tr>" & _
                                "<tr height=115>" & _
                                "<td width=22><img src=" & imgURL & "cert6.gif border=0></td>" & _
                                "<td width=743>&nbsp;</td>" & _
                                "<td width=22><img src=" & imgURL & "cert7.gif border=0></td>" & _
                                "</tr>" & _
                                "<tr height=24>" & _
                                "<td colspan=3><img src=" & imgURL & "cert4.gif border=0></td>" & _
                                "</tr></table></td></tr></table>"

                                MSG1 = "<div class=""ui-grid-a"">" & _
                                "<div class=""ui-block-a""><a href=" & NextLink & " id=""navbutton1"" data-role=""button"" rel=""external"" data-theme=""b"">Return to the portal</a></div>" & _
                                "<div class=""ui-block-b""><a href=""#"" value=""Print"" onClick=""javascript:window.print()"" id=""navbutton2"" data-role=""button"" rel=""external"" data-theme=""b"">Print Certificate</a></div>" & _
                                "</div>"
                        End Select

                    End If
                    If Debug = "Y" Then mydebuglog.Debug("  .. MSG2: " & MSG2 & vbCrLf)
                End If
            Else
                Select Case LANG_CD
                    Case "ESN"
                        MSG1 = "<center><table width=""800""><tr align=""Center"" valign=""middle"" BGCOLOR=""#000000""><td>" & _
                        "<table cellpadding=0 cellspacing=0 width=100% HEIGHT=100% border=0 BGCOLOR=""#f9f9f9"">" & _
                        "<tr align=""Center"" valign=""middle"">" & _
                        "<td>" & _
                        "<span class=""BigHeader"">Felicitaciones por completar el examen de certificaci&oacute;n. <br> Lo sentimos ... No pasaste.</span>" & _
                        "<br><br><span class=""BigHeader"">InstrLas subastas se le enviar&aacute;n por correo electr&oacute;nico sobre c&oacute;mo proceder. Gracias.</span>" & _
                        "<br><br><h3><a href=" & NextLink & " id=""navbutton1"" data-role=""button"" rel=""external"" data-theme=""b"">Haga clic para Continuar</a></h3></td></tr></table>" & _
                        "</td></tr></table></center>"
                    Case Else
                        MSG1 = "<center><table width=""800""><tr align=""Center"" valign=""middle"" BGCOLOR=""#000000""><td>" & _
                        "<table cellpadding=0 cellspacing=0 width=100% HEIGHT=100% border=0 BGCOLOR=""#f9f9f9"">" & _
                        "<tr align=""Center"" valign=""middle"">" & _
                        "<td>" & _
                        "<span class=""BigHeader"">Congratulations for completing the certification exam.<br>Sorry.. You did not pass.</span>" & _
                        "<br><br><span class=""BigHeader"">Instructions will be sent to you via email on how to proceed.  Thank you.</span>" & _
                        "<br><br><h3><a href=" & NextLink & " id=""navbutton1"" data-role=""button"" rel=""external"" data-theme=""b"">Click to Proceed</a></h3></td></tr></table>" & _
                        "</td></tr></table></center>"
                End Select
            End If
            
            ' ================================================
            ' PREPARE MSG1
            ' If a Certification Results product has been generated
            If CERT_LINK <> "" Then
                NextLink = Replace(NextLink, "#reg", "#cert")
                If Debug = "Y" Then mydebuglog.Debug("  .. NextLink: " & NextLink)
                Select Case LANG_CD
                    Case "ESN"
                        MSG1 = "<div class=""ui-grid-a"">" & _
                        "<div class=""ui-block-a""><a href=" & NextLink & " id=""navbutton1"" data-role=""button"" rel=""external"" data-theme=""b"">Volver al portal</a></div>" & _
                        "<div class=""ui-block-b""><a href=""#"" value=""Print"" onClick=""javascript:window.print()"" id=""navbutton2"" data-role=""button"" rel=""external"" data-theme=""b"">Certificado de Impresi&oacute;n</a></div>" & _
                        "</div>"
                    Case Else
                        MSG1 = "<div class=""ui-grid-a"">" & _
                        "<div class=""ui-block-a""><a href=" & NextLink & " id=""navbutton1"" data-role=""button"" rel=""external"" data-theme=""b"">Return to the portal</a></div>" & _
                        "<div class=""ui-block-b""><a href=""#"" value=""Print"" onClick=""javascript:window.print()"" id=""navbutton2"" data-role=""button"" rel=""external"" data-theme=""b"">Print Certificate</a></div>" & _
                        "</div>"
                End Select
      
                ' If no Certification Results product exists
            Else
                If SCORM_FLG = "Y" Then
                    temp = "<center><table width=""100%"" height=""100%"" border=""0"" cellpadding=""2"" cellspacing=""1"" bgcolor=""#f9f9f9"" style=""margin-top: -30px;""><tr valign=""Center""><td bgcolor=""#f9f9f9"" class=""body"">"
                    If CERT_LINK <> "" Then
                        temp = temp & "<center>" & HLINK & "</center>"
                    End If
                    temp = temp & "</td></tr>"
                    If PASSED_FLG = "Y" Then
                        If DOMAIN = "TIPS" Then
                            temp = temp & "<tr><td width=""100%"" height=""800"" style=""background: #f9f9f9 url(/ls/images/certificate.jpg) center no-repeat;"">"
                        Else
                            temp = temp & "<tr><td style=""background: #f9f9f9 url(/ls/images/certificate2.jpg) left no-repeat;"">"
                        End If
                    Else
                        temp = temp & "<tr><td width=""*"" height=""*"" class=""Header"">"
                    End If
                    temp = temp & "<div id=""info"" style=""color: #000000;"">"
                    temp = temp & MSG1
                    temp = temp & "</div></td></tr></table></center>"
                Else
                    temp = "<div id=""info"">"
                    temp = temp & MSG1
                    temp = temp & "</div>"
                End If
                MSG2 = temp
                MSG1 = ""

                If MSG2.IndexOf("navbutton1") = 0 Or PASSED_FLG = "Y" Then
                    Select Case LANG_CD
                        Case "ESN"
                            MSG1 = "<div class=""ui-grid-a"">" & _
                            "<div class=""ui-block-a""><a href=" & NextLink & " id=""navbutton1"" data-role=""button"" rel=""external"" data-theme=""b"">Volver al portal</a></div>" & _
                            "<div class=""ui-block-b""><a href=""#"" value=""Print"" onClick=""javascript:window.print()"" id=""navbutton2"" data-role=""button"" rel=""external"" data-theme=""b"">Certificado de Impresi&oacute;n</a></div>" & _
                            "</div>"
                        Case Else
                            MSG1 = "<div class=""ui-grid-a"">" & _
                            "<div class=""ui-block-a""><a href=" & NextLink & " id=""navbutton1"" data-role=""button"" rel=""external"" data-theme=""b"">Return to the portal</a></div>" & _
                            "<div class=""ui-block-b""><a href=""#"" value=""Print"" onClick=""javascript:window.print()"" id=""navbutton2"" data-role=""button"" rel=""external"" data-theme=""b"">Print Certificate</a></div>" & _
                            "</div>"
                    End Select
                End If
            End If
        Else
            errmsg = errmsg & "Unable to open the database connection. " & vbCrLf
        End If
        
        ' ================================================
        ' RETURN TO USER
        ' This creates a frame which loads the content.  If the content is an exam, it also loads a Javascript library
        ' that handles the unload event
ReturnControl:
        GoTo PrepareExit
        
NotFound:
        If Debug = "Y" Then mydebuglog.Debug(">>NotFound")
        ErrLvl = "Warning"
        Select Case LANG_CD
            Case "ESN"
                MSG2 = "Este " & TRAIN_TYPE & " no se encontr&oacute;"
            Case Else
                MSG2 = "This " & TRAIN_TYPE & " was not found"
        End Select
        GoTo PrepareExit
   
DBError:
        If Debug = "Y" Then mydebuglog.Debug(">>DBError")
        Select Case LANG_CD
            Case "ESN"
                MSG2 = "El sistema puede no estar disponible ahora. Por favor, int&eacute;ntelo de nuevo m&aacute;s tarde"
            Case Else
                MSG2 = "The system may be unavailable now.  Please try again later"
        End Select
        GoTo PrepareExit
   
AccessError:
        If Debug = "Y" Then mydebuglog.Debug(">>AccessError")
        ErrLvl = "Warning"
        Select Case LANG_CD
            Case "ESN"
                MSG2 = "Este certificado no es accesible para ti o est&aacute;s desconectado"
            Case Else
                MSG2 = "This certificate is not accessible to you or you are logged out"
        End Select
        GoTo PrepareExit
   
DataError:
        If Debug = "Y" Then mydebuglog.Debug(">>DataError")
        ErrLvl = "Warning"
        Select Case LANG_CD
            Case "ESN"
                MSG2 = "No se encontr&oacute; el registro del examen de certificaci&oacute;n para " & TRAIN_TYPE
            Case Else
                MSG2 = "The certification exam record for the " & TRAIN_TYPE & " was not found"
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
        If NextLink = "" Then
            If LANG_CD <> "ENU" Then
                NextLink = "https://www.gettips.com/mobile/" & LANG_CD & "/index.html?PP=" & DOMAIN & "#reg"
            Else
                NextLink = "https://www.gettips.com/mobile/index.html?PP=" & DOMAIN & "#reg"
            End If
        End If
        If CRSE_NAME = "" Then
            Select Case LANG_CD
                Case "ESN"
                    If DOMAIN = "TIPS" Then
                        CRSE_NAME = "eTIPS"
                    Else
                        CRSE_NAME = "Entrenamiento en linea"
                    End If
                Case Else
                    If DOMAIN = "TIPS" Then
                        CRSE_NAME = "eTIPS"
                    Else
                        CRSE_NAME = "OnLine Training"
                    End If
            End Select
        End If

        If MSG2.IndexOf("navbutton1") = 0 Then
            Select Case LANG_CD
                Case "ESN"
                    MSG1 = "<div class=""ui-grid-a"">" & _
                    "<div class=""ui-block-a""><a href=" & NextLink & " id=""navbutton1"" data-role=""button"" rel=""external"" data-theme=""b"">Volver al portal</a></div>" & _
                    "</div>"
                Case Else
                    MSG1 = "<div class=""ui-grid-a"">" & _
                    "<div class=""ui-block-a""><a href=" & NextLink & " id=""navbutton1"" data-role=""button"" rel=""external"" data-theme=""b"">Return to the portal</a></div>" & _
                    "</div>"
            End Select
        End If
        
CloseOut:
        If MSG1 = "" And InStr(MSG2, NextLink) > 0 Then
            NextLink = ""
        End If
        If Debug = "Y" Then
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & ">>Final Values")
            mydebuglog.Debug("  .. REG_STATUS_CD: " & REG_STATUS_CD)
            mydebuglog.Debug("  .. TEST_STATUS: " & TEST_STATUS)
            mydebuglog.Debug("  .. CRSE_NAME: " & CRSE_NAME)
            mydebuglog.Debug("  .. NextLink: " & NextLink)
            mydebuglog.Debug("  .. DOMAIN: " & DOMAIN)
            mydebuglog.Debug("  .. PASSED_FLG: " & PASSED_FLG)
            mydebuglog.Debug("  .. SCORM_FLG: " & SCORM_FLG)
            mydebuglog.Debug("  .. CERT_LINK: " & CERT_LINK)
            mydebuglog.Debug("  .. MSG1: " & MSG1)
            mydebuglog.Debug("  .. MSG2: " & MSG2)
            mydebuglog.Debug("  .. ErrMsg: " & errmsg)
        End If

        ' ============================================
        ' Finalize output      
        outdata = ""
        outdata = outdata & """REG_STATUS_CD"":""" & REG_STATUS_CD & ""","
        outdata = outdata & """TEST_STATUS"":""" & EscapeJSON(TEST_STATUS) & ""","
        outdata = outdata & """CRSE_NAME"":""" & EscapeJSON(CRSE_NAME) & ""","
        outdata = outdata & """NextLink"":""" & EscapeJSON(NextLink) & ""","
        outdata = outdata & """DOMAIN"":""" & EscapeJSON(DOMAIN) & ""","
        outdata = outdata & """PASSED_FLG"":""" & PASSED_FLG & ""","
        outdata = outdata & """SCORM_FLG"":""" & SCORM_FLG & ""","
        outdata = outdata & """CERT_LINK"":""" & EscapeJSON(CERT_LINK) & ""","
        outdata = outdata & """MSG1"":""" & EscapeJSON(MSG1) & ""","
        outdata = outdata & """MSG2"":""" & EscapeJSON(MSG2) & ""","
        outdata = outdata & """ErrMsg"":""" & errmsg & """ "
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
        If Trim(errmsg) <> "" Then myeventlog.Error("WsGetCertificate.ashx: " & ErrLvl & ": " & Trim(errmsg))
        myeventlog.Info("WsGetCertificate.ashx : Contact Id: " & CONTACT_ID & ", Reg Id: " & REG_ID & " - ContentLink: " & NextLink)
        If Debug = "Y" Or (Logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("  outdata: " & outdata & vbCrLf)
                mydebuglog.Debug("Results:  Contact Id: " & CONTACT_ID & ", Reg Id: " & REG_ID & " - ContentLink: " & NextLink)
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
                LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, "WSGETCERTIFICATE", LogStartTime, VersionNum, Debug)
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