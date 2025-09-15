<%@ WebHandler Language="VB" Class="WsGetDocument" %>

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
Imports Amazon
Imports Amazon.S3

Public Class WsGetDocument : Implements IHttpHandler

    Enum enumObjectType
        StrType = 0
        IntType = 1
        DblType = 2
        DteType = 3
    End Enum

    Public Class clsSSL
        Public Function AcceptAllCertifications(ByVal sender As Object, ByVal certification As System.Security.Cryptography.X509Certificates.X509Certificate, ByVal chain As System.Security.Cryptography.X509Certificates.X509Chain, ByVal sslPolicyErrors As System.Net.Security.SslPolicyErrors) As Boolean
            Return True
        End Function
    End Class

    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest

        ' Parameter Declarations
        Dim Debug, RETURN_PAGE, DOMAIN, DMID As String
        Dim UID, SessID, CURRENT_PAGE, HOME_PAGE, LANG_CD, callback, myprotocol As String

        ' Result Declarations
        Dim jdoc As String

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String

        ' Minio declarations
        Dim minio_flg, d_verid As String
        Dim AccessKey, AccessSecret, AccessRegion, AccessBucket As String
        Dim sslhttps As clsSSL = New clsSSL

        ' Logging declarations
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("GetDocumentDebugLog")
        Dim logfile, tempdebug As String
        Dim Logging As String
        Dim LogStartTime As String = Now.ToString
        Dim VersionNum As String = "101"
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
        Dim cookieid, cookiesessid As String
        cookiesessid = ""
        cookieid = ""
        If context.Request.Cookies.Count > 0 Then
            cookieid = Trim(context.Request.Cookies.Item("ID").Value.ToString())
            cookiesessid = Trim(context.Request.Cookies.Item("Sess").Value.ToString())
        End If

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.basic.Service
        Dim Processing As New com.certegrity.cloudsvc.processing.Service

        ' Variable declarations
        Dim errmsg As String                    ' Error message (if any)
        Dim EOL, LaunchProtocol, ReturnDestination, NextLink, outdata, ErrLvl As String
        Dim NewParticipation As Boolean
        Dim MSG1, CRSE_NAME, temp, HLINK As String
        Dim MSG2, CERT_LINK, SCORM_FLG, PASSED_FLG, TEST_STATUS, REG_STATUS_CD, TRAIN_TYPE As String
        Dim LOGGED_IN, CONTACT_ID, SUB_ID, CONTACT_OU_ID, ltemp, USER_DOMAIN As String
        Dim DOC_FN, VERSION_ID, dName, newwindow, FileName, ReDirect, PublicKey, ItemName, isFirefox As String

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
        USER_DOMAIN = ""
        LANG_CD = "ENU"
        callback = ""
        myprotocol = ""
        HOME_PAGE = ""
        RETURN_PAGE = ""
        CURRENT_PAGE = ""
        errmsg = ""
        EOL = Chr(10)
        ltemp = ""
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
        newwindow = ""
        FileName = ""
        ReDirect = ""
        DOC_FN = ""
        VERSION_ID = ""
        dName = ""
        DMID = ""
        PublicKey = ""
        ItemName = ""
        isFirefox = ""
        minio_flg = ""
        d_verid = ""

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
            tempdebug = System.Configuration.ConfigurationManager.AppSettings.Get("GetDocument_debug")
            If tempdebug = "Y" And Debug <> "T" Then Debug = "Y"
            AccessKey = System.Configuration.ConfigurationManager.AppSettings("minio-key")
            If AccessKey = "" Then AccessKey = "dms"
            AccessSecret = System.Configuration.ConfigurationManager.AppSettings("minio-secret")
            If AccessSecret <> "" Then AccessSecret = System.Web.HttpUtility.HtmlDecode(AccessSecret)
            If AccessSecret = "" Then AccessSecret = "TptbjrNTVQDRYFJzNmw27BV5"
            AccessRegion = System.Configuration.ConfigurationManager.AppSettings("minio-region")
            If AccessRegion = "" Then AccessRegion = "us-east"
            AccessBucket = System.Configuration.ConfigurationManager.AppSettings("minio-bucket")
            If AccessBucket = "" Then AccessBucket = "dms"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open log file if applicable
        If Logging = "Y" Then
            logfile = "C:\Logs\WsGetDocument.log"
            Try
                log4net.GlobalContext.Properties("GetDocumentLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                GoTo CloseOut2
            End Try
        End If

        ' ============================================
        ' Get parameters    
        If Not context.Request.QueryString("DMI") Is Nothing Then
            DMID = context.Request.QueryString("DMI")
        End If

        If Not context.Request.QueryString("PP") Is Nothing Then
            DOMAIN = context.Request.QueryString("PP")
        End If

        If Not context.Request.QueryString("UID") Is Nothing Then
            UID = context.Request.QueryString("UID")
            If InStr(UID, ",") > 0 Then
                If cookieid <> "" Then UID = cookieid
            End If
        End If

        If Not context.Request.QueryString("SES") Is Nothing Then
            SessID = context.Request.QueryString("SES")
            If InStr(SessID, ",") > 0 Then
                If cookiesessid<>"" THEN SessID = cookiesessid
            End If
        End If

        If Not context.Request.QueryString("HP") Is Nothing Then
            HOME_PAGE = context.Request.QueryString("HP")
        End If

        If Not context.Request.QueryString("FF") Is Nothing Then
            isFirefox = context.Request.QueryString("FF")
        End If

        If Not context.Request.QueryString("LANG") Is Nothing Then
            LANG_CD = UCase(context.Request.QueryString("LANG"))
        End If

        If Not context.Request.QueryString("CUR") Is Nothing Then
            CURRENT_PAGE = context.Request.QueryString("CUR")
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
        myprotocol = "https:"
        If myprotocol = "" Then myprotocol = "https:"
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
            mydebuglog.Debug("  cookiesessid: " & cookiesessid)
            mydebuglog.Debug("  DMID: " & DMID)
            mydebuglog.Debug("  myprotocol: " & myprotocol)
            mydebuglog.Debug("  HOME_PAGE: " & HOME_PAGE)
            mydebuglog.Debug("  RETURN_PAGE: " & RETURN_PAGE)
            mydebuglog.Debug("  CURRENT_PAGE: " & CURRENT_PAGE)
            mydebuglog.Debug("  LANG_CD: " & LANG_CD)
            mydebuglog.Debug("  DOMAIN: " & DOMAIN)
            mydebuglog.Debug("  BROWSER: " & BROWSER)
            mydebuglog.Debug("  isFirefox: " & isFirefox)
            mydebuglog.Debug("  callback: " & callback)
        End If

        If DMID = "" Then GoTo AvailError
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

            ' Validate access
            If UID <> "" And SessID <> "" Then
                SqlS = "SELECT TOP 1 (SELECT CASE WHEN H.LOGOUT_DT IS NULL THEN 'Y' ELSE 'N' END) AS LOGGED_IN, SC.CON_ID, SC.SUB_ID, C.PR_DEPT_OU_ID, S.DOMAIN, C.X_PR_LANG_CD " & _
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
                            USER_DOMAIN = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                            ltemp = Trim(CheckDBNull(dr(5), enumObjectType.StrType))
                            If LANG_CD <> ltemp And ltemp <> "" Then LANG_CD = ltemp
                        End While
                    End If
                Catch ex As Exception
                    GoTo AccessError
                End Try
                dr.Close()
                If Debug = "Y" Then
                    mydebuglog.Debug("  .. LOGGED_IN: " & LOGGED_IN)
                    mydebuglog.Debug("  .. CONTACT_ID: " & CONTACT_ID)
                    mydebuglog.Debug("  .. LANG_CD: " & LANG_CD)
                    mydebuglog.Debug("  .. USER_DOMAIN: " & USER_DOMAIN & vbCrLf)
                End If
                If LOGGED_IN <> "Y" Then GoTo AccessError
            End If

            ' ================================================
            ' GET DOCUMENT PATH AND VERIFY THAT THE DOCUMENT EXISTS AND IS IN A PUBLIC GROUP
            If DMID <> "" Then
                If UID = "" Then
                    SqlS = "SELECT D.dfilename, V.version, D.name, V.minio_flg, V.row_id, D.row_id " & _
                        "FROM DMS.dbo.Documents D " & _
                        "LEFT OUTER JOIN DMS.dbo.Document_Versions V ON V.row_id=D.last_version_id " & _
                        "LEFT OUTER JOIN DMS.dbo.Document_Categories DC ON DC.doc_id=D.row_id " & _
                        "LEFT OUTER JOIN DMS.dbo.Categories C ON C.row_id=DC.cat_id " & _
                        "INNER JOIN DMS.dbo.Document_Users DU ON DU.doc_id=D.row_id " & _
                        "LEFT OUTER JOIN DMS.dbo.User_Group_Access UA ON UA.row_id=DU.user_access_id " & _
                        "LEFT OUTER JOIN DMS.dbo.Groups G1 ON G1.row_id=UA.access_id AND UA.type_id='G' AND G1.type_cd='Domain' " & _
                        "LEFT OUTER JOIN DMS.dbo.Category_Access CA ON CA.cat_id=C.row_id AND CA.user_access_id=UA.row_id " & _
                        "WHERE D.row_id='" & DMID & "' AND D.deleted IS NULL AND " & _
                        "G1.name='" & DOMAIN & "' AND CHARINDEX('P',CA.access_type)>0 " & _
                        "GROUP BY D.dfilename, V.version, D.name"
                Else
                    If USER_DOMAIN <> DOMAIN Then
                        temp = "G1.name IN ('" & DOMAIN & "','" & USER_DOMAIN & "')"
                    Else
                        If DOMAIN <> "TIPS" Then
                            temp = "G1.name IN ('" & DOMAIN & "','TIPS')"
                        Else
                            temp = "G1.name='" & USER_DOMAIN & "'"
                        End If
                    End If
                    SqlS = "SELECT D.dfilename, V.version, D.name, V.minio_flg, V.row_id, D.row_id " &
                        "FROM DMS.dbo.Documents D " &
                        "LEFT OUTER JOIN DMS.dbo.Document_Versions V ON V.row_id=D.last_version_id " &
                        "LEFT OUTER JOIN DMS.dbo.Document_Categories DC ON DC.doc_id=D.row_id " &
                        "LEFT OUTER JOIN DMS.dbo.Categories C ON C.row_id=DC.cat_id " &
                        "INNER JOIN DMS.dbo.Document_Users DU ON DU.doc_id=D.row_id " &
                        "LEFT OUTER JOIN DMS.dbo.User_Group_Access UA ON UA.row_id=DU.user_access_id " &
                        "LEFT OUTER JOIN DMS.dbo.Groups G1 ON G1.row_id=UA.access_id AND UA.type_id='G' AND G1.type_cd='Domain' " &
                        "LEFT OUTER JOIN DMS.dbo.Category_Access CA ON CA.cat_id=C.row_id AND CA.user_access_id=UA.row_id " &
                        "WHERE D.row_id='" & DMID & "' AND D.deleted IS NULL AND " & temp & " AND " &
                        "(CHARINDEX('P',CA.access_type)>0 OR CHARINDEX('R',DU.access_type)>0) " &
                        "GROUP BY D.dfilename, V.version, D.name, V.minio_flg, V.row_id, D.row_id"
                End If
                If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Get the Document Path: " & vbCrLf & "  " & SqlS)
                Try
                    cmd.CommandText = SqlS
                    dr = cmd.ExecuteReader()
                    If Not dr Is Nothing Then
                        While dr.Read()
                            DOC_FN = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                            VERSION_ID = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                            dName = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                            minio_flg = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                            d_verid = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                            If DMID = "" Then DMID = Trim(CheckDBNull(dr(5), enumObjectType.StrType))
                        End While
                    End If
                Catch ex As Exception
                    GoTo AccessError
                End Try
                dr.Close()
                If Debug = "Y" Then
                    mydebuglog.Debug("  .. DOC_FN: " & DOC_FN)
                    mydebuglog.Debug("  .. VERSION_ID: " & VERSION_ID)
                    mydebuglog.Debug("  .. dName: " & dName)
                    mydebuglog.Debug("  .. minio_flg: " & minio_flg)
                    mydebuglog.Debug("  .. d_verid: " & d_verid)
                    mydebuglog.Debug("  .. DMID: " & DMID)
                End If
            End If
            If VERSION_ID = "" Then GoTo AvailError
            If InStr(DOC_FN, " ") > 0 Then DOC_FN = Replace(DOC_FN, " ", "_")
            If InStr(DOC_FN, "%20") > 0 Then DOC_FN = Replace(DOC_FN, "%20", "_")

            ' ================================================
            ' PREPARE LINK TO THE DOCUMENT
            If InStr(LCase(DOC_FN), ".pdf") > 0 Then newwindow = "True"
            'If isFirefox="true" then newwindow = "True" 
            PublicKey = GenerateKey(DMID)
            ItemName = DOC_FN
            If InStr(ItemName, "getti.ps") > 0 Then
                ReDirect = ItemName
                newwindow = "True"
            Else
                ReDirect = "http://hciscorm.certegrity.com/media/GetDImage.ashx?Domain=" & USER_DOMAIN & "&PublicKey=" & PublicKey & "&ItemName=" & HttpUtility.UrlEncode(ItemName)
            End If
            If Debug = "Y" Then
                mydebuglog.Debug("  .. newwindow: " & newwindow)
                mydebuglog.Debug("  .. PublicKey: " & PublicKey)
                mydebuglog.Debug("  .. ItemName: " & ItemName)
                mydebuglog.Debug("  .. ReDirect: " & ReDirect)
            End If

            ' Log access
            SqlS = "INSERT INTO reports.dbo.CM_LOG(REG_ID, SESSION_ID, DOC_ID, TRANSACTION_ID, ACTION, BROWSER) " & _
                "VALUES('" & UID & "','" & SessID & "','" & DMID & "','" & DMID & "','GET DMS DOCUMENT','" & BROWSER & "')"
            temp = ExecQuery("Insert", "CM_LOG", cmd, SqlS, mydebuglog, Debug)

            If LANG_CD = "ESN" Then
                MSG1 = "<div class=""ui-grid-a"">" & _
                    "<div class=""ui-block-a""><a href=""#"" onClick=""javascript:window.close()"" id=""navbutton1"" data-role=""button"" rel=""external"" data-theme=""b"">Cerrar Ventana</a></div>" & _
                    "<div class=""ui-block-b""><a href=""#"" value=""Print"" onClick=""javascript:window.print()"" id=""navbutton2"" data-role=""button"" rel=""external"" data-theme=""b"">Imprimir Documento</a></div>" & _
                    "</div>"
            Else
                MSG1 = "<div class=""ui-grid-a"">" & _
                    "<div class=""ui-block-a""><a href=""#"" onClick=""javascript:window.close()"" id=""navbutton1"" data-role=""button"" rel=""external"" data-theme=""b"">Close Window</a></div>" & _
                    "<div class=""ui-block-b""><a href=""#"" value=""Print"" onClick=""javascript:window.print()"" id=""navbutton2"" data-role=""button"" rel=""external"" data-theme=""b"">Print Document</a></div>" & _
                    "</div>"
            End If
        Else
            errmsg = errmsg & "Unable to open the database connection. " & vbCrLf
        End If

        ' ================================================
        ' RETURN TO USER
        ' This creates a frame which loads the content.  If the content is an exam, it also loads a Javascript library
        ' that handles the unload event
ReturnControl:
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

AvailError:
        If Debug = "Y" Then mydebuglog.Debug(">>AvailError")
        ErrLvl = "Warning"
        Select Case LANG_CD
            Case "ESN"
                errmsg = "Los documentos no est&aacute;n disponibles en este momento. Por favor, int&eacute;ntelo de nuevo m&aacute;s tarde."
            Case Else
                errmsg = "The document(s) is/are unavailable right now.  Please try again later."
        End Select
        GoTo CloseOut

AccessError:
        If Debug = "Y" Then mydebuglog.Debug(">>AccessError")
        ErrLvl = "Warning"
        Select Case LANG_CD
            Case "ESN"
                errmsg = "Usted no tiene acceso a este documento. Por favor, cierre la sesi&oacute;n y vuelva a intentarlo."
            Case Else
                errmsg = "You do not have access to this document.  Please logout and in again and try again"
        End Select
        GoTo CloseOut

DataError:
        If Debug = "Y" Then mydebuglog.Debug(">>DataError")
        ErrLvl = "Warning"
        Select Case LANG_CD
            Case "ESN"
                errmsg = "No se encontr&oacute; el registro del examen de certificaci&oacute;n para " & TRAIN_TYPE
            Case Else
                errmsg = "The certification exam record for the " & TRAIN_TYPE & " was not found"
        End Select

CloseOut:
        If MSG1 = "" Then MSG1 = "<div class=""ui-grid-a"">" & _
            "<div class=""ui-block-a""><a href=""#"" onClick=""javascript:window.close()"" id=""navbutton1"" data-role=""button"" rel=""external"" data-theme=""b"">Close Window</a></div>" & _
            "</div>"
        If dName = "" Then dName = "Certification Manager for " & DOMAIN

        If Debug = "Y" Then
            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & ">>Final Values")
            mydebuglog.Debug("  .. newwindow: " & newwindow)
            mydebuglog.Debug("  .. dName: " & dName)
            mydebuglog.Debug("  .. FileName: " & FileName)
            mydebuglog.Debug("  .. ReDirect: " & ReDirect)
            mydebuglog.Debug("  .. MSG1: " & MSG1)
            mydebuglog.Debug("  .. DOMAIN: " & DOMAIN)
            mydebuglog.Debug("  .. ErrMsg: " & errmsg)
        End If

        ' ============================================
        ' Finalize output      
        outdata = ""
        outdata = outdata & """newwindow"":""" & newwindow & ""","
        outdata = outdata & """dName"":""" & EscapeJSON(dName) & ""","
        outdata = outdata & """FileName"":""" & EscapeJSON(DOC_FN) & ""","
        outdata = outdata & """ReDirect"":""" & EscapeJSON(ReDirect) & ""","
        outdata = outdata & """MSG1"":""" & EscapeJSON(MSG1) & ""","
        outdata = outdata & """DOMAIN"":""" & EscapeJSON(DOMAIN) & ""","
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
        If Trim(errmsg) <> "" Then myeventlog.Error("WsGetDocument.ashx: " & ErrLvl & ": " & Trim(errmsg))
        myeventlog.Info("WsGetDocument.ashx : DMID : " & DMID & ", UID: " & UID & ", SessID: " & SessID)
        If Debug = "Y" Or (Logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("  outdata: " & outdata & vbCrLf)
                mydebuglog.Debug("Results:  DMID : " & DMID & ", UID: " & UID & ", SessID: " & SessID)
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
                LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, "WSGETDOCUMENT", LogStartTime, VersionNum, Debug)
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
    Public Function GenerateKey(ByVal UID As String) As String
        Dim btUID() As Byte
        Dim encText As New System.Text.UTF8Encoding()
        btUID = encText.GetBytes(UID)
        GenerateKey = ReverseString(ToBase64(btUID))
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