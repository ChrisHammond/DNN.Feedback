Imports System.ComponentModel
Imports System.Globalization
Imports System.Net
Imports System.Web.Script.Serialization
Imports System.Web.UI
Imports System.Web.UI.WebControls

Namespace DotNetNuke.Modules.Feedback

    Public Class GoogleVerificationResult
        Private _success As Boolean

        Public Property Success As Boolean
            Get
                Return _success
            End Get
            Set(value As Boolean)
                _success = value
            End Set
        End Property
    End Class

    <ToolboxData("<{0}:ReCaptcha runat=server><{0}:ReCaptcha>")>
    Public Class ReCaptcha
        Inherits WebControl

        Dim ValidateResponseField As String = "g-recaptcha-response"
        Dim LanguageCode As String = CultureInfo.CurrentUICulture.TwoLetterISOLanguageName

        <Bindable(True)>
        <Category("Appearance")>
        <DefaultValue("")>
        <Localizable(False)>
        Public Property SiteKey As String
            Get
                Dim _siteKey As String = DirectCast(ViewState("_ReCaptchaSiteKey"), [String])
                Return _siteKey
            End Get
            Set(value As String)
                ViewState("_ReCaptchaSiteKey") = value
            End Set
        End Property

        <Bindable(True)>
        <Category("Appearance")>
        <DefaultValue("")>
        <Localizable(False)>
        Public Property SecretKey As String
            Get
                Dim _secretKey As String = DirectCast(ViewState("_ReCaptchaSecretKey"), [String])
                Return _secretKey
            End Get
            Set(value As String)
                ViewState("_ReCaptchaSecretKey") = value
            End Set
        End Property

        Protected Overrides Sub RenderContents(writer As HtmlTextWriter)
            If Not String.IsNullOrEmpty(SiteKey) AndAlso Not String.IsNullOrEmpty(SecretKey) Then
                Dim ReCaptchaHTML As String = String.Format("<div class='g-recaptcha' data-sitekey='{0}'></div>", SiteKey)

                Dim ReCaptchaScript As String = String.Format(" <script src=""https://www.google.com/recaptcha/api.js?hl={0}"" async defer></script>", LanguageCode)
                writer.Write(ReCaptchaHTML)
                Me.Page.ClientScript.RegisterStartupScript(Me.GetType, "ReCaptchaScript", ReCaptchaScript, False)
            End If
        End Sub

        Public Function Validate() As Boolean
            Dim EncodedResponse As String = Me.Page.Request.Form(ValidateResponseField)
            If String.IsNullOrEmpty(EncodedResponse) Or String.IsNullOrEmpty(SecretKey) Then
                Return False
            Else
                ' Bypass certificate validation
                System.Net.ServicePointManager.ServerCertificateValidationCallback = Function()
                                                                                         Return True
                                                                                     End Function
                System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
                Dim client As New System.Net.WebClient()
                Dim GoogleReply As String = client.DownloadString(String.Format("https://www.google.com/recaptcha/api/siteverify?secret={0}&response={1}", SecretKey, EncodedResponse))

                Dim serializer As New JavaScriptSerializer()
                Dim gOutput As GoogleVerificationResult = serializer.Deserialize(Of GoogleVerificationResult)(GoogleReply)

                Return gOutput.Success
            End If
        End Function

    End Class

End Namespace
