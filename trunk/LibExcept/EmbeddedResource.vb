Imports System.Drawing
Imports System.Reflection
Imports System.IO

'--
'-- A simple class to load Embedded Resources from the current Assembly
'--
'-- currently only supports loading images, but could be extended to load any binary filetype
'--
'--
'-- Jeff Atwood
'-- http://www.codinghorror.com
'--
Friend Class EmbeddedResource

    Private Shared _objMyAssembly As System.Reflection.Assembly = Nothing
    Private Shared _strAssemblyName As String = ""

    Private Sub New()
        '-- to keep this class from being creatable as an instance.
    End Sub

    '--
    '-- get the currently executing assembly
    '--
    Private Shared Sub GetAssembly()
        If _objMyAssembly Is Nothing Then
            _objMyAssembly = System.Reflection.Assembly.GetExecutingAssembly
        End If
        If _strAssemblyName = "" Then
            _strAssemblyName = _objMyAssembly.GetName.Name + "."
        End If
    End Sub


    '--
    '-- return a placeholder image 
    '-- contains text of image we failed to load
    '--
    Private Shared Function GetFailureImage(ByVal strResourceName As String) As Image
        Dim objBitmap As New Bitmap(128, 32)
        Dim objGraphics As Graphics = Graphics.FromImage(objBitmap)
        Dim strText As String = _strAssemblyName + Environment.NewLine + strResourceName
        Dim drawBrush As New SolidBrush(Color.Red)
        Dim objFont As New Font(System.Drawing.FontFamily.GenericSansSerif, 7)

        objGraphics.DrawString(strText, objFont, drawBrush, 1, 1)
        Return objBitmap
    End Function

    '--
    '-- notify user of failure to load a resource
    '--
    Private Shared Sub NotifyFailure(ByVal strResourceType As String, ByVal strResourceName As String)
        Debug.WriteLine("** failed to load embedded " + strResourceType + " resource named: ")
        Debug.WriteLine("  " + _strAssemblyName + strResourceName)
    End Sub

    '--
    '-- grab an image stored as an embedded resource in this project
    '--
    Public Shared Function GetImage(ByVal strResourceName As String) As Image
        GetAssembly()
        Dim objImage As Image
        Try
            objImage = Image.FromStream(_objMyAssembly.GetManifestResourceStream(_strAssemblyName + strResourceName))
        Catch e As Exception
            '-- build placeholder error image
            objImage = GetFailureImage(strResourceName)
            NotifyFailure("Image", strResourceName)
        End Try
        Return objImage
    End Function

End Class