Option Strict On
Option Explicit On 
Imports Microsoft.Win32
Public Class ClassRegistry
    Public Sub DelSetting(ByVal subKey As String, ByVal name As String)
        Dim regSubKey As RegistryKey = Registry.LocalMachine.CreateSubKey("Software\Faena\" + subKey)
        regSubKey.DeleteValue(name)
    End Sub

    Public Sub SaveSetting(ByVal subKey As String, ByVal name As String, ByVal value As String)
        Dim regSubKey As RegistryKey = Registry.LocalMachine.CreateSubKey("Software\Faena\" + subKey)
        regSubKey.SetValue(name, value)
    End Sub

    Public Sub SaveSetting(ByVal subKey As String, ByVal name As String, ByVal value As Long)
        Dim regSubKey As RegistryKey = Registry.LocalMachine.CreateSubKey("Software\Faena\" + subKey)
        regSubKey.SetValue(name, CInt(value))
    End Sub

    Public Sub SaveSetting(ByVal subKey As String, ByVal name As String, ByVal value As Double)
        Dim regSubKey As RegistryKey = Registry.LocalMachine.CreateSubKey("Software\Faena\" + subKey)
        regSubKey.SetValue(name, CInt(value))
    End Sub
    Public Sub SaveSetting(ByVal subKey As String, ByVal name As String, ByVal value As Boolean)
        Dim regSubKey As RegistryKey = Registry.LocalMachine.CreateSubKey("Software\Faena\" + subKey)
        If value Then
            regSubKey.SetValue(name, 1)
        Else
            regSubKey.SetValue(name, 0)
        End If
    End Sub

    Public Sub SaveSetting(ByVal subKey As String, ByVal name As String, ByVal value As Decimal)
        Dim regSubKey As RegistryKey = Registry.LocalMachine.CreateSubKey("Software\Faena\" + subKey)
        regSubKey.SetValue(name, CInt(value))
    End Sub
    Public Sub SaveSetting(ByVal subKey As String, ByVal name As String, ByVal value As Short)
        Dim regSubKey As RegistryKey = Registry.LocalMachine.CreateSubKey("Software\Faena\" + subKey)
        regSubKey.SetValue(name, CInt(value))
    End Sub

    Public Function GetSetting(ByVal subKey As String, ByVal name As String, Optional ByVal defaultValue As String = "") As String
        Dim regSubKey As RegistryKey = Registry.LocalMachine.OpenSubKey("Software\Faena\" + subKey, False)
        If regSubKey Is Nothing Then
            Return defaultValue
        Else
            Return regSubKey.GetValue(name, defaultValue).ToString
        End If
    End Function

    Public Function GetSetting(ByVal subKey As String, ByVal name As String, ByVal defaultValue As Integer) As Integer
        Dim regSubKey As RegistryKey = Registry.LocalMachine.OpenSubKey("Software\Faena\" + subKey, False)
        If regSubKey Is Nothing Then
            Return defaultValue
        Else
            Return CInt(regSubKey.GetValue(name, defaultValue))
        End If
    End Function

    Public Function GetSetting(ByVal subKey As String, ByVal name As String, ByVal defaultValue As Boolean) As Boolean
        Dim regSubKey As RegistryKey = Registry.LocalMachine.OpenSubKey("Software\Faena\" + subKey, False)
        If regSubKey Is Nothing Then
            Return defaultValue
        Else
            Return CBool(regSubKey.GetValue(name, defaultValue))
        End If
    End Function

    Public Function GetSetting(ByVal subKey As String, ByVal name As String, ByVal defaultValue As Short) As Short
        Dim regSubKey As RegistryKey = Registry.LocalMachine.OpenSubKey("Software\Faena\" + subKey, False)
        If regSubKey Is Nothing Then
            Return defaultValue
        Else
            Return CType(regSubKey.GetValue(name, defaultValue), Short)
        End If
    End Function
    Public Function GetSetting(ByVal subKey As String, ByVal name As String, ByVal defaultValue As Long) As Long
        Dim regSubKey As RegistryKey = Registry.LocalMachine.OpenSubKey("Software\Faena\" + subKey, False)
        If regSubKey Is Nothing Then
            Return defaultValue
        Else
            Return CType(regSubKey.GetValue(name, defaultValue), Long)
        End If
    End Function
    Public Function GetSetting(ByVal subKey As String, ByVal name As String, ByVal defaultValue As Double) As Double
        Dim regSubKey As RegistryKey = Registry.LocalMachine.OpenSubKey("Software\Faena\" + subKey, False)
        If regSubKey Is Nothing Then
            Return defaultValue
        Else
            Return CType(regSubKey.GetValue(name, defaultValue), Double)
        End If
    End Function
    Public Function GetSetting(ByVal subKey As String, ByVal name As String, ByVal defaultValue As Decimal) As Decimal
        Dim regSubKey As RegistryKey = Registry.LocalMachine.OpenSubKey("Software\Faena\" + subKey, False)
        If regSubKey Is Nothing Then
            Return defaultValue
        Else
            Return CType(regSubKey.GetValue(name, defaultValue), Decimal)
        End If
    End Function
End Class
