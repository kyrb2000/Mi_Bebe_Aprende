Attribute VB_Name = "Bebe_Aprende"
Private Type RegMillonario
    Preguntas As String
    Respuestas As String
End Type
Public MILLONARIO() As RegMillonario
Public ALEATORIO() As RegMillonario
Public NumRegistros As Integer
Public Nivel As Byte 'Hay 3 niveles
Public Veces As Integer
Public Segundos As Byte
Public Segundos3 As Integer
Public Tiempo_Activo As Boolean

Sub Main()
    'frmLector.Show
    'frmCalendario.Show
    frmLetrar.Show
End Sub
