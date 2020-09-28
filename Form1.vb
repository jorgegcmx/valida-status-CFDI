Imports cfdi = estadodecfdi.cfdistatus
Public Class Form1
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim estatus = ValidaVigenciaCFDI("GPT181004NV8", "AQC880627A23", "24.97", "3E326422-4AEF-4264-801D-8843AC073015")
        If estatus = True Then
            MsgBox("Vigente, Pasa a Flujo")
        Else
            MsgBox("cancelado, No pasa a Fujo")
        End If
    End Sub


    Function ValidaVigenciaCFDI(ByVal RFCEMISOR As String, ByVal RFCRECEPTOR As String, ByVal IMPORTE As String, ByVal UUID As String) As Boolean
      Try
            Dim st As New cfdistatus.ConsultaCFDIServiceClient
            Dim querySend As String
            Dim Acuse
            Dim statusCFDI
            querySend = "?re=" + RFCEMISOR + "&rr=" + RFCRECEPTOR + "&tt=" + IMPORTE + "&id=" + UUID + ""
            Acuse = st.Consulta(querySend)
            statusCFDI = Acuse.Estado.ToLower()
            If statusCFDI = "vigente" Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MessageBox.Show("error", ex.Message)
            MsgBox("El Servicio del SAT No esta Activo")
        End Try
    End Function
End Class
