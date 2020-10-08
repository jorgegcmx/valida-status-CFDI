Imports cfdi = estadodecfdi.cfdistatus
Imports System.Xml
Public Class Form1
    Dim Ruta As String = "C:\Users\AuxDes03\Desktop\versionxml\Egresos\"
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
    Private Sub Cargamosxmlsinpasiovo(ByVal Ruta As String)
        Dim Namefile As String = ""
        ':::Capturador de errores Try
        Try
            ':::Contamos cuanto archivos de texto hay en la carpeta
            Dim Total = My.Computer.FileSystem.GetFiles(Ruta, FileIO.SearchOption.SearchAllSubDirectories, "*.xml")
            ':::Realizamos la búsqueda de la ruta de cada archivo de texto y los agregamos al ListBox
            Dim conter = 0
            ProgressBar1.Value = conter
            For Each archivos As String In My.Computer.FileSystem.GetFiles(Ruta, FileIO.SearchOption.SearchAllSubDirectories, "*.xml")
                Namefile = Replace(archivos, Ruta, "")
                Call leeXmlSinPasivo(archivos, Namefile)
                conter = conter + 1
                ProgressBar1.Value = conter
            Next
            ProgressBar1.Value = 100
        Catch ex As Exception
            MsgBox("No se realizó la operación por: " & ex.Message)
        End Try
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

#Region "EXTRAEMOS LOS DATOS DE CADA XML PARA REVISAR SI YA SE ENCUANTRA UN PASIVO PARA LOS PENDIENTES DE CAPTURA"
    Private Sub leeXmlSinPasivo(ByVal ruta As String, ByVal FileName As String)
        Try
            Dim Version_xml As String = ""
            Dim implocaltraladados As Double = 0
            Dim implocalretenidos As Double = 0
            Dim retenidos As Double = 0
            Dim trasladados As Double = 0
            Dim Subtotal As Double = 0
            Dim valor_noIdentificacion As String = ""
            Dim preciounitarioxml As Double = 0
            Dim importeconcepto As Double = 0
            Dim cantidadvendida As Double = 0
            Dim subtotalclave As Decimal = 0
            Dim descripcionxml As String = ""
            Dim Emisor_Nombre As String = ""
            Dim Emisor_Rfc As String = ""
            Dim Receptor_UsoCFDI As String = ""
            Dim total As Double = 0
            Dim Serie As String = ""
            Dim Folio As String = ""
            Dim SerieFolio As String = ""
            Dim UUID As String = ""
            Dim MetodoPago As String = ""
            Dim TipoDeComprobante As String = ""

            Dim Emisor_RegimenFiscal As String = ""
            Dim Fecha As String = ""
            Dim Sello As String = ""
            Dim NoCertificado As String = ""
            Dim Certificado As String = ""
            Dim Moneda As String = ""
            Dim LugarExpedicion As String = ""
            Dim Receptor_Rfc As String = ""
            Dim Receptor_Nombre As String = ""
            Dim FechaTimbrado As String = ""
            Dim RfcProvCertif As String = ""
            Dim SelloCFD As String = ""
            Dim NoCertificadoSAT As String = ""
            Dim SelloSAT As String = ""
            Dim CtaBeneficiario As String = ""
            Dim CtaOrdenante As String = ""
            Dim FechaPago As String = ""
            Dim FormaDePagoP As String = ""
            Dim MonedaP As String = ""
            Dim Monto As String = 0
            Dim RfcEmisorCtaBen As String = ""
            Dim RfcEmisorCtaOrd As String = ""
            Dim IdDocumento As String = ""
            Dim Version As String = ""
            Dim FormaPago As String = ""


            Dim documentoxml As XmlDocument
            documentoxml = New XmlDocument
            Dim VarManager As XmlNamespaceManager = New XmlNamespaceManager(documentoxml.NameTable)
            documentoxml.Load(Trim(ruta))

            Dim root As XmlElement = documentoxml.DocumentElement
            Dim nodeToFind As XmlNode

            VarManager.AddNamespace("cfdi", "http://www.sat.gob.mx/cfd/3")
            VarManager.AddNamespace("tfd", "http://www.sat.gob.mx/TimbreFiscalDigital")
            VarManager.AddNamespace("implocal", "http://www.sat.gob.mx/implocal")
            VarManager.AddNamespace("pago10", "http://www.sat.gob.mx/Pagos")

            Dim archivo As String = FileName
            Version_xml = documentoxml.SelectSingleNode("/cfdi:Comprobante/@Version", VarManager).InnerText
            TipoDeComprobante = documentoxml.SelectSingleNode("/cfdi:Comprobante/@TipoDeComprobante", VarManager).InnerText
            'DATOS GENERALES COMPROBANTE
            nodeToFind = root.SelectSingleNode("/cfdi:Comprobante/@Serie", VarManager)
            If nodeToFind Is Nothing Then
                Serie = ""
            Else
                Serie = documentoxml.SelectSingleNode("/cfdi:Comprobante/@Serie", VarManager).InnerText
            End If

            Dim nodeToFolio As XmlNode = root.SelectSingleNode("/cfdi:Comprobante/@Folio", VarManager)
            If nodeToFolio Is Nothing Then
                Folio = ""
            Else
                Folio = documentoxml.SelectSingleNode("/cfdi:Comprobante/@Folio", VarManager).InnerText
            End If

            Dim nodeToMetodoPago As XmlNode = root.SelectSingleNode("/cfdi:Comprobante/@MetodoPago", VarManager)
            If nodeToMetodoPago Is Nothing Then
                MetodoPago = ""
            Else
                MetodoPago = documentoxml.SelectSingleNode("/cfdi:Comprobante/@MetodoPago", VarManager).InnerText
            End If

            Dim nodeToFormaPago As XmlNode = root.SelectSingleNode("/cfdi:Comprobante/@FormaPago", VarManager)
            If nodeToFormaPago Is Nothing Then
                FormaPago = ""
            Else
                FormaPago = documentoxml.SelectSingleNode("/cfdi:Comprobante/@FormaPago", VarManager).InnerText
            End If

            SerieFolio = Trim(Serie & Folio)
            Fecha = documentoxml.SelectSingleNode("/cfdi:Comprobante/@Fecha", VarManager).InnerText
            Sello = documentoxml.SelectSingleNode("/cfdi:Comprobante/@Sello", VarManager).InnerText
            NoCertificado = documentoxml.SelectSingleNode("/cfdi:Comprobante/@NoCertificado", VarManager).InnerText
            Certificado = documentoxml.SelectSingleNode("/cfdi:Comprobante/@Certificado", VarManager).InnerText

            Subtotal = documentoxml.SelectSingleNode("/cfdi:Comprobante/@SubTotal", VarManager).InnerText
            Moneda = documentoxml.SelectSingleNode("/cfdi:Comprobante/@Moneda", VarManager).InnerText
            Certificado = documentoxml.SelectSingleNode("/cfdi:Comprobante/@Certificado", VarManager).InnerText
            total = documentoxml.SelectSingleNode("/cfdi:Comprobante/@Total", VarManager).InnerText
            LugarExpedicion = documentoxml.SelectSingleNode("/cfdi:Comprobante/@LugarExpedicion", VarManager).InnerText

            'DATOS GENERALES EMISOR
            Emisor_RegimenFiscal = documentoxml.SelectSingleNode("/cfdi:Comprobante/cfdi:Emisor/@RegimenFiscal", VarManager).InnerText
            Emisor_Rfc = documentoxml.SelectSingleNode("/cfdi:Comprobante/cfdi:Emisor/@Rfc", VarManager).InnerText
            Emisor_Nombre = documentoxml.SelectSingleNode("/cfdi:Comprobante/cfdi:Emisor/@Nombre", VarManager).InnerText

            'DATOS GENERALES RECEPTOR
            Receptor_Rfc = documentoxml.SelectSingleNode("/cfdi:Comprobante/cfdi:Receptor/@Rfc", VarManager).InnerText
            Receptor_UsoCFDI = documentoxml.SelectSingleNode("/cfdi:Comprobante/cfdi:Receptor/@UsoCFDI", VarManager).InnerText

            Dim nodeToNombre As XmlNode = root.SelectSingleNode("/cfdi:Comprobante/cfdi:Receptor/@Nombre", VarManager)

            If nodeToNombre Is Nothing Then
                Receptor_Nombre = ""
            Else
                Receptor_Nombre = documentoxml.SelectSingleNode("/cfdi:Comprobante/cfdi:Receptor/@Nombre", VarManager).InnerText
            End If

            'DATOS GENERALES COMPLEMENTO
            UUID = documentoxml.SelectSingleNode("/cfdi:Comprobante/cfdi:Complemento/tfd:TimbreFiscalDigital/@UUID", VarManager).InnerText
            FechaTimbrado = documentoxml.SelectSingleNode("/cfdi:Comprobante/cfdi:Complemento/tfd:TimbreFiscalDigital/@FechaTimbrado", VarManager).InnerText
            RfcProvCertif = documentoxml.SelectSingleNode("/cfdi:Comprobante/cfdi:Complemento/tfd:TimbreFiscalDigital/@RfcProvCertif", VarManager).InnerText
            SelloCFD = documentoxml.SelectSingleNode("/cfdi:Comprobante/cfdi:Complemento/tfd:TimbreFiscalDigital/@SelloCFD", VarManager).InnerText
            NoCertificadoSAT = documentoxml.SelectSingleNode("/cfdi:Comprobante/cfdi:Complemento/tfd:TimbreFiscalDigital/@NoCertificadoSAT", VarManager).InnerText
            SelloSAT = documentoxml.SelectSingleNode("/cfdi:Comprobante/cfdi:Complemento/tfd:TimbreFiscalDigital/@SelloSAT", VarManager).InnerText

            '----------------------------------------------------------------------------------------------------'
            Dim item As New ListViewItem()

            Dim estatus = ValidaVigenciaCFDI(Emisor_Rfc, Receptor_Rfc, total, UUID)
            If estatus = True Then
                item.Text = "CFDI ( VIGENTE ) Emisor: (" & Emisor_Rfc & ") Receptor: (" & Receptor_Rfc & ") Total: (" & total & ") UUID: " & UUID
                item.ForeColor = Color.Blue
                ListView1.Items.Add(item)
            Else
               item.Text = "CFDI (CANCELADO) Emisor: (" & Emisor_Rfc & ") Receptor: (" & Receptor_Rfc & ") Total: (" & total & ") UUID: " & UUID
                item.ForeColor = Color.Red
                ListView1.Items.Add(item)
            End If
            '----------------------------------------------------------------------------------------------------'

        Catch ex As Exception
            MessageBox.Show("error", ex.Message)
        End Try
    End Sub
#End Region
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Cargamosxmlsinpasiovo(Ruta)
    End Sub
End Class
