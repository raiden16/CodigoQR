Imports System.Net
Imports System.IO
Imports System.Drawing

Public Class CrearQR

    Private cSBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Private cSBOCompany As SAPbobsCOM.Company


    '//----- METODO DE CREACION DE LA CLASE
    Public Sub New()
        MyBase.New()
        cSBOApplication = oCatchingEvents.SBOApplication
        cSBOCompany = oCatchingEvents.SBOCompany
    End Sub


    Public Function CreateQR(ByVal DocNum As String, ByVal DocDate As String)

        Dim stQueryH As String
        Dim oRecSetH As SAPbobsCOM.Recordset
        Dim URL As String
        Dim QRPath, QRSaved As String
        Dim webclient As WebClient
        Dim stream As Stream
        Dim bitmap As Bitmap

        oRecSetH = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try
            '///URL codificacion de caracteres https://zainex.es/guia-rapida/html/html-url-encode-codificacion
            stQueryH = "Select ""BitmapPath"" from OADP"
            oRecSetH.DoQuery(stQueryH)

            If oRecSetH.RecordCount > 0 Then

                QRPath = oRecSetH.Fields.Item("BitmapPath").Value.ToString & "Codigos QR\"
                DocDate = DocDate.Substring(6, 2) & "/" & DocDate.Substring(4, 2) & "/" & DocDate.Substring(0, 4)

                webclient = New WebClient()
                URL = "https://www.algoryt.com/qr/qr-generator.php?" & DocNum & "%09" & DocDate & "%09%09"

                stream = webclient.OpenRead(URL)
                bitmap = New Bitmap(stream)

                If bitmap IsNot Nothing Then

                    QRSaved = Nothing
                    QRSaved = QRPath & DocNum & ".png"
                    bitmap.Save(QRSaved, System.Drawing.Imaging.ImageFormat.Png)

                End If

            End If

        Catch ex As Exception

            cSBOApplication.MessageBox("Error en CreateQR. " & ex.Message)

        End Try

    End Function

    Public Function ValidarImg(ByVal DocNum As String)

        Dim Ruta, qr As String
        Dim stQueryH As String
        Dim oRecSetH As SAPbobsCOM.Recordset
        qr = Nothing

        oRecSetH = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try

            stQueryH = "Select ""BitmapPath"" from OADP"
            oRecSetH.DoQuery(stQueryH)

            If oRecSetH.RecordCount > 0 Then

                Ruta = oRecSetH.Fields.Item("BitmapPath").Value.ToString & "Codigos QR\"

                Dim dir As New System.IO.DirectoryInfo(Ruta)

                Dim fileList = dir.GetFiles("*.png", System.IO.SearchOption.TopDirectoryOnly)

                Dim FileQuery = From file In fileList
                                Where file.Extension = ".png" And file.Name.Trim.ToString.EndsWith(DocNum & ".png") And file.Name.Trim.ToString.StartsWith(DocNum & ".png")
                                Order By file.CreationTime
                                Select file

                qr = Ruta & DocNum & ".png"

                If FileQuery.Count > 0 Then

                    UpdateQR(DocNum, qr)

                End If

            End If

        Catch ex As Exception

            cSBOApplication.MessageBox("Error en ValidarImg. " & ex.Message)

        End Try

    End Function

    Public Function UpdateQR(ByVal DocNum As String, ByVal qr As String)

        Dim stQueryH, stQueryH2 As String
        Dim oRecSetH, oRecSetH2 As SAPbobsCOM.Recordset
        Dim QRField As String

        oRecSetH = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH2 = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try

            stQueryH = "Select ""U_QRImage"" from OINV where ""DocNum""=" & DocNum
            oRecSetH.DoQuery(stQueryH)

            If oRecSetH.RecordCount = 1 Then

                QRField = oRecSetH.Fields.Item("U_QRImage").Value

                If QRField = Nothing Then

                    stQueryH2 = "Update OINV set ""U_QRImage""='" & qr & "' where ""DocNum""=" & DocNum
                    oRecSetH2.DoQuery(stQueryH2)

                End If

            End If

        Catch ex As Exception

            cSBOApplication.MessageBox("Error en UpdateQR. " & ex.Message)

        End Try

    End Function

End Class
