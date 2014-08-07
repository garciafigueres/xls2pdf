'POR LUIS GARCÍA FIGUERES para ***** ***** 24/07/2014
Imports Microsoft.Office.Interop.Excel
Imports iTextSharp.text.pdf
Imports iTextSharp.text
Imports System.IO

Public Class frmPrincipal
    Dim archivoExcel As String
    'Arrays para guardar los datos de cada factura
    'Tamaños hardcodeados para abreviar
    Dim numero(0 To 4) As String
    Dim fecha(0 To 4) As String
    Dim nombre(0 To 4) As String
    Dim direccion(0 To 4) As String
    Dim poblacion(0 To 4) As String
    Dim pais(0 To 4) As String
    Dim base(0 To 4) As String
    Dim tipo(0 To 4) As String
    Dim iva(0 To 4) As String
    Dim importe(0 To 4) As String
    Dim factura(0 To 4) As String

    Private Sub btnCargar_Click(sender As Object, e As EventArgs) Handles btnCargar.Click
        'Este método carga el archivo Excel que se indique en el cuadro de diálogo
        'lblEstado.Text = "Botón deshabilitado para agilizar el testeo. Por favor, revisa el código de [btnCargar_Click]"
        ofd.FileName = "llistat_factures.xlsx"
        ofd.Filter = "Listado de facturas | *.xlsx"
        If ofd.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            archivoExcel = ofd.FileName
            lblEstado.Text = "Archivo " & archivoExcel & " cargado correctamente."
        End If
    End Sub

    Private Sub btnLeer_Click(sender As Object, e As EventArgs) Handles btnLeer.Click
        'Este método lee los contenidos del archivo Excel cargado previamente
        If archivoExcel = "" Then
            lblEstado.Text = "Por favor, cargue primero el archivo .xlsx"
        Else
            Dim oExcel As Application = CreateObject("Excel.Application")
            Dim oBook As Workbook = oExcel.Workbooks.Open(archivoExcel, , False)
            'Dim oBook As Workbook = oExcel.Workbooks.Open("C:\Users\Luis\Desktop\YorgaPDF\llistat_factures.xlsx", , False)
            Dim oSheet As Worksheet
            oSheet = oBook.Worksheets("Hoja1")

            'Recorremos cada una de las celdas y guardamos su contenido en el array correspondiente
            'Los límites de fila y columna están hardcodeados por abreviar. Soy consciente de ello.
            For fila As Integer = 2 To 6
                Dim columna As Char
                For col As Integer = 1 To 10
                    Select Case col
                        Case 1
                            columna = "A"
                            numero(fila - 2) = oSheet.Range(columna & fila).Text
                        Case 2
                            columna = "B"
                            fecha(fila - 2) = oSheet.Range(columna & fila).Text
                        Case 3
                            columna = "C"
                            nombre(fila - 2) = oSheet.Range(columna & fila).Text
                        Case 4
                            columna = "D"
                            direccion(fila - 2) = oSheet.Range(columna & fila).Text
                        Case 5
                            columna = "E"
                            poblacion(fila - 2) = oSheet.Range(columna & fila).Text
                        Case 6
                            columna = "F"
                            pais(fila - 2) = oSheet.Range(columna & fila).Text
                        Case 7
                            columna = "G"
                            base(fila - 2) = oSheet.Range(columna & fila).Text
                        Case 8
                            columna = "H"
                            tipo(fila - 2) = oSheet.Range(columna & fila).Text
                        Case 9
                            columna = "I"
                            iva(fila - 2) = oSheet.Range(columna & fila).Text
                        Case 10
                            columna = "J"
                            importe(fila - 2) = oSheet.Range(columna & fila).Text
                        Case Else
                            MsgBox("Error en la columna")
                    End Select
                Next col
                factura(fila - 2) = numero(fila - 2) & " -> " & nombre(fila - 2) & " = " & importe(fila - 2)
                lstFacturas.Items.Add(factura(fila - 2))
            Next fila

            lblEstado.Text = "Datos leídos correctamente."

            'Cerramos el libro
            oBook.Close()
        End If

    End Sub

    Private Sub lstFacturas_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles lstFacturas.MouseDoubleClick
        'Al hacer doble click en alguna de las facturas llamamos al método correspondiente para generar el PDF.
        If lstFacturas.SelectedItem = "" Then
            lblEstado.Text = "No ha elegido nada..."
        Else
            Dim detallesFactura As String
            Dim numFra, fecFra, nomFra, dirFra, pobFra, paiFra, basFra, tipFra, ivaFra, impFra As String

            With lstFacturas
                numFra = numero(.SelectedIndex)
                fecFra = fecha(.SelectedIndex)
                nomFra = nombre(.SelectedIndex)
                dirFra = direccion(.SelectedIndex)
                pobFra = poblacion(.SelectedIndex)
                paiFra = pais(.SelectedIndex)
                basFra = base(.SelectedIndex)
                tipFra = tipo(.SelectedIndex)
                ivaFra = iva(.SelectedIndex)
                impFra = importe(.SelectedIndex)
            End With


            detallesFactura = "Nº factura: " & numFra & _
                "Fecha: " & fecFra & vbCrLf & _
                "Nombre: " & nomFra & vbCrLf & _
                "Dirección: " & dirFra & vbCrLf &
                "Poblacion: " & pobFra & vbCrLf & _
                "Pais: " & paiFra & vbCrLf & _
                "Base: " & basFra & vbCrLf & _
                "Tipo: " & tipFra & vbCrLf & _
                "IVA: " & ivaFra & vbCrLf & _
                "Importe: " & impFra

            Dim resp As DialogResult
            resp = MessageBox.Show("Se va a generar el PDF de la factura siguiente " & vbCrLf & _
                                   detallesFactura,
                                   "GENERAR PDF",
                                   MessageBoxButtons.OKCancel)
            If resp = vbOK Then
                With sfd
                    .FileName = "factura" & numFra & ".pdf"
                    .Filter = "Archivo PDF |*.pdf"
                    If .ShowDialog = DialogResult.OK Then
                        generarPDF(numFra, fecFra, nomFra, dirFra, pobFra, paiFra, basFra, tipFra, ivaFra, impFra, .FileName)
                    End If
                End With
                lblEstado.Text = "PDF generado correctamente."
            Else
                lblEstado.Text = "Generación de PDF cancelada."
            End If
        End If
    End Sub

    Private Sub generarPDF(numero As String, fecha As String, nombre As String, direccion As String, poblacion As String, pais As String, base As String, tipo As String, iva As String, importe As String, ficheroPDF As String)
        'Es en este método donde construyo el archivo PDF. Como ve, está basado en una tabla de 4 columnas,
        'y voy jugando con el formato de las celdas. Los datos los recibo como parámetros.
        Dim logo As Image
        Dim pdfDoc As New Document()
        Dim pdfWrite As PdfWriter = PdfWriter.GetInstance(pdfDoc, New FileStream(ficheroPDF, FileMode.Create))

        'La ruta del logo no debería estar hardcodeada. Nuevamente, por abreviar ;-)
        'Por favor, adáptela a donde tenga usted el archivo
        logo = Image.GetInstance("C:\Users\Luis\Desktop\YorgaPDF\logo.png")

        'Abro lo que será el futuro documento PDF
        pdfDoc.Open()

        'Declaro las variables que almacenan las tablas y las celdas
        Dim tabla As New PdfPTable(4)
        Dim celdaLogo As New PdfPCell(logo)
        Dim filaEnBlanco As New PdfPCell(New Phrase(" "))
        Dim datosCliente As New PdfPCell(New Paragraph("DATOS CLIENTE:" & vbCrLf & nombre & vbCrLf & direccion & vbCrLf & poblacion & "(" & pais & ")"))
        Dim datosFactura As New PdfPCell(New Paragraph("FACTURA Nº: " & numero & " DE FECHA: " & fecha))
        Dim pieBase As New PdfPCell(New Paragraph("Base imponible: "))
        Dim pieBaseNum As New PdfPCell(New Paragraph(base))
        Dim pieIva As New PdfPCell(New Paragraph("IVA (" & tipo & "%): "))
        Dim pieIvaNum As New PdfPCell(New Paragraph(iva))
        Dim pieImporte As New PdfPCell(New Paragraph("Importe: "))
        Dim pieImporteNum As New PdfPCell(New Paragraph(importe))

        'Doy formato a las celdas
        celdaLogo.Colspan = 2
        celdaLogo.Padding = 20
        celdaLogo.Border = PdfPCell.NO_BORDER
        filaEnBlanco.Colspan = 4
        filaEnBlanco.Border = PdfPCell.NO_BORDER
        datosCliente.Colspan = 2
        datosCliente.HorizontalAlignment = 0
        datosCliente.PaddingTop = 30
        datosFactura.Colspan = 4
        datosFactura.Padding = 8
        datosFactura.Border = PdfPCell.NO_BORDER
        pieBase.Colspan = 3
        pieBase.Padding = 8
        pieBase.HorizontalAlignment = 2
        pieBase.Border = PdfPCell.NO_BORDER
        pieIva.Colspan = 3
        pieIva.Padding = 8
        pieIva.HorizontalAlignment = 2
        pieIva.Border = PdfPCell.NO_BORDER
        pieImporte.Colspan = 3
        pieImporte.Padding = 8
        pieImporte.HorizontalAlignment = 2
        pieImporte.Border = PdfPCell.NO_BORDER
        pieBaseNum.HorizontalAlignment = 2
        pieIvaNum.HorizontalAlignment = 2
        pieImporteNum.HorizontalAlignment = 2

        'Añado las celdas a la tabla
        tabla.AddCell(celdaLogo)
        tabla.AddCell(datosCliente)
        tabla.AddCell(filaEnBlanco)
        tabla.AddCell(datosFactura)
        tabla.AddCell(filaEnBlanco)
        tabla.AddCell("Cantidad")
        tabla.AddCell("Descripción")
        tabla.AddCell("PVP")
        tabla.AddCell("Importe")
        tabla.AddCell("1")
        tabla.AddCell("Artículo genérico")
        tabla.AddCell(base)
        tabla.AddCell(base)

        'Añado unas cuantas celdas vacías para rellenar el cuerpo y que quede bien estéticamente :D
        tabla.AddCell(" ")
        tabla.AddCell(" ")
        tabla.AddCell(" ")
        tabla.AddCell(" ")
        tabla.AddCell(" ")
        tabla.AddCell(" ")
        tabla.AddCell(" ")
        tabla.AddCell(" ")
        tabla.AddCell(" ")
        tabla.AddCell(" ")
        tabla.AddCell(" ")
        tabla.AddCell(" ")
        tabla.AddCell(" ")
        tabla.AddCell(" ")
        tabla.AddCell(" ")
        tabla.AddCell(" ")
        tabla.AddCell(" ")
        tabla.AddCell(" ")
        tabla.AddCell(" ")
        tabla.AddCell(" ")
        tabla.AddCell(" ")
        tabla.AddCell(" ")
        tabla.AddCell(" ")
        tabla.AddCell(" ")

        'Añado una doble separación antes del pie de la factura
        tabla.AddCell(filaEnBlanco)
        tabla.AddCell(filaEnBlanco)

        'Añado el pie de factura con los datos numéricos de la misma
        tabla.AddCell(pieBase)
        tabla.AddCell(pieBaseNum)
        tabla.AddCell(pieIva)
        tabla.AddCell(pieIvaNum)
        tabla.AddCell(pieImporte)
        tabla.AddCell(pieImporteNum)

        'Añado la propia tabla al documento y lo cierro
        pdfDoc.Add(tabla)
        pdfDoc.Close()
    End Sub

    Private Sub btnSalir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSalir.Click
        'Antes de salir pido confirmación del usuario
        Dim resp As DialogResult
        resp = MessageBox.Show("¿Seguro que desea salir?", "SALIR DE LA APLICACIÓN", MessageBoxButtons.YesNo)
        If resp = vbYes Then
            Close()
        End If
    End Sub

    
End Class
