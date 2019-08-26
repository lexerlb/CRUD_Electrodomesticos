Imports System.Data.OleDb
Public Class Form1
    Dim CONEC As New OleDbConnection
    Dim CMDO As New OleDbCommand
    Dim ADAP As New OleDbDataAdapter

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim TABLA As New DataTable
        Dim ARCHIVO As String REM Ubicacion y Nombre del Archivo
        ARCHIVO = "PROVIDER = MICROSOFT.ACE.OLEDB.12.0 ; DATA SOURCE = D:\ELECTROMUNDO.ACCDB"
        CONEC.ConnectionString = ARCHIVO
        CMDO.Connection = CONEC
        CMDO.CommandText = TextBox5.Text
        ADAP.SelectCommand = CMDO
        ADAP.Fill(TABLA)
        DataGridView1.DataSource = TABLA
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
      
    End Sub

    Private Sub ELECTRODOMESTICOSToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Form3.Show()
    End Sub

    Private Sub PROVEEDORESToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Form4.Show()

    End Sub

  
    Private Sub CODPROVEEDORToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub NOMPROVEEDORToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub DIRECCIONToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub TELEFONOToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub NOMBREToolStripMenuItem2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub PRECIOToolStripMenuItem2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub MARCAToolStripMenuItem2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub CANTIDADENALMACENToolStripMenuItem1_Click(sender As Object, e As EventArgs)
       
    End Sub

    Private Sub CODIGOToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub NOMBREToolStripMenuItem3_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub TELEFONOToolStripMenuItem1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        
    End Sub

    Private Sub SALIRToolStripMenuItem_Click(sender As Object, e As EventArgs)
        End
    End Sub

    Private Sub Label9_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label11_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Form3.Show()
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        Form4.Show()
    End Sub

    Private Sub SALIRToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles SALIRToolStripMenuItem.Click
        End
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TextBox5.Text = ""
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Label1.Text = Date.Now.ToLongTimeString()
    End Sub
End Class
