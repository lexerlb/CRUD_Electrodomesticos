Imports System.Data.OleDb
Public Class Form3
    Dim CONEC As New OleDbConnection
    Dim CMDO As New OleDbCommand
    Dim ADAP As New OleDbDataAdapter
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TABLA As New DataTable
        Dim ARCHIVO As String REM Ubicacion y Nombre del Archivo
        ARCHIVO = "PROVIDER = MICROSOFT.ACE.OLEDB.12.0 ; DATA SOURCE = D:\ELECTROMUNDO.ACCDB"
        CONEC.ConnectionString = ARCHIVO
        CMDO.Connection = CONEC
        CMDO.CommandText = "select * from ELECTRODOMESTICOS"
        ADAP.SelectCommand = CMDO
        ADAP.Fill(TABLA)
        DataGridView1.DataSource = TABLA
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim TABLA As New DataTable
        Dim ARCHIVO As String REM Ubicacion y Nombre del Archivo
        ARCHIVO = "PROVIDER = MICROSOFT.ACE.OLEDB.12.0 ; DATA SOURCE = D:\ELECTROMUNDO.ACCDB"
        CONEC.ConnectionString = ARCHIVO
        CMDO.Connection = CONEC
        CMDO.CommandText = "insert INTO ELECTRODOMESTICOS(NOMBRE, PRECIO, MARCA, CANT_ALMACEN,COD_PROVEEDOR) VALUES ('" + TextBox1.Text + "'," + TextBox2.Text + ",'" + TextBox3.Text + "'," + TextBox4.Text + "," + TextBox5.Text + ")"
        ADAP.SelectCommand = CMDO
        ADAP.Fill(TABLA)
        DataGridView1.DataSource = TABLA

        
        CMDO.CommandText = "select * from ELECTRODOMESTICOS"
        ADAP.SelectCommand = CMDO
        ADAP.Fill(TABLA)
        DataGridView1.DataSource = TABLA
    End Sub

    Private Sub IDToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles IDToolStripMenuItem.Click
        Dim TABLA As New DataTable
        Dim ARCHIVO As String REM Ubicacion y Nombre del Archivo
        ARCHIVO = "PROVIDER = MICROSOFT.ACE.OLEDB.12.0 ; DATA SOURCE = D:\ELECTROMUNDO.ACCDB"
        CONEC.ConnectionString = ARCHIVO
        CMDO.Connection = CONEC
        CMDO.CommandText = "select * from ELECTRODOMESTICOS WHERE ID_PRODUCTO = " & TextBox6.Text + ""
        ADAP.SelectCommand = CMDO
        ADAP.Fill(TABLA)
        DataGridView1.DataSource = TABLA
    End Sub

    Private Sub ATRASToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ATRASToolStripMenuItem.Click
        Form1.Show()
        Me.Hide()
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub NOMBREToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles NOMBREToolStripMenuItem.Click
        Dim TABLA As New DataTable
        Dim ARCHIVO As String REM Ubicacion y Nombre del Archivo
        ARCHIVO = "PROVIDER = MICROSOFT.ACE.OLEDB.12.0 ; DATA SOURCE = D:\ELECTROMUNDO.ACCDB"
        CONEC.ConnectionString = ARCHIVO
        CMDO.Connection = CONEC
        CMDO.CommandText = "select * from ELECTRODOMESTICOS WHERE NOMBRE = '" & TextBox6.Text + "'"
        ADAP.SelectCommand = CMDO
        ADAP.Fill(TABLA)
        DataGridView1.DataSource = TABLA
    End Sub

    Private Sub PRECIOToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles PRECIOToolStripMenuItem.Click
        Dim TABLA As New DataTable
        Dim ARCHIVO As String REM Ubicacion y Nombre del Archivo
        ARCHIVO = "PROVIDER = MICROSOFT.ACE.OLEDB.12.0 ; DATA SOURCE = D:\ELECTROMUNDO.ACCDB"
        CONEC.ConnectionString = ARCHIVO
        CMDO.Connection = CONEC
        CMDO.CommandText = "select * from ELECTRODOMESTICOS WHERE PRECIO = " & TextBox6.Text + ""
        ADAP.SelectCommand = CMDO
        ADAP.Fill(TABLA)
        DataGridView1.DataSource = TABLA
    End Sub

    Private Sub MARCAToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles MARCAToolStripMenuItem.Click
        Dim TABLA As New DataTable
        Dim ARCHIVO As String REM Ubicacion y Nombre del Archivo
        ARCHIVO = "PROVIDER = MICROSOFT.ACE.OLEDB.12.0 ; DATA SOURCE = D:\ELECTROMUNDO.ACCDB"
        CONEC.ConnectionString = ARCHIVO
        CMDO.Connection = CONEC
        CMDO.CommandText = "select * from ELECTRODOMESTICOS WHERE MARCA = '" & TextBox6.Text + "'"
        ADAP.SelectCommand = CMDO
        ADAP.Fill(TABLA)
        DataGridView1.DataSource = TABLA
    End Sub

    Private Sub PROVEEDORToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PROVEEDORToolStripMenuItem.Click
        Dim TABLA As New DataTable
        Dim ARCHIVO As String REM Ubicacion y Nombre del Archivo
        ARCHIVO = "PROVIDER = MICROSOFT.ACE.OLEDB.12.0 ; DATA SOURCE = D:\ELECTROMUNDO.ACCDB"
        CONEC.ConnectionString = ARCHIVO
        CMDO.Connection = CONEC
        CMDO.CommandText = "select * from ELECTRODOMESTICOS WHERE COD_PROVEEDOR = '" & TextBox6.Text + "'"
        ADAP.SelectCommand = CMDO
        ADAP.Fill(TABLA)
        DataGridView1.DataSource = TABLA
    End Sub

    Private Sub NOMBREToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles NOMBREToolStripMenuItem1.Click
        Dim TABLA As New DataTable
        Dim ARCHIVO As String REM Ubicacion y Nombre del Archivo
        ARCHIVO = "PROVIDER = MICROSOFT.ACE.OLEDB.12.0 ; DATA SOURCE = D:\ELECTROMUNDO.ACCDB"
        CONEC.ConnectionString = ARCHIVO
        CMDO.Connection = CONEC
        CMDO.CommandText = "UPDATE ELECTRODOMESTICOS set NOMBRE = '" + TextBox7.Text + "' WHERE ID = " + TextBox8.Text
        ADAP.SelectCommand = CMDO
        ADAP.Fill(TABLA)
        DataGridView1.DataSource = TABLA
    End Sub

    Private Sub CANTIDADENALMACENToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CANTIDADENALMACENToolStripMenuItem.Click
        Dim TABLA As New DataTable
        Dim ARCHIVO As String REM Ubicacion y Nombre del Archivo
        ARCHIVO = "PROVIDER = MICROSOFT.ACE.OLEDB.12.0 ; DATA SOURCE = D:\ELECTROMUNDO.ACCDB"
        CONEC.ConnectionString = ARCHIVO
        CMDO.Connection = CONEC
        CMDO.CommandText = "UPDATE ELECTRODOMESTICOS set PRECIO = '" + TextBox7.Text + "' WHERE ID = " + TextBox8.Text
        ADAP.SelectCommand = CMDO
        ADAP.Fill(TABLA)
        DataGridView1.DataSource = TABLA
    End Sub

    Private Sub MARCAToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles MARCAToolStripMenuItem1.Click
        Dim TABLA As New DataTable
        Dim ARCHIVO As String REM Ubicacion y Nombre del Archivo
        ARCHIVO = "PROVIDER = MICROSOFT.ACE.OLEDB.12.0 ; DATA SOURCE = D:\ELECTROMUNDO.ACCDB"
        CONEC.ConnectionString = ARCHIVO
        CMDO.Connection = CONEC
        CMDO.CommandText = "UPDATE ELECTRODOMESTICOS set MARCA = '" + TextBox7.Text + "' WHERE ID = " + TextBox8.Text
        ADAP.SelectCommand = CMDO
        ADAP.Fill(TABLA)
        DataGridView1.DataSource = TABLA
    End Sub

    Private Sub CANTIDADToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CANTIDADToolStripMenuItem.Click
        Dim TABLA As New DataTable
        Dim ARCHIVO As String REM Ubicacion y Nombre del Archivo
        ARCHIVO = "PROVIDER = MICROSOFT.ACE.OLEDB.12.0 ; DATA SOURCE = D:\ELECTROMUNDO.ACCDB"
        CONEC.ConnectionString = ARCHIVO
        CMDO.Connection = CONEC
        CMDO.CommandText = "UPDATE ELECTRODOMESTICOS set MARCA = '" + TextBox7.Text + "' WHERE ID = " + TextBox8.Text
        ADAP.SelectCommand = CMDO
        ADAP.Fill(TABLA)
        DataGridView1.DataSource = TABLA
    End Sub

    Private Sub PROVEEDORToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles PROVEEDORToolStripMenuItem1.Click
        Dim TABLA As New DataTable
        Dim ARCHIVO As String REM Ubicacion y Nombre del Archivo
        ARCHIVO = "PROVIDER = MICROSOFT.ACE.OLEDB.12.0 ; DATA SOURCE = D:\ELECTROMUNDO.ACCDB"
        CONEC.ConnectionString = ARCHIVO
        CMDO.Connection = CONEC
        CMDO.CommandText = "UPDATE ELECTRODOMESTICOS set COD_PROVEEDOR = '" + TextBox7.Text + "' WHERE ID = " + TextBox8.Text
        ADAP.SelectCommand = CMDO
        ADAP.Fill(TABLA)
        DataGridView1.DataSource = TABLA
    End Sub

    Private Sub VERDATOSToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VERDATOSToolStripMenuItem.Click
        Dim TABLA As New DataTable
        Dim ARCHIVO As String REM Ubicacion y Nombre del Archivo
        ARCHIVO = "PROVIDER = MICROSOFT.ACE.OLEDB.12.0 ; DATA SOURCE = D:\ELECTROMUNDO.ACCDB"
        CONEC.ConnectionString = ARCHIVO
        CMDO.Connection = CONEC
        CMDO.CommandText = "select * from ELECTRODOMESTICOS"
        ADAP.SelectCommand = CMDO
        ADAP.Fill(TABLA)
        DataGridView1.DataSource = TABLA
    End Sub

   
    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Dim TABLA As New DataTable
        Dim ARCHIVO As String REM Ubicacion y Nombre del Archivo
        ARCHIVO = "PROVIDER = MICROSOFT.ACE.OLEDB.12.0 ; DATA SOURCE = D:\ELECTROMUNDO.ACCDB"
        CONEC.ConnectionString = ARCHIVO
        CMDO.Connection = CONEC
        CMDO.CommandText = "delete from ELECTRODOMESTICOS WHERE ID_PRODUCTO = " & TextBox9.Text
        ADAP.SelectCommand = CMDO
        ADAP.Fill(TABLA)
        DataGridView1.DataSource = TABLA
    End Sub
End Class