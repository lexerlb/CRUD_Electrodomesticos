Imports System.Data.OleDb
Public Class Form4
    Dim CONEC As New OleDbConnection
    Dim CMDO As New OleDbCommand
    Dim ADAP As New OleDbDataAdapter
    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TABLA As New DataTable
        Dim ARCHIVO As String REM Ubicacion y Nombre del Archivo
        ARCHIVO = "PROVIDER = MICROSOFT.ACE.OLEDB.12.0 ; DATA SOURCE = D:\ELECTROMUNDO.ACCDB"
        CONEC.ConnectionString = ARCHIVO
        CMDO.Connection = CONEC
        CMDO.CommandText = "select * from PROVEEDORES"
        ADAP.SelectCommand = CMDO
        ADAP.Fill(TABLA)
        DataGridView1.DataSource = TABLA
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim TABLA As New DataTable
        Dim ARCHIVO As String REM Ubicacion y Nombre del Archivo
        ARCHIVO = "PROVIDER = MICROSOFT.ACE.OLEDB.12.0 ; DATA SOURCE = D:\ELECTROMUNDO.ACCDB"
        CONEC.ConnectionString = ARCHIVO
        CMDO.Connection = CONEC
        CMDO.CommandText = "insert INTO PROVEEDORES(NOM_PROVEEDOR, DIRECCION, TELEFONO) VALUES ('" + TextBox12.Text + "','" + TextBox11.Text + "'," + TextBox10.Text + ")"
        ADAP.SelectCommand = CMDO
        ADAP.Fill(TABLA)
        DataGridView1.DataSource = TABLA

        CMDO.CommandText = "select * from PROVEEDORES"
        ADAP.SelectCommand = CMDO
        ADAP.Fill(TABLA)
        DataGridView1.DataSource = TABLA
    End Sub

    Private Sub ATRASToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ATRASToolStripMenuItem.Click
        Form1.Show()
        Me.Hide()
    End Sub

    Private Sub CODPROVEEDORToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CODPROVEEDORToolStripMenuItem.Click
        Dim TABLA As New DataTable
        Dim ARCHIVO As String REM Ubicacion y Nombre del Archivo
        ARCHIVO = "PROVIDER = MICROSOFT.ACE.OLEDB.12.0 ; DATA SOURCE = D:\ELECTROMUNDO.ACCDB"
        CONEC.ConnectionString = ARCHIVO
        CMDO.Connection = CONEC
        CMDO.CommandText = "select * from PROVEEDORES WHERE COD_PROVEEDOR = " & TextBox6.Text + ""
        ADAP.SelectCommand = CMDO
        ADAP.Fill(TABLA)
        DataGridView1.DataSource = TABLA
    End Sub

    Private Sub NOMPROVEEDORToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NOMPROVEEDORToolStripMenuItem.Click
        Dim TABLA As New DataTable
        Dim ARCHIVO As String REM Ubicacion y Nombre del Archivo
        ARCHIVO = "PROVIDER = MICROSOFT.ACE.OLEDB.12.0 ; DATA SOURCE = D:\ELECTROMUNDO.ACCDB"
        CONEC.ConnectionString = ARCHIVO
        CMDO.Connection = CONEC
        CMDO.CommandText = "select * from PROVEEDORES WHERE NOM_PROVEEDOR = '" & TextBox6.Text + "'"
        ADAP.SelectCommand = CMDO
        ADAP.Fill(TABLA)
        DataGridView1.DataSource = TABLA
    End Sub

    Private Sub DIRECCIONToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DIRECCIONToolStripMenuItem.Click
        Dim TABLA As New DataTable
        Dim ARCHIVO As String REM Ubicacion y Nombre del Archivo
        ARCHIVO = "PROVIDER = MICROSOFT.ACE.OLEDB.12.0 ; DATA SOURCE = D:\ELECTROMUNDO.ACCDB"
        CONEC.ConnectionString = ARCHIVO
        CMDO.Connection = CONEC
        CMDO.CommandText = "select * from PROVEEDORES WHERE DIRECCION = '" & TextBox6.Text + "'"
        ADAP.SelectCommand = CMDO
        ADAP.Fill(TABLA)
        DataGridView1.DataSource = TABLA
    End Sub

    Private Sub TELEFONOToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TELEFONOToolStripMenuItem.Click
        Dim TABLA As New DataTable
        Dim ARCHIVO As String REM Ubicacion y Nombre del Archivo
        ARCHIVO = "PROVIDER = MICROSOFT.ACE.OLEDB.12.0 ; DATA SOURCE = D:\ELECTROMUNDO.ACCDB"
        CONEC.ConnectionString = ARCHIVO
        CMDO.Connection = CONEC
        CMDO.CommandText = "select * from PROVEEDORES WHERE TELEFONO = " & TextBox6.Text + ""
        ADAP.SelectCommand = CMDO
        ADAP.Fill(TABLA)
        DataGridView1.DataSource = TABLA
    End Sub

    Private Sub CODIGOToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CODIGOToolStripMenuItem.Click
        Dim TABLA As New DataTable
        Dim ARCHIVO As String REM Ubicacion y Nombre del Archivo
        ARCHIVO = "PROVIDER = MICROSOFT.ACE.OLEDB.12.0 ; DATA SOURCE = D:\ELECTROMUNDO.ACCDB"
        CONEC.ConnectionString = ARCHIVO
        CMDO.Connection = CONEC
        CMDO.CommandText = "UPDATE PROVEEDORES set NOM_PROVEEDOR = '" + TextBox7.Text + "' WHERE ID = " + TextBox8.Text
        ADAP.SelectCommand = CMDO
        ADAP.Fill(TABLA)
        DataGridView1.DataSource = TABLA
    End Sub

    Private Sub NOMBREToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles NOMBREToolStripMenuItem3.Click
        Dim TABLA As New DataTable
        Dim ARCHIVO As String REM Ubicacion y Nombre del Archivo
        ARCHIVO = "PROVIDER = MICROSOFT.ACE.OLEDB.12.0 ; DATA SOURCE = D:\ELECTROMUNDO.ACCDB"
        CONEC.ConnectionString = ARCHIVO
        CMDO.Connection = CONEC
        CMDO.CommandText = "UPDATE PROVEEDORES set DIRECCION = '" + TextBox7.Text + "' WHERE ID = " + TextBox8.Text
        ADAP.SelectCommand = CMDO
        ADAP.Fill(TABLA)
        DataGridView1.DataSource = TABLA
    End Sub

    Private Sub TELEFONOToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles TELEFONOToolStripMenuItem1.Click
        Dim TABLA As New DataTable
        Dim ARCHIVO As String REM Ubicacion y Nombre del Archivo
        ARCHIVO = "PROVIDER = MICROSOFT.ACE.OLEDB.12.0 ; DATA SOURCE = D:\ELECTROMUNDO.ACCDB"
        CONEC.ConnectionString = ARCHIVO
        CMDO.Connection = CONEC
        CMDO.CommandText = "UPDATE PROVEEDORES set TELEFONO = '" + TextBox7.Text + "' WHERE ID = " + TextBox8.Text
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
        CMDO.CommandText = "select * from PROVEEDORES"
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
        CMDO.CommandText = "delete from PROVEEDORES WHERE COD_PROVEEDOR = " & TextBox9.Text
        ADAP.SelectCommand = CMDO
        ADAP.Fill(TABLA)
        DataGridView1.DataSource = TABLA
    End Sub
End Class