Public Class Form1

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim directorio As IO.FileInfo() = New IO.DirectoryInfo(TextBox1.Text).GetFiles()
        Dim count As Integer = directorio.Length
        MsgBox(count)
        Dim fichero As IO.FileInfo
        ProgressBar1.Maximum = count
        ProgressBar1.Value = 0

        For Each fichero In directorio
            'MsgBox(fichero.FullName)
            'MsgBox(fichero.Name)
            ProgressBar1.Value = ProgressBar1.Value + 1
            If fichero.Name.Contains(TextBox2.Text) Then
                My.Computer.FileSystem.RenameFile(fichero.FullName, fichero.Name.Replace(TextBox2.Text, TextBox3.Text))
            End If
        Next
    End Sub
End Class
