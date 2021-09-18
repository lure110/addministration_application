Imports System.Data.OleDb ' MS access pasiekti
Public Class Stats
    Public Function GetArrayOfStats()
        Dim stats(4) As Integer

        ' 0 - Visas mokinių skaičius duombazėje
        ' 1 - Tinkamų įkelti į AIS mokinių skaičius
        ' 2 - Pateiktų sutarčių skaičius
        Dim allStudent As ArrayList = Form1.person.getAllStudents
        stats(0) = allStudent.Count


        Dim root = Form1.session.appFolder
        Dim connectionACString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Form1.session.rootwithAccdb



        Using cn As New OleDbConnection(connectionACString)
            '1
            Dim selectString As String = "SELECT * FROM Mokiniai "
            Dim cmd As OleDbCommand = New OleDbCommand(selectString, cn)
            cn.Open()

            cn.Close()
            '2
            selectString = "SELECT COUNT(*) FROM Mokiniai WHERE [Sutarties nr] IS NOT NULL"
            cmd = New OleDbCommand(selectString, cn)
            cn.Open()
            stats(2) = cmd.ExecuteScalar
            MsgBox("Išviso mokinių duomenų bazėje: " + stats(0).ToString + Environment.NewLine + "Sutarčių kiekis: " + stats(2).ToString)
            cn.Close()
        End Using

        Return stats
    End Function
End Class
