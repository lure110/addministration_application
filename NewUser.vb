Imports System.Data.OleDb
Imports System.IO
Imports System

Public Class NewUser
    ' Atgal
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub AfterClose() Handles MyBase.FormClosed
        If Application.OpenForms().OfType(Of all).Any Then
            all.listOfStudent = New ArrayList
            all.listOfStudent = Form1.person.getAllStudents
            all.LoadInfo(1)
        End If
    End Sub


    'Išsaugoti
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If (VardasTextBox.Text.Equals("") Or PavardėTextBox.Text.Equals("")) Then
            MsgBox("Vardas ir pavardė negali būti tušti")
            Return
        End If

        Try
            Form1.person.Pamaina = PamainaComboBox.SelectedItem.ToString()
        Catch ex As Exception
            Form1.person.Pamaina = ""
            MsgBox("Pamaina turi būti užpildyta")
            Return
        End Try

        Try
            Form1.person.Programa = Mokymosi_programaComboBox.SelectedItem.ToString()
        Catch ex As Exception
            Form1.person.Programa = ""
            MsgBox("Mokymosi programa turi būti užpildyta")
            Return
        End Try


        If (Form1.CheckThis(VardasTextBox.Text, PavardėTextBox.Text)) Then
            MsgBox("Toks žmogus jau yra...")
            Button3.Enabled = False
            Return
        End If

        Try

            Form1.person.Asm_kodas = Asmens_kodasTextBox.Text
            Form1.person.Vardas = VardasTextBox.Text
            Form1.person.Pavarde = PavardėTextBox.Text
            Form1.person.Mok_el = Mokinio_el_paštasTextBox.Text
            Form1.person.Grupe = TextBox1.Text 'grupes
            Form1.person.Mokykla = MokyklaTextBox.Text
            Form1.person.Klase = KlasėTextBox.Text
            Form1.person.Miestas = MiestasTextBox.Text
            Form1.person.Mamos_tel = Mamos_tel_numerisTextBox.Text
            Form1.person.Tevo_tel = Tėvo_tel_numerisTextBox.Text
            Form1.person.Mok_tel = Moksleivio_tel_numerisTextBox.Text
            Form1.person.Tevu_el = Tėvų_el_paštasTextBox.Text
            Form1.person.Pastabos = PastabosTextBox.Text
            Form1.person.Gyv_adressas = TextBox2.Text

            Try
                If (Marškinėlių_dydisComboBox.SelectedItem.ToString() IsNot Nothing) Then
                    Form1.person.marsD = Marškinėlių_dydisComboBox.SelectedItem.ToString()
                Else Form1.person.marsD = ""
                End If
            Catch ex As Exception
                Form1.person.marsD = ""
            End Try

            Try
                If (Iš_kur_sužinojoComboBox.SelectedItem.ToString() IsNot Nothing) Then
                    Form1.person.Suzino = Iš_kur_sužinojoComboBox.SelectedItem.ToString()
                Else Form1.person.Suzino = ""
                End If
            Catch ex As Exception
                Form1.person.Suzino = ""
            End Try

            If Form1.person.ValidateData() = False Then
                Return
            End If

            Form1.person.AIS = False
            Form1.person.Sutarties_Nr = ""




            Dim result As MsgBoxResult = False
            Dim chc As Boolean
            result = MsgBox(" Ar tikrai norite įkelti " + VardasTextBox.Text + " " + PavardėTextBox.Text, vbYesNo, "example")
            If result = MsgBoxResult.Yes Then
                chc = True
            ElseIf result = MsgBoxResult.No Then
                chc = False
            End If

            Form1.person.InsertPerson()

            'Jei viską įkelia tuomet galime atlaisvinti spausdinimą
            If (chc = True) Then


                LoadAll()
                'Form1.person.Sutarties_Nr = Sutarties_nrTextBox.Text
                'Form1.person.JKM_kodas = Label1.Text
                'Sutarties_nrTextBox.Text = Form1.person.Sutarties_Nr

                If (Form1.person.JKM_kodas Is Nothing Or Form1.person.JKM_kodas = "") Then
                    GenerateStudentID()
                    Form1.person.JKM_kodas = Label1.Text
                End If

                Form1.person.UpdatePerson()
                Button2.Enabled = False
                If (Form1.person.Vardas.Length > 0 And Form1.person.Asm_kodas.Length > 0 And
                    Form1.person.Pavarde.Length > 0 And Form1.person.Programa.Length > 0) Then
                    Button3.Enabled = True
                End If


            Else
                Button3.Enabled = False
            End If


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    'Spausdinti sutartį
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Print.Show()
    End Sub

    ' On load
    Private Sub NewUser_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' Užkrauti sąrašus
        Dim root = Form1.session.appFolder
        Dim connectionACString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Form1.session.rootwithAccdb

        Using cn As New OleDbConnection(connectionACString)

            'Programos
            Dim selectString As String = "SELECT * FROM [Mokymosi programos]"
            Dim cmd As OleDbCommand = New OleDbCommand(selectString, cn)
            cn.Open()
            Dim reader As OleDbDataReader = cmd.ExecuteReader()
            While reader.Read()
                Mokymosi_programaComboBox.Items.Add(reader.GetValue(1).ToString())
            End While
            cn.Close()

            'Marskiniai
            selectString = "SELECT * FROM [Dydis]"
            cmd = New OleDbCommand(selectString, cn)
            cn.Open()
            reader = cmd.ExecuteReader()
            While reader.Read()
                Marškinėlių_dydisComboBox.Items.Add(reader.GetValue(1).ToString())
            End While
            cn.Close()

            'Pamainos
            selectString = "SELECT * FROM [Pamaina]"
            cmd = New OleDbCommand(selectString, cn)
            cn.Open()
            reader = cmd.ExecuteReader()
            While reader.Read()
                PamainaComboBox.Items.Add(reader.GetValue(1).ToString())
            End While
            cn.Close()

            'Is kur suzinojo
            selectString = "SELECT * FROM [Informacija]"
            cmd = New OleDbCommand(selectString, cn)
            cn.Open()
            reader = cmd.ExecuteReader()
            While reader.Read()
                Iš_kur_sužinojoComboBox.Items.Add(reader.GetValue(1).ToString())
            End While
            cn.Close()
        End Using





    End Sub

    Private Sub MakeContractID()
        If (Sutarties_nrTextBox.Text Is Nothing Or Sutarties_nrTextBox.Text = "") Then


            Dim contc As Contract = New Contract
            ' cia id sudarom
            Dim currID As Integer = Integer.Parse(Form1.person.ID)
            contc.ID = Integer.Parse(Form1.person.ID)

            Dim contractID As String = currID
            contractID = Form1.person.GenerateNewContractID
            'If (contractID.Length <= 1) Then
            '    contractID = "0" + contractID
            'End If
            'If (contractID.Length <= 2) Then
            '    contractID = "0" + contractID
            'End If

            Edit.contract.ContractID = contractID
            Sutarties_nrTextBox.Text = contractID
            Form1.person.Sutarties_Nr = contractID
            '------
            IDTextBox.Text = Form1.person.ID


            '    cn.Close()
            'End Using
        End If
    End Sub

    Private Sub LoadAll()

        Dim root = Form1.session.appFolder
        Dim connectionACString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Form1.session.rootwithAccdb

        Dim output As New List(Of String)()

        Dim selectString As String = "SELECT * FROM Mokiniai WHERE Vardas = ? AND Pavardė = ?"

        Using cn As New OleDbConnection(connectionACString)
            Dim cmd As OleDbCommand = New OleDbCommand(selectString, cn)
            cmd.Parameters.AddWithValue("Vardas", VardasTextBox.Text)
            cmd.Parameters.AddWithValue("Pavardė", PavardėTextBox.Text)
            cn.Open()
            Dim reader As OleDbDataReader = cmd.ExecuteReader()

            Dim index As Integer = 0
            Form1.session.SetIndex(index)

            While reader.Read()
                Form1.person.ID = reader.GetValue(0).ToString()
                Form1.person.Asm_kodas = reader.GetValue(1).ToString()
                Form1.person.Vardas = reader.GetValue(2).ToString()
                Form1.person.Pavarde = reader.GetValue(3).ToString()
                Form1.person.marsD = reader.GetValue(4).ToString()
                Form1.person.Pamaina = reader.GetValue(5).ToString()
                Form1.person.Programa = reader.GetValue(6).ToString()
                Form1.person.Mok_el = reader.GetValue(7).ToString()
                Form1.person.Mokykla = reader.GetValue(8).ToString()
                Form1.person.Klase = reader.GetValue(9).ToString()
                Form1.person.Miestas = reader.GetValue(10).ToString()
                Form1.person.Mamos_tel = reader.GetValue(11).ToString()
                Form1.person.Tevo_tel = reader.GetValue(12).ToString()
                Form1.person.Mok_tel = reader.GetValue(13).ToString()
                Form1.person.Tevu_el = reader.GetValue(14).ToString()
                Form1.person.Suzino = reader.GetValue(15).ToString()
                Form1.person.AIS = reader.GetValue(16).ToString()
                Form1.person.Sutarties_Nr = reader.GetValue(17).ToString()
                Form1.person.JKM_kodas = reader.GetValue(19).ToString()
                Form1.person.Pastabos = reader.GetValue(18).ToString()
                Form1.person.Grupe = reader.GetValue(21).ToString()
                Form1.person.Gyv_adressas = reader.GetValue(22).ToString()
                Form1.person.Metai = reader.GetValue(23)
            End While
            cn.Close()
        End Using


        MakeContractID()

        If (Form1.person.JKM_kodas Is Nothing Or Form1.person.JKM_kodas = "") Then
            GenerateStudentID()
        End If
    End Sub

    Private Sub GenerateStudentID()
        Dim jkmID As String = GetJKMID()

        If Not jkmID.Equals("") Then
            Label1.Text = jkmID
            Exit Sub
        End If


        Dim rnd As New Random

        Dim ID As String = ""
        Dim Rn As String = "ABCDEFGHJKLMNPRSTUVYZ"

        Dim num As Integer = 0

        ID = GenerateProgrammeID(Mokymosi_programaComboBox.Text) ' Programos ID
        Dim TempID As String = Form1.person.Sutarties_Nr

        For i = TempID.Length To 4
            TempID = "0" + TempID
        Next

        ID = ID + TempID 'Registracijos ID

        num = rnd.Next(0, 20)
        ID = ID + Rn.Substring(num, 1)
        Dim T = Rn
        num = rnd.Next(0, 20)
        ID = ID + T.Substring(num, 1)


        Label1.Text = ID
    End Sub

    Private Function GenerateProgrammeID(var As String)
        Dim ID As String = ""

        Dim root = Form1.session.appFolder
        Dim connectionACString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Form1.session.rootwithAccdb
        Dim selectString As String = "SELECT * FROM `Mokymosi programos`"

        Using cn As New OleDbConnection(connectionACString)
            Dim cmd As OleDbCommand = New OleDbCommand(selectString, cn)
            cn.Open()
            Dim reader As OleDbDataReader = cmd.ExecuteReader()

            While reader.Read()
                Dim idtemp As String = reader.GetValue(1).ToString
                If (idtemp.Equals(var)) Then
                    ID = reader.GetValue(2).ToString
                End If
            End While
            cn.Close()
        End Using

        Return ID
    End Function

    Private Function GetJKMID()
        Dim ID As String = ""
        Dim root = Form1.session.appFolder
        Dim connectionACString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Form1.session.rootwithAccdb
        Dim selectString = "SELECT * FROM Mokiniai " &
                            "WHERE ID = @ID"
        Try
            Using cn As New OleDbConnection(connectionACString)
                Dim cmd As OleDbCommand = New OleDbCommand(selectString, cn)
                cmd.Parameters.AddWithValue("@ID", Form1.person.ID)
                cn.Open()
                Dim reader As OleDbDataReader = cmd.ExecuteReader()

                While reader.Read()
                    ID = reader.GetValue(19).ToString
                End While
                cn.Close()
            End Using
        Catch ex As Exception

        End Try

        Return ID
    End Function

End Class