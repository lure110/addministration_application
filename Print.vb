Imports System.IO
Imports Word = Microsoft.Office.Interop.Word

Public Class Print

	Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

	End Sub

	' Spausdinti / Trišalė 
	Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
		Print(True)
	End Sub

	Private Sub Print_Load(sender As Object, e As EventArgs) Handles MyBase.Load

		Label1.Text = "Asmens kodas: " + Form1.person.Asm_kodas
		Label2.Text = "Asmens vardas: " + Form1.person.Vardas
		Label3.Text = "Asmens pavardė: " + Form1.person.Pavarde
		Label5.Text = "Sutarties numeris: " + Form1.person.Sutarties_Nr
		'Form1.person.Sutarties_Nr = Edit.contract.ContractID
		Label6.Text = "Programa: " + Form1.person.Programa

	End Sub

	' Atgal
	Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
		Me.Close()
	End Sub

	' Spausdinimo sablonas
	Private Sub Print(flag As Boolean)
		If (TextBox1.Text.Length < 3) Then
			MsgBox("Įveskite atstovo vardą ir pavardę")
			Exit Sub
		End If
		If Not TextBox3.Text.Length = 11 Then
			MsgBox("Netinkamas asmens kodas ilgis(" & TextBox3.Text.Length & ")")
			Exit Sub
		End If




		Dim year As String = Date.Today.Year.ToString
		Dim fullYear = year
		Dim nextfullYear As String = Date.Today.AddYears(1).Year.ToString
		year = year.Substring(2)
		Dim todayDate As String = "2021-04-25"
		Dim ddate As DateTime = Date.Today
		todayDate = ddate.ToString("yyyy-MM-dd")

		Dim contractID As String = Form1.person.Sutarties_Nr

		Dim Course As String = Form1.person.Programa

		' Vaikas 
		Dim firstName As String = Form1.person.Vardas
		Dim lastName As String = Form1.person.Pavarde
		Dim PersonID As String = Form1.person.Asm_kodas
		Dim StudentPhone As String = Form1.person.Mok_tel
		Dim StudentEmail As String = Form1.person.Mok_el

		' Atstovas
		Dim Parent As String = TextBox1.Text
		' Kazkaip reikia apdoroti
		Dim LivingAddress As String = Form1.person.Gyv_adressas
		Dim ParentID As String = TextBox3.Text
		Dim ParentPhone As String = "+37010111011"
		If Form1.person.Mamos_tel.Length < 5 Then
			ParentPhone = Form1.person.Tevo_tel
		Else
			ParentPhone = Form1.person.Mamos_tel
		End If
		Dim ParentEmail As String = Form1.person.Tevu_el




		Dim TodayDateM_DD As String = ddate.ToString("MM-dd")

		If Not Directory.Exists(Form1.session.appFolder + "\sutartys") Then
			Directory.CreateDirectory(Form1.session.appFolder + "\sutartys")
		End If

		Dim ContractDOCX As String = Form1.session.appFolder + "\sutartys\" + lastName + "_" + firstName + "_" + year + ".docx"
		Dim ContractPdf As String = Form1.session.appFolder + "\sutartys\" + lastName + "_" + firstName + "_" + year + ".pdf"

		Dim objWordApp As Word.Application = Nothing
		Try

			objWordApp = New Word.Application


			' Atveriam docx
			Dim str As String = ""

			' Auto atskirimui kazkodel kai kada neveikia!!
			' Galibuti neranda Farom1.person.Metai reiksmes
			'If Form1.person.Metai > 13 Then
			'	str = Form1.session.appFolder + "\sut.docx"
			'Else
			'	str = Form1.session.appFolder + "\sut13.docx"
			'End If

			If flag = True Then
				str = Form1.session.appFolder + "\sut.docx"
			ElseIf flag = False Then
				str = Form1.session.appFolder + "\sut13.docx"
			End If

			' Gal jau atidarytas?

			Dim objDoc As Word.Document
			Dim AppId = DateTime.Now.Ticks()
			objWordApp.Application.Caption = AppId
			objWordApp.Application.Visible = True

			Dim WordPid = Process.GetProcessesByName(AppId)
			objWordApp.Application.Visible = False

			objDoc = objWordApp.Documents.OpenNoRepairDialog(str)


			objDoc = objWordApp.ActiveDocument
			' Ieškome ir pakeičiame
			objDoc.Content.Find.Execute(FindText:="[firstName]", ReplaceWith:=firstName, Replace:=Word.WdReplace.wdReplaceAll)
			objDoc.Content.Find.Execute(FindText:="[lastName]", ReplaceWith:=lastName, Replace:=Word.WdReplace.wdReplaceAll)
			objDoc.Content.Find.Execute(FindText:="[year]", ReplaceWith:=year, Replace:=Word.WdReplace.wdReplaceAll)
			objDoc.Content.Find.Execute(FindText:="[contractID]", ReplaceWith:=contractID, Replace:=Word.WdReplace.wdReplaceAll)
			objDoc.Content.Find.Execute(FindText:="[personID]", ReplaceWith:=PersonID, Replace:=Word.WdReplace.wdReplaceAll)
			objDoc.Content.Find.Execute(FindText:="[course]", ReplaceWith:=Course, Replace:=Word.WdReplace.wdReplaceAll)
			objDoc.Content.Find.Execute(FindText:="[parentFirstLastName]", ReplaceWith:=Parent, Replace:=Word.WdReplace.wdReplaceAll)

			objDoc.Content.Find.Execute(FindText:="[livingAddress]", ReplaceWith:=LivingAddress, Replace:=Word.WdReplace.wdReplaceAll)
			objDoc.Content.Find.Execute(FindText:="[todayMonth_day]", ReplaceWith:=TodayDateM_DD, Replace:=Word.WdReplace.wdReplaceAll)
			objDoc.Content.Find.Execute(FindText:="[studentPhone]", ReplaceWith:=StudentPhone, Replace:=Word.WdReplace.wdReplaceAll)
			objDoc.Content.Find.Execute(FindText:="[studentEmail]", ReplaceWith:=StudentEmail, Replace:=Word.WdReplace.wdReplaceAll)
			objDoc.Content.Find.Execute(FindText:="[parentID]", ReplaceWith:=ParentID, Replace:=Word.WdReplace.wdReplaceAll)
			objDoc.Content.Find.Execute(FindText:="[parentPhone]", ReplaceWith:=ParentPhone, Replace:=Word.WdReplace.wdReplaceAll)
			objDoc.Content.Find.Execute(FindText:="[parentEmail]", ReplaceWith:=ParentEmail, Replace:=Word.WdReplace.wdReplaceAll)



			'Save and close
			If File.Exists(ContractPdf) Then File.Delete(ContractPdf)

			'objDoc.SaveAs2(ContractDOCX)
			objDoc.SaveAs2(ContractPdf, Word.WdSaveFormat.wdFormatPDF)

			objDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges)


			KillPid("WINWORD")




			Process.Start(ContractPdf)

		Catch ex As Exception
			MsgBox(ex.Message)
		Finally
			If Not IsNothing(objWordApp) Then
				'objWordApp.Quit()
				objWordApp = Nothing
			End If
		End Try
		GC.Collect()
		GC.WaitForPendingFinalizers()


	End Sub

	' Mygtukas spausdinti dvišalė
	Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
		Print(False)
	End Sub

	Private Sub KillPid(name As String)
		For Each Proc In Process.GetProcesses
			If Process.GetCurrentProcess.Id = Proc.Id Then
				Continue For
			End If
			If Proc.ProcessName.Equals(name) Then
				If Proc.MainWindowTitle.Equals("") Then
					Proc.Kill()
				End If
			End If
		Next

	End Sub
End Class