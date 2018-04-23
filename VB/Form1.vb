Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports System.IO

Namespace WindowsApplication1
	Partial Public Class Form1
		Inherits Form
				Private Function CreateTable(ByVal RowCount As Integer) As DataTable
			Dim tbl As New DataTable()
			tbl.Columns.Add("Name", GetType(String))
			tbl.Columns.Add("ID", GetType(Integer))
			tbl.Columns.Add("Number", GetType(Integer))
			tbl.Columns.Add("Date", GetType(DateTime))
			For i As Integer = 0 To RowCount - 1
				tbl.Rows.Add(New Object() { String.Format("Name{0}", i), i, 3 - i, DateTime.Now.AddDays(i) })
			Next i
			Return tbl
				End Function


		Public Sub New()
			InitializeComponent()
			gridControl1.DataSource = CreateTable(20)
		End Sub

		Public Sub WriteLog(ByVal testName As String, ByVal result As Object)
			memoEdit1.Text += String.Format("[{0}] = {1}{2}", testName, result, Environment.NewLine)
		End Sub

		Public Function IsPrintingAvailable() As Boolean
			Return gridControl1.IsPrintingAvailable
		End Function

		Private Shared Function GetFileName() As String
			Dim fileDialog As FileDialog = New SaveFileDialog()
			fileDialog.FileName = "C:\test.xls"
			fileDialog.ShowDialog()
			Dim fileName As String = fileDialog.FileName
			Return fileName
		End Function
		Public Function ApplicationHasRequiredRights() As Object
			Dim result As Object
			Try
				Dim fileName As String = String.Format("{0}EmptyFile.xls", targetFileName)
				File.Create(fileName)
				Return File.Exists(fileName)

			Catch ex As Exception
			 result = ex.Message
			End Try
			Return result
		End Function

		Public Function IsFileCreated() As Object
			Dim fileName As String = String.Format("{0}GridControl.xls", targetFileName)
			gridControl1.ExportToXls(fileName)
			Return File.Exists(fileName)
		End Function
		Private Shared targetFileName As String
		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			targetFileName = GetFileName()
			WriteLog("IsPrintingAvailable", IsPrintingAvailable())
			WriteLog("ApplicationHasRequiredRights", ApplicationHasRequiredRights())
			WriteLog("IsFileCreated", IsFileCreated())
		End Sub
	End Class
End Namespace