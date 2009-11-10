<%
'*************** PJblog4 基本设置 *******************
' PJblog4 Copyright 2009
' Update: 2009-11-03
' Author: evio
' Owner : PuterJam
' Mode  : Class
' MoName: Access Connect
'***************************************************

Class Sys_Connection

	Private pj_version, pj_UpdateTime, pj_CookieName, pj_AccessFile
	Private pj_Conn

	' ***********************************************
	'	设置或获取当前版本号
	' ***********************************************
	Public Property Let version(ByVal Str)
		pj_version = Str
	End Property
	Public Property Get version
		version = pj_version
	End Property
	
	' ***********************************************
	'	设置或获取当前版本更新时间
	' ***********************************************
	Public Property Let UpdateTime(ByVal Str)
		pj_UpdateTime = Str
	End Property
	Public Property Get UpdateTime
		UpdateTime = pj_UpdateTime
	End Property
	
	' ***********************************************
	'	设置或获取CookieName的属性方法
	' ***********************************************
	Public Property Let CookieName(ByVal Str)
		pj_CookieName = Str
	End Property
	Public Property Get CookieName
		CookieName = pj_CookieName
	End Property
	
	' ***********************************************
	'	设置或获取数据库地址的方法
	' ***********************************************
	Public Property Let AccessFile(ByVal Str)
		pj_AccessFile = Str
	End Property
	Public Property Get AccessFile
		AccessFile = pj_AccessFile
	End Property
	
	' ***********************************************
	'	数据库连接类初始化
	' ***********************************************
	Private Sub Class_Initialize
    End Sub 
     
	' ***********************************************
	'	数据库连接类终结化
	' ***********************************************
    Private Sub Class_Terminate
    End Sub
	
	' ***********************************************
	'	解决数据库连接问题
	' ***********************************************
	Private Function SolveConnection(ByVal ConnectionString, ByVal i)
		On Error Resume Next
		Dim j, Note, Str
		Note = "../" : Str = ""
		For j = 1 To i
			Str = Str & "../"
		Next
		Set pj_Conn = Server.CreateObject("ADODB.CONNECTION")
		pj_Conn.open("provider=Microsoft.jet.oledb.4.0;data source=" & Server.MapPath(Str & ConnectionString))
		If Err Then
			Err.Clear
			i = i + 1
			Set pj_Conn = SolveConnection(ConnectionString, i)
		End If
		Set SolveConnection = pj_Conn
	End Function
	
	' ***********************************************
	'	连接数据库
	' ***********************************************
	Public Function open
		Set open = SolveConnection(AccessFile, 0)
	End Function
	
	' ***********************************************
	'	关闭数据库
	' ***********************************************
	Public Sub Close
		pj_Conn.Close
		Set pj_Conn = Nothing
	End Sub
	
	' ***********************************************
	'	数据库增加修改操作
	' ***********************************************
	Public Function doRecord(ByVal table, ByVal DBArray, ByVal Action, ByVal Primarykey, ByVal ID)
		On Error Resume Next
		Dim AddCount, TempDB, i, v, Sql, BackPrimarykeyID
		If LCase(Action) <> "insert" And LCase(Action) <> "update" Then Action = "insert"
		If LCase(Action) = "insert" Then v = 2 Else v = 3
		If Not IsArray(DBArray) Then
			doRecord = Array(False, lang.Set.Asp(3))
			Exit Function
		Else
			pj_Conn.BeginTrans
			Set TempDB = Server.CreateObject("ADODB.RecordSet")
				If LCase(Action) = "insert" Then
					Sql = table
				Else
					Sql = "Select * From " & table & " Where " & Primarykey & "=" & ID
				End If
				TempDB.Open Sql, pj_Conn, 1, v
				If LCase(Action) = "insert" Then TempDB.AddNew : BackPrimarykeyID = TempDB(Primarykey)
				AddCount = UBound(DBArray, 1)
				For i = 0 To AddCount
					TempDB(DBArray(i)(0)) = DBArray(i)(1)
				Next
				TempDB.Update
			TempDB.Close
			Set TempDB = Nothing
			If Err.Number > 0 Then
				pj_Conn.RollBackTrans
				doRecord = Array(False, Err.Description)
			Else
				pj_Conn.CommitTrans
				doRecord = Array(True, BackPrimarykeyID)
			End If
		End If
	End Function
	
	' ***********************************************
	'	数据库删除操作
	'	DBArray 	: Array
	'	How To Use 	: Array(Array(Table, Primarykey, ID), ...)
	' ***********************************************
	Public Function doRecDel(ByVal DBArray)
		On Error Resume Next
		Dim AddCount, i, TempDB, Sql
		If Not IsArray(DBArray) Then
			doRecDel = Array(False, lang.Set.Asp(3))
			Exit Function
		Else
			pj_Conn.BeginTrans
			Set TempDB = Server.CreateObject("ADODB.RecordSet")
				AddCount = UBound(DBArray, 1)
				For i = 0 To AddCount
					Sql = "Select * From " & DBArray(i)(0) & " Where " & DBArray(i)(1) & "=" & DBArray(i)(2)
					TempDB.open Sql, pj_Conn, 1, 3
					TempDB.Delete
					TempDB.Close
				Next
			Set TempDB = Nothing
			If Err.Number > 0 Then
				pj_Conn.RollBackTrans
				doRecDel = Array(False, Err.Description)
			Else
				pj_Conn.CommitTrans
				doRecDel = Array(True, lang.Set.Asp(1))
			End If
		End If
	End Function
	
	' ***********************************************
	'	数据库查询单个信息
	' ***********************************************
	Public Function doGet(ByVal Str)
		doGet = pj_Conn.Execute(Str)
	End Function
	
End Class
%>