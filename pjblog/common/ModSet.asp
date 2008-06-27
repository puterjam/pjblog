<%
'=================================================
' moduleSetting Class for PJBlog2
' Author: PuterJam
' UpdateDate: 2005-7-31
'=================================================

Class ModSet
    Private ModSetArray
    Private ModName
    Private state

    Private Sub Class_Initialize()

    End Sub

    Private Sub Class_Terminate()

    End Sub

    '=================================================
    ' 打开模块Open(ModName)
    '=================================================

    Public Function Open(LoadName)
        ModName = LoadName
        If Not IsArray(Application(CookieName&"_Mod_"&ModName))Then
            state = -18902
            ReLoad()
        Else
            ModSetArray = Application(CookieName&"_Mod_"&ModName)
            state = 0
        End If
    End Function

    '=================================================
    ' 从数据库里重新读取模块到缓存ReLoad()
    '=================================================

    Public Function ReLoad()
        If ModName = "" Then
            state = -18901
            Exit Function
        End If
        Dim ModDB, KeyLen, i, GetPlugPath
        i = 0
        KeyLen = conn.Execute("select count(*) from blog_ModSetting where set_ModName='"&ModName&"'")(0)
        Set ModDB = conn.Execute("select * from blog_ModSetting where set_ModName='"&ModName&"'")
        ReDim ModSetArray(KeyLen, 1)
        Do Until ModDB.EOF
            ModSetArray(i, 0) = ModDB("set_KeyName")
            ModSetArray(i, 1) = ModDB("set_KeyValue")
            i = i + 1
            ModDB.movenext
        Loop
        ModSetArray(KeyLen, 0) = "PlugingPath"
        Set GetPlugPath = conn.Execute("select InstallFolder from blog_module where name='"&ModName&"'")
        If GetPlugPath.EOF Then
            state = -18903
            Exit Function
        Else
            ModSetArray(KeyLen, 1) = GetPlugPath(0)
        End If
        Application.Lock
        Application(CookieName&"_Mod_"&ModName) = ModSetArray
        Application.UnLock
        state = 0
    End Function

    '=================================================
    ' 读取字段名称getKeyValue(KeyName)
    '=================================================

    Public Function getKeyValue(KeyName)
        Dim KeysLen, i
        getKeyValue = ""
        KeysLen = UBound(ModSetArray, 1)
        For i = 0 To KeysLen
            If ModSetArray(i, 0) = KeyName Then
                getKeyValue = ModSetArray(i, 1)
                Exit Function
            End If
        Next
    End Function

    '=================================================
    ' 获得出错信息ReLoad()
    '=================================================

    Public Function PasreError
        PasreError = state
        ' -18901 没有打开模块
        ' -18902 缓存里没有任何信息
        ' -18903 没有安装插件
    End Function

    '=================================================
    ' 获得插件所在路径
    '=================================================

    Public Function GetPath
        Dim KeysLen, i
        GetPath = ""
        KeysLen = UBound(ModSetArray, 1)
        GetPath = ModSetArray(KeysLen, 1)
    End Function

    '=================================================
    ' 清除插件占用的 Application 地址
    '=================================================

    Public Function RemoveApplication
        Application.Lock
        Application.Contents.Remove(CookieName&"_Mod_"&ModName)
        Application.UnLock
    End Function

End Class
%>
