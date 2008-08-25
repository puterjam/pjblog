<%@CodePage="65001"%>
<script language="vbscript" runat="server">
	Option Explicit '显示声明
	Dim showErrorImg
	showErrorImg = False

	function ChkPost() 
		dim server_v1,server_v2
		chkpost=false
		server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
		server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
		if instr(server_v1, replace(replace(server_v2, "http://", ""), "www.", ""))=0 then
'		If Mid(server_v1,8,Len(server_v2))<>server_v2 then
			chkpost=False
		else
			chkpost=True
		end If
	end Function

	If Not ChkPost() Then
		showErrorImg = true
	End If
	
	
    '''''''''''''''''''''''''''''''''''''''''''''
    ' Author: Layen support@ssaw.net 84815733(QQ)
    ' Thanks: Laomi, Laomiao, NetRube
    '                                2006-01-02
    ' URL:    http://www.layen.cn/
    '''''''''''''''''''''''''''''''''''''''''''''

	Class Com_GifCode_Class
		Public Noisy, Count, Width, Height, Angle, Offset, Border
		Private Graph(), Margin(3)
		Private Sub Class_Initialize()
			Randomize 
			Noisy = 3 ' 干扰点出现的概率 
			Count = 4  ' 字符数量
			Width = 40 ' 图片宽度
			Height = 20 ' 图片高度
			Angle = 2  ' 角度随机变化量
			Offset = 15 ' 偏移随机变化量
			Border = 1.5 ' 边框大小
		End Sub 
		
		Public Function Create2(str)
			Dim i
			Dim vIndex
			ReDim Graph(Width-1, Height-1)
			For i = 0 To Count - 1
				vIndex=CInt(Mid(str,i+1,1)-1)
				SetDraw vIndex, i
			Next
		End Function
	
		Sub SetDot(pX, pY)
			If pX * (Width-pX-1) >= 0 And pY * (Height-pY-1) >= 0 Then
				Graph(pX, pY) = 1
			End If
		End Sub

		Public Sub SetDraw(pIndex, pNumber)
			' 字符数据字典
			If pIndex=-1 Then pIndex=9

			Dim DotData(9)
			DotData(0) = Array(30, 15, 50, 1, 50, 100)
			DotData(1) = Array(1 ,34 ,30 ,1 ,71, 1, 100, 34, 1, 100, 93, 100, 100, 86) 
			DotData(2) = Array(1, 1, 100, 1, 42, 42, 100, 70, 50, 100, 1, 70)
			DotData(3) = Array(100, 73, 6, 73, 75, 6, 75, 100)
			DotData(4) = Array(100, 1, 1, 1, 1, 50, 50, 35, 100, 55, 100, 80, 50, 100, 1, 95)
			DotData(5) = Array(100, 20, 70, 1, 20, 1, 1, 30, 1, 80, 30, 100, 70, 100, 100, 80, 100, 60, 70, 50, 30, 50, 1, 60) 
			DotData(6) = Array(6, 26, 6, 6, 100, 6, 53, 100)
			DotData(7) = Array(100, 30, 100, 20, 70, 1, 30, 1, 1, 20, 1, 30, 100, 70, 100, 80, 70, 100, 30, 100, 1, 80, 1, 70, 100, 30)
			DotData(8) = Array(1, 80, 30, 100, 80, 100, 100, 70, 100, 20, 70, 1, 30, 1, 1, 20, 1, 40, 30, 50, 70, 50, 100, 40) 
			DotData(9) = Array(100, 20, 70, 1, 20, 1, 1, 30, 1, 80, 30, 100, 70, 100, 100, 80, 100, 60, 90, 20, 80,3)


			Dim vExtent : vExtent = Width / Count
			
			Margin(0) = Border + vExtent * (Rnd * Offset) / 100 + Margin(1)
			
			Margin(1) = vExtent * (pNumber + 1) - Border - vExtent * (Rnd * Offset) / 100
			
			Margin(2) = Border + Height * (Rnd * Offset) / 100
			
			Margin(3) = Height - Border - Height * (Rnd * Offset) / 100

			Dim vStartX, vEndX, vStartY, vEndY
			Dim vWidth, vHeight, vDX, vDY, vDeltaT
			Dim vAngle, vLength

			vWidth =Int(Margin(1) - Margin(0))
			vHeight =Int(Margin(3) - Margin(2))
			vStartX = Int((DotData(pIndex)(0)-1) * vWidth / 100)
			vStartY = Int((DotData(pIndex)(1)-1) * vHeight / 100)

			Dim i, j
			For i = 1 To UBound(DotData(pIndex), 1)/2
				If DotData(pIndex)(2*i-2) <> 0 And DotData(pIndex)(2*i) <> 0 Then 
					vEndX = (DotData(pIndex)(2*i)-1) * vWidth / 100
					vEndY = (DotData(pIndex)(2*i+1)-1) * vHeight / 100
					vDX = vEndX - vStartX
					vDY = vEndY - vStartY
					If vDX = 0 Then
						vAngle = Sgn(vDY) * 3.14 /2
					Else
						vAngle = Atn(vDY / vDX)
					End If
					If Sin(vAngle) = 0 Then
						vLength = vDX
					Else
						vLength = vDY / Sin(vAngle)
					End If
					vAngle = vAngle + (Rnd - 0.5) * 2 * Angle * 3.14 * 2 / 100
					vDX = Int(Cos(vAngle) * vLength)
					vDY = Int(Sin(vAngle) * vLength)
					If Abs(vDX) > Abs(vDY) Then vDeltaT = Abs(vDX) Else vDeltaT = Abs(vDY)
					For j = 1 To vDeltaT
						SetDot Margin(0) + vStartX + j * vDX / vDeltaT, Margin(2) + vStartY + j * vDY / vDeltaT 
					Next
					vStartX = vStartX + vDX
					vStartY = vStartY + vDY
				End If
			Next
		End Sub

		Public Sub Output()
			Response.Expires = -9999
			Response.AddHeader "pragma", "no-cache"
			Response.AddHeader "cache-ctrol", "no-cache"
			Response.ContentType = "image/gif" 
			Response.BinaryWrite ChrB(Asc("G")) & ChrB(Asc("I")) & ChrB(Asc("F"))
			Response.BinaryWrite ChrB(Asc("8")) & ChrB(Asc("9")) & ChrB(Asc("a")) 
			Response.BinaryWrite ChrB(Width Mod 256) & ChrB((Width \ 256) Mod 256)
			Response.BinaryWrite ChrB(Height Mod 256) & ChrB((Height \ 256) Mod 256)
			Response.BinaryWrite ChrB(128) & ChrB(0) & ChrB(0) 
			Response.BinaryWrite ChrB(250) & ChrB(250) & ChrB(250)
			Response.BinaryWrite ChrB(0) & ChrB(0) & ChrB(0)
			Response.BinaryWrite ChrB(Asc(","))
			Response.BinaryWrite ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) 
			Response.BinaryWrite ChrB(Width Mod 256) & ChrB((Width \ 256) Mod 256)
			Response.BinaryWrite ChrB(Height Mod 256) & ChrB((Height \ 256) Mod 256)
			Response.BinaryWrite ChrB(0) & ChrB(7) & ChrB(255) 

			Dim x, y, i : i = 0
			For y = 0 To Height - 1
				For x = 0 To Width - 1
					If Rnd < Noisy / 100 Then
						Response.BinaryWrite ChrB(1-Graph(x, y))
					ElseIf x * (x-Width) = 0 Or y * (y-Height) = 0 Then 
						Response.BinaryWrite ChrB(Graph(x, y))
					ElseIf Graph(x-1, y) = 1 Or Graph(x, y) Or Graph(x, y-1) = 1 Then
						Response.BinaryWrite ChrB(1)
					Else
						Response.BinaryWrite ChrB(0)
					End If

					If (y * Width + x + 1) Mod 126 = 0 Then
						Response.BinaryWrite ChrB(128)
						i = i + 1
					End If
					If (y * Width + x + i + 1) Mod 255 = 0 Then
						If (Width*Height - y * Width - x - 1) > 255 Then 
							Response.BinaryWrite ChrB(255)
						Else
							Response.BinaryWrite ChrB(Width * Height Mod 255)
						End If
					End If
				Next
			Next
			Response.BinaryWrite ChrB(128) & ChrB(0) & ChrB(129) & ChrB(0) & ChrB(59) 
		End Sub
	End Class
	''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	
	Dim mCode
	Set mCode = New Com_GifCode_Class

    Randomize
    Dim code
    code = Int(Rnd * 9000 + 1000)
	If showErrorImg Then
		Session("GetCode") = Int(Rnd * 9000 + 1000)
	else
		Session("GetCode") = code
	End If
		
	mCode.Create2(code)
	mCode.Output()

	Set mCode = Nothing
</script>
