<%
function Error2Info(Error)
	select case Error
	case -1:Error2Info = "ERROR -1: 上传没有开始，请先 'Open' 。"
	case 0: Error2Info = "ERROR 0: 上传成功。"
	case 1: Error2Info = "ERROR 1: 上传生效，但有一些文件因大于 'MaxSize' 而未被保存。"
	case 2: Error2Info = "ERROR 2: 上传生效，但有一些文件因不匹配 'FileType' 而未被保存。"
	case 3: Error2Info = "ERROR 3: 上传生效，但有一些文件因大于 'MaxSize' 并且不匹配 'FileType' 而未被保存。"
	case 4: Error2Info = "ERROR 4: 异常，不存在上传。"
	case 5: Error2Info = "ERROR 5: 上传已经取消，请检查总上载数据是否小于 'TotalSize' 。"
	end select
end function

function Err2Info(Err)
	select case Err
	case -1:Err2Info = "ERR -1: 没有文件上传。"
	case 0: Err2Info = "ERR 0: 文件保存成功。"
	case 1: Err2Info = "ERR 1: 文件因大于 'MaxSize' 而未被保存。"
	case 2: Err2Info = "ERR 2: 文件因不匹配 'FileType' 而未被保存。"
	case 3: Err2Info = "ERR 3: 文件因大于 'MaxSize' 并且不匹配 'FileType' 而未被保存。"
	case 4:	Err2Info = "ERR 4: 文件保存失败。"	
	end select
end function
%>