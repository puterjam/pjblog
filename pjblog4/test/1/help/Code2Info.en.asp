<%
function Error2Info(Error)
	select case Error
	case -1:Error2Info = "ERROR -1: UpLoad is not started, please 'Open' first."
	case 0: Error2Info = "ERROR 0: Upload is success."
	case 1: Error2Info = "ERROR 1: Upload is taken effect, but some file is not saveed for mismatching 'MaxSize'."
	case 2: Error2Info = "ERROR 2: Upload is taken effect, but some file is not saveed for mismatching 'FileType'."
	case 3: Error2Info = "ERROR 3: Upload is taken effect, but some file is not saveed for mismatching 'MaxSize' and 'FileType'."
	case 4: Error2Info = "ERROR 4: Exception, upload is not exist."
	case 5: Error2Info = "ERROR 5: Upload is canceled, please check the requestsize is less than 'TotalSize'."
	end select
end function

function Err2Info(Err)
	select case Err
	case -1:Err2Info = "ERR -1: No file is to be upload."
	case 0: Err2Info = "ERR 0: File is saveed success."
	case 1: Err2Info = "ERR 1: File is not saveed for mismatching 'MaxSize'."
	case 2: Err2Info = "ERR 2: File is not saveed for mismatching 'FileType'."
	case 3: Err2Info = "ERR 3: File is not saveed for mismatching 'MaxSize' and 'FileType'."
	case 4:	Err2Info = "ERR 4: File is saveed failed."	
	end select
end function
%>