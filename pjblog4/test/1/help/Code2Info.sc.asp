<%
function Error2Info(Error)
	select case Error
	case -1:Error2Info = "ERROR -1: �ϴ�û�п�ʼ������ 'Open' ��"
	case 0: Error2Info = "ERROR 0: �ϴ��ɹ���"
	case 1: Error2Info = "ERROR 1: �ϴ���Ч������һЩ�ļ������ 'MaxSize' ��δ�����档"
	case 2: Error2Info = "ERROR 2: �ϴ���Ч������һЩ�ļ���ƥ�� 'FileType' ��δ�����档"
	case 3: Error2Info = "ERROR 3: �ϴ���Ч������һЩ�ļ������ 'MaxSize' ���Ҳ�ƥ�� 'FileType' ��δ�����档"
	case 4: Error2Info = "ERROR 4: �쳣���������ϴ���"
	case 5: Error2Info = "ERROR 5: �ϴ��Ѿ�ȡ�������������������Ƿ�С�� 'TotalSize' ��"
	end select
end function

function Err2Info(Err)
	select case Err
	case -1:Err2Info = "ERR -1: û���ļ��ϴ���"
	case 0: Err2Info = "ERR 0: �ļ�����ɹ���"
	case 1: Err2Info = "ERR 1: �ļ������ 'MaxSize' ��δ�����档"
	case 2: Err2Info = "ERR 2: �ļ���ƥ�� 'FileType' ��δ�����档"
	case 3: Err2Info = "ERR 3: �ļ������ 'MaxSize' ���Ҳ�ƥ�� 'FileType' ��δ�����档"
	case 4:	Err2Info = "ERR 4: �ļ�����ʧ�ܡ�"	
	end select
end function
%>