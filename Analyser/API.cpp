
#include "stdafx.h"
#include "API.h"
#include "setSelectID.h"

BOOL TestDir(CString dir,CString& hint)
{
	DWORD result = GetFileAttributes(dir);

	// ������
	if(result == INVALID_FILE_ATTRIBUTES)
	{
		hint	= "�޷����ʴ��������Ŀ¼\r\n����·���������Ƿ���ȷ";
		return	FALSE;
	}

	// ֻ��
	if(result & FILE_ATTRIBUTE_READONLY)
	{
		hint	= "�޷�д���������Ŀ¼\r\n��Ŀ¼������Ϊֻ������\r\n�뽫���Ըı�Ϊ��д";
		return	FALSE;
	}

	//��ȷ
	return TRUE;
}

BOOL TestDirectory( CString dir )
{
	DWORD result = GetFileAttributes(dir);

	// ������
	if(result == INVALID_FILE_ATTRIBUTES)
	{
		return	FALSE;
	}

	return TRUE;
}
