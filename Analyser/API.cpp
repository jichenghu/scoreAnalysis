
#include "stdafx.h"
#include "API.h"
#include "setSelectID.h"

BOOL TestDir(CString dir,CString& hint)
{
	DWORD result = GetFileAttributes(dir);

	// 不存在
	if(result == INVALID_FILE_ATTRIBUTES)
	{
		hint	= "无法访问存贮结果的目录\r\n请检查路径的设置是否正确";
		return	FALSE;
	}

	// 只读
	if(result & FILE_ATTRIBUTE_READONLY)
	{
		hint	= "无法写存贮结果的目录\r\n该目录被设置为只读属性\r\n请将属性改变为可写";
		return	FALSE;
	}

	//正确
	return TRUE;
}

BOOL TestDirectory( CString dir )
{
	DWORD result = GetFileAttributes(dir);

	// 不存在
	if(result == INVALID_FILE_ATTRIBUTES)
	{
		return	FALSE;
	}

	return TRUE;
}
