// setCutSheetEx.h : CSetCutSheetEx 类的实现



// CSetCutSheetEx 实现

// 代码生成在 2005年8月31日, 10:29

#include "stdafx.h"
#include "setCutSheetEx.h"
IMPLEMENT_DYNAMIC(CSetCutSheetEx, CRecordset)

CSetCutSheetEx::CSetCutSheetEx(CDatabase* pdb)
	: CRecordset(pdb)
{
	m_CutSheetID = 0;
	m_SubjectID = 0;
	m_RoomID = 0;
	m_SheetID = 0;
	m_QuestionID = 0;
	m_Flag = 0;
	m_nFields = 6;
	m_nDefaultType = snapshot;
}
//#error Security Issue: The connection string may contain a password
//// 此连接字符串中可能包含密码
//// 下面的连接字符串中可能包含明文密码和/或
//// 其他重要信息。请在查看完
//// 此连接字符串并找到所有与安全有关的问题后移除 #error。可能需要
//// 将此密码存储为其他格式或使用其他的用户身份验证。
//CString CSetCutSheetEx::GetDefaultConnect()
//{
//	return _T("DSN=eTest;UID=sa;PWD=87211237,;APP=Microsoft\x00ae Visual Studio .NET;WSID=SKY;DATABASE=eTest;LANGUAGE=\x7b80\x4f53\x4e2d\x6587");
//}

CString CSetCutSheetEx::GetDefaultSQL()
{
	return _T("[dbo].[CutSheetEx]");
}

void CSetCutSheetEx::DoFieldExchange(CFieldExchange* pFX)
{
	pFX->SetFieldType(CFieldExchange::outputColumn);
// RFX_Text() 和 RFX_Int() 这类宏依赖的是
// 成员变量的类型，而不是数据库字段的类型。
// ODBC 尝试自动将列值转换为所请求的类型
	RFX_Long(pFX, _T("[CutSheetID]"), m_CutSheetID);
	RFX_Int(pFX, _T("[SubjectID]"), m_SubjectID);
	RFX_Long(pFX, _T("[RoomID]"), m_RoomID);
	RFX_Long(pFX, _T("[SheetID]"), m_SheetID);
	RFX_Long(pFX, _T("[QuestionID]"), m_QuestionID);
	RFX_Byte(pFX, _T("[Flag]"), m_Flag);

}
/////////////////////////////////////////////////////////////////////////////
// CSetCutSheetEx 诊断

#ifdef _DEBUG
void CSetCutSheetEx::AssertValid() const
{
	CRecordset::AssertValid();
}

void CSetCutSheetEx::Dump(CDumpContext& dc) const
{
	CRecordset::Dump(dc);
}
#endif //_DEBUG


