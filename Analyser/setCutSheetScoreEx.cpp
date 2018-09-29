// setCutSheetScoreEx.h : CSetCutSheetScoreEx 类的实现



// CSetCutSheetScoreEx 实现

// 代码生成在 2005年8月31日, 10:34

#include "stdafx.h"
#include "setCutSheetScoreEx.h"
IMPLEMENT_DYNAMIC(CSetCutSheetScoreEx, CRecordset)

CSetCutSheetScoreEx::CSetCutSheetScoreEx(CDatabase* pdb)
	: CRecordset(pdb)
{
	m_ID			= 0;
	m_CutSheetID	= 0;
	m_TeacherID		= 0;
	m_Step			= 0;
	m_Score			= 0;
	m_Flag			= 0;
	m_nFields		= 6;

	m_nDefaultType	= snapshot;
}
//#error Security Issue: The connection string may contain a password
//// 此连接字符串中可能包含密码
//// 下面的连接字符串中可能包含明文密码和/或
//// 其他重要信息。请在查看完
//// 此连接字符串并找到所有与安全有关的问题后移除 #error。可能需要
//// 将此密码存储为其他格式或使用其他的用户身份验证。
//CString CSetCutSheetScoreEx::GetDefaultConnect()
//{
//	return _T("DSN=eTest;APP=Microsoft\x00ae Visual Studio .NET;WSID=SKY;DATABASE=eTest;LANGUAGE=\x7b80\x4f53\x4e2d\x6587;Trusted_Connection=Yes");
//}

CString CSetCutSheetScoreEx::GetDefaultSQL()
{
	return _T("[dbo].[CutSheetScoreEx]");
}

void CSetCutSheetScoreEx::DoFieldExchange(CFieldExchange* pFX)
{
	pFX->SetFieldType(CFieldExchange::outputColumn);
// RFX_Text() 和 RFX_Int() 这类宏依赖的是
// 成员变量的类型，而不是数据库字段的类型。
// ODBC 尝试自动将列值转换为所请求的类型
	RFX_Long(pFX, _T("[ID]"), m_ID);
	RFX_Long(pFX, _T("[CutSheetID]"), m_CutSheetID);
	RFX_Long(pFX, _T("[TeacherID]"), m_TeacherID);
	RFX_Byte(pFX, _T("[Step]"), m_Step);
	RFX_Byte(pFX, _T("[Score]"), m_Score);
	RFX_Byte(pFX, _T("[Flag]"), m_Flag);

}
/////////////////////////////////////////////////////////////////////////////
// CSetCutSheetScoreEx 诊断

#ifdef _DEBUG
void CSetCutSheetScoreEx::AssertValid() const
{
	CRecordset::AssertValid();
}

void CSetCutSheetScoreEx::Dump(CDumpContext& dc) const
{
	CRecordset::Dump(dc);
}
#endif //_DEBUG


