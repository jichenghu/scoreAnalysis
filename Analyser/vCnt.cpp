// vCnt.h : CvCnt 类的实现



// CvCnt 实现

// 代码生成在 2005年9月28日, 16:56

#include "stdafx.h"
#include "vCnt.h"
IMPLEMENT_DYNAMIC(CvCnt, CRecordset)

CvCnt::CvCnt(CDatabase* pdb)
	: CRecordset(pdb)
{
	m_Cnt = 0;
//	m_TeacherID = 0;
	m_nFields = 1;
	m_nDefaultType = snapshot;
}
//#error Security Issue: The connection string may contain a password
// 此连接字符串中可能包含密码
// 下面的连接字符串中可能包含明文密码和/或
// 其他重要信息。请在查看完
// 此连接字符串并找到所有与安全有关的问题后移除 #error。可能需要
// 将此密码存储为其他格式或使用其他的用户身份验证。
//CString CvCnt::GetDefaultConnect()
//{
//	return _T("DSN=eTest;UID=sa;PWD=87211237,;APP=Microsoft\x00ae Visual Studio .NET;WSID=SKY;DATABASE=eTest;LANGUAGE=\x7b80\x4f53\x4e2d\x6587");
//}

CString CvCnt::GetDefaultSQL()
{
	return _T("[dbo].[vCntAccepted]");
}

void CvCnt::DoFieldExchange(CFieldExchange* pFX)
{
	pFX->SetFieldType(CFieldExchange::outputColumn);
// RFX_Text() 和 RFX_Int() 这类宏依赖的是
// 成员变量的类型，而不是数据库字段的类型。
// ODBC 尝试自动将列值转换为所请求的类型
	RFX_Long(pFX, _T("[Cnt]"), m_Cnt);
//	RFX_Long(pFX, _T("[TeacherID]"), m_TeacherID);

}
/////////////////////////////////////////////////////////////////////////////
// CvCnt 诊断

#ifdef _DEBUG
void CvCnt::AssertValid() const
{
	CRecordset::AssertValid();
}

void CvCnt::Dump(CDumpContext& dc) const
{
	CRecordset::Dump(dc);
}
#endif //_DEBUG


