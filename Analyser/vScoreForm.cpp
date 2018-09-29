// vScoreForm.h : CvScoreForm 类的实现



// CvScoreForm 实现

// 代码生成在 2006年2月1日, 11:15

#include "stdafx.h"
#include "vScoreForm.h"
IMPLEMENT_DYNAMIC(CvScoreForm, CRecordset)

CvScoreForm::CvScoreForm(CDatabase* pdb)
	: CRecordset(pdb)
{
	m_ID = 0;
	m_StudentID = 0;
	m_FormID = 0;
	m_Score = 0.0;
	m_ClassID = 0;
	m_SubjectID = 0;
	m_nFields = 6;
	m_nDefaultType = snapshot;
}
//#error Security Issue: The connection string may contain a password
//// 此连接字符串中可能包含密码
//// 下面的连接字符串中可能包含明文密码和/或
//// 其他重要信息。请在查看完
//// 此连接字符串并找到所有与安全有关的问题后移除 #error。可能需要
//// 将此密码存储为其他格式或使用其他的用户身份验证。
//CString CvScoreForm::GetDefaultConnect()
//{
//	return _T("DSN=eTest;APP=Microsoft\x00ae Visual Studio .NET;WSID=SKY;DATABASE=eTest;LANGUAGE=\x7b80\x4f53\x4e2d\x6587;Trusted_Connection=Yes");
//}

CString CvScoreForm::GetDefaultSQL()
{
	return _T("[dbo].[vScoreForm]");
}

void CvScoreForm::DoFieldExchange(CFieldExchange* pFX)
{
	pFX->SetFieldType(CFieldExchange::outputColumn);
// RFX_Text() 和 RFX_Int() 这类宏依赖的是
// 成员变量的类型，而不是数据库字段的类型。
// ODBC 尝试自动将列值转换为所请求的类型
	RFX_Long(pFX, _T("[ID]"), m_ID);
	RFX_Long(pFX, _T("[StudentID]"), m_StudentID);
	RFX_Long(pFX, _T("[FormID]"), m_FormID);
	RFX_Single(pFX, _T("[Score]"), m_Score);
	RFX_Long(pFX, _T("[ClassID]"), m_ClassID);
	RFX_Int(pFX, _T("[SubjectID]"), m_SubjectID);

}
/////////////////////////////////////////////////////////////////////////////
// CvScoreForm 诊断

#ifdef _DEBUG
void CvScoreForm::AssertValid() const
{
	CRecordset::AssertValid();
}

void CvScoreForm::Dump(CDumpContext& dc) const
{
	CRecordset::Dump(dc);
}
#endif //_DEBUG


