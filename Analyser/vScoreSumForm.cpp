// vScoreSumForm.h : CvScoreSumForm 类的实现



// CvScoreSumForm 实现

// 代码生成在 2006年2月15日, 13:41

#include "stdafx.h"
#include "vScoreSumForm.h"
IMPLEMENT_DYNAMIC(CvScoreSumForm, CRecordset)

CvScoreSumForm::CvScoreSumForm(CDatabase* pdb)
	: CRecordset(pdb)
{
	m_StudentID			= 0		;
	m_QuestionFormID	= 0		;
	m_fValue			= 0.0	;

	m_nFields		= 3			;
	m_nDefaultType	= snapshot	;
}
//#error Security Issue: The connection string may contain a password
//// 此连接字符串中可能包含密码
//// 下面的连接字符串中可能包含明文密码和/或
//// 其他重要信息。请在查看完
//// 此连接字符串并找到所有与安全有关的问题后移除 #error。可能需要
//// 将此密码存储为其他格式或使用其他的用户身份验证。
//CString CvScoreSumForm::GetDefaultConnect()
//{
//	return _T("DSN=eTest;APP=Microsoft\x00ae Visual Studio .NET;WSID=SKY;DATABASE=eTest;LANGUAGE=\x7b80\x4f53\x4e2d\x6587;Trusted_Connection=Yes");
//}

CString CvScoreSumForm::GetDefaultSQL()
{
	return _T("[dbo].[vScoreSumForm]");
}

void CvScoreSumForm::DoFieldExchange(CFieldExchange* pFX)
{
	pFX->SetFieldType(CFieldExchange::outputColumn);
// RFX_Text() 和 RFX_Int() 这类宏依赖的是
// 成员变量的类型，而不是数据库字段的类型。
// ODBC 尝试自动将列值转换为所请求的类型
	RFX_Long(pFX, _T("[StudentID]"), m_StudentID);
	RFX_Long(pFX, _T("[QuestionFormID]"), m_QuestionFormID);
	RFX_Single(pFX, _T("[fValue]"), m_fValue);

}
/////////////////////////////////////////////////////////////////////////////
// CvScoreSumForm 诊断

#ifdef _DEBUG
void CvScoreSumForm::AssertValid() const
{
	CRecordset::AssertValid();
}

void CvScoreSumForm::Dump(CDumpContext& dc) const
{
	CRecordset::Dump(dc);
}
#endif //_DEBUG


