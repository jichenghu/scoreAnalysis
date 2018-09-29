// setQuestionInfo.h : CSetQuestionInfo 类的实现



// CSetQuestionInfo 实现

// 代码生成在 2005年8月31日, 16:23

#include "stdafx.h"
#include "setQuestionInfo.h"
IMPLEMENT_DYNAMIC(CSetQuestionInfo, CRecordset)

CSetQuestionInfo::CSetQuestionInfo(CDatabase* pdb)
	: CRecordset(pdb)
{
	m_QuestionID = 0;
	m_SubjectID = 0;
	m_QuestionNum = 0;
	m_QuestionType = 0;
	m_ReservedMark = 0.0;
	m_WorkLoad = 0;
	m_Gap = 0;
	m_Note = "";
	m_QuestionTitle = "";
	m_SubQuestion = "";
//	m_CntStep = 0;
	m_nFields = 10;
	m_nDefaultType = snapshot;
}
//#error Security Issue: The connection string may contain a password
//// 此连接字符串中可能包含密码
//// 下面的连接字符串中可能包含明文密码和/或
//// 其他重要信息。请在查看完
//// 此连接字符串并找到所有与安全有关的问题后移除 #error。可能需要
//// 将此密码存储为其他格式或使用其他的用户身份验证。
//CString CSetQuestionInfo::GetDefaultConnect()
//{
//	return _T("DSN=eTest;APP=Microsoft\x00ae Visual Studio .NET;WSID=SKY;DATABASE=eTest;LANGUAGE=\x7b80\x4f53\x4e2d\x6587;Trusted_Connection=Yes");
//}

CString CSetQuestionInfo::GetDefaultSQL()
{
	return _T("[dbo].[QuestionInfo]");
}

void CSetQuestionInfo::DoFieldExchange(CFieldExchange* pFX)
{
	pFX->SetFieldType(CFieldExchange::outputColumn);
// RFX_Text() 和 RFX_Int() 这类宏依赖的是
// 成员变量的类型，而不是数据库字段的类型。
// ODBC 尝试自动将列值转换为所请求的类型
	RFX_Long(pFX, _T("[QuestionID]"), m_QuestionID);
	RFX_Int(pFX, _T("[SubjectID]"), m_SubjectID);
	RFX_Byte(pFX, _T("[QuestionNum]"), m_QuestionNum);
	RFX_Byte(pFX, _T("[QuestionType]"), m_QuestionType);
	RFX_Single(pFX, _T("[ReservedMark]"), m_ReservedMark);
	RFX_Byte(pFX, _T("[WorkLoad]"), m_WorkLoad);
	RFX_Byte(pFX, _T("[Gap]"), m_Gap);
	RFX_Text(pFX, _T("[Note]"), m_Note);
	RFX_Text(pFX, _T("[QuestionTitle]"), m_QuestionTitle);
	RFX_Text(pFX, _T("[SubQuestion]"), m_SubQuestion);
//	RFX_Byte(pFX, _T("[CntStep]"), m_CntStep);

}
/////////////////////////////////////////////////////////////////////////////
// CSetQuestionInfo 诊断

#ifdef _DEBUG
void CSetQuestionInfo::AssertValid() const
{
	CRecordset::AssertValid();
}

void CSetQuestionInfo::Dump(CDumpContext& dc) const
{
	CRecordset::Dump(dc);
}
#endif //_DEBUG


