// vChoicePercentage.h : CvChoicePercentage 类的实现



// CvChoicePercentage 实现

// 代码生成在 2006年2月5日, 15:32

#include "stdafx.h"
#include "vChoicePercentage.h"
IMPLEMENT_DYNAMIC(CvChoicePercentage, CRecordset)

CvChoicePercentage::CvChoicePercentage(CDatabase* pdb)
	: CRecordset(pdb)
{
	m_ID			= 0	;
	m_QuestionFormID= 0	;
	m_Choice		= 0	;
	m_percentage	= 0	;
	m_SubjectID		= 0	;
	m_SchoolID		= 0	;
	m_Name			= "";
	m_Type			= 0	;

	m_nFields		= 8			;
	m_nDefaultType	= snapshot	;
}
//#error Security Issue: The connection string may contain a password
//// 此连接字符串中可能包含密码
//// 下面的连接字符串中可能包含明文密码和/或
//// 其他重要信息。请在查看完
//// 此连接字符串并找到所有与安全有关的问题后移除 #error。可能需要
//// 将此密码存储为其他格式或使用其他的用户身份验证。
//CString CvChoicePercentage::GetDefaultConnect()
//{
//	return _T("DSN=eTest;APP=Microsoft\x00ae Visual Studio .NET;WSID=SKY;DATABASE=eTest;LANGUAGE=\x7b80\x4f53\x4e2d\x6587;Trusted_Connection=Yes");
//}

CString CvChoicePercentage::GetDefaultSQL()
{
	return _T("[dbo].[vChoiceCnt]");
}

void CvChoicePercentage::DoFieldExchange(CFieldExchange* pFX)
{
	pFX->SetFieldType(CFieldExchange::outputColumn);
// RFX_Text() 和 RFX_Int() 这类宏依赖的是
// 成员变量的类型，而不是数据库字段的类型。
// ODBC 尝试自动将列值转换为所请求的类型
	RFX_Long(pFX, _T("[ClassID]"), m_ID);
	RFX_Long(pFX, _T("[QuestionFormID]"), m_QuestionFormID);
	RFX_Byte(pFX, _T("[Choice]"), m_Choice);
	RFX_Single(pFX, _T("[percentage]"), m_percentage);
	RFX_Int(pFX, _T("[SubjectID]"), m_SubjectID);
	RFX_Long(pFX, _T("[SchoolID]"), m_SchoolID);
	RFX_Text(pFX, _T("[ClassName]"), m_Name);
	RFX_Byte(pFX, _T("[Type]"), m_Type);
}
/////////////////////////////////////////////////////////////////////////////
// CvChoicePercentage 诊断

#ifdef _DEBUG
void CvChoicePercentage::AssertValid() const
{
	CRecordset::AssertValid();
}

void CvChoicePercentage::Dump(CDumpContext& dc) const
{
	CRecordset::Dump(dc);
}
#endif //_DEBUG


