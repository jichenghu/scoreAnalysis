// vAVGClass.h : CvAVGClass 类的实现



// CvAVGClass 实现

// 代码生成在 2006年2月21日, 8:20

#include "stdafx.h"
#include "vAVGClass.h"
IMPLEMENT_DYNAMIC(CvAVGClass, CRecordset)

CvAVGClass::CvAVGClass(CDatabase* pdb)
	: CRecordset(pdb)
{
	m_ClassID = 0;
	m_FormID = 0;
	m_avgScore = 0.0;
	m_SubjectID = 0.0;
	m_SchoolID = 0;
	m_ClassName = "";
	m_Type = 0;
	m_nFields = 7;
	m_nDefaultType = snapshot;
}
//#error Security Issue: The connection string may contain a password
//// 此连接字符串中可能包含密码
//// 下面的连接字符串中可能包含明文密码和/或
//// 其他重要信息。请在查看完
//// 此连接字符串并找到所有与安全有关的问题后移除 #error。可能需要
//// 将此密码存储为其他格式或使用其他的用户身份验证。
//CString CvAVGClass::GetDefaultConnect()
//{
//	return _T("DSN=eTest;APP=Microsoft\x00ae Visual Studio .NET;WSID=NETSURF;DATABASE=eTest;Trusted_Connection=Yes");
//}

CString CvAVGClass::GetDefaultSQL()
{
	return _T("[dbo].[vAVGClass]");
}

void CvAVGClass::DoFieldExchange(CFieldExchange* pFX)
{
	pFX->SetFieldType(CFieldExchange::outputColumn);
// RFX_Text() 和 RFX_Int() 这类宏依赖的是
// 成员变量的类型，而不是数据库字段的类型。
// ODBC 尝试自动将列值转换为所请求的类型
	RFX_Long(pFX, _T("[ClassID]"), m_ClassID);
	RFX_Long(pFX, _T("[FormID]"), m_FormID);
	RFX_Single(pFX, _T("[avgScore]"), m_avgScore);
	RFX_Single(pFX, _T("[SubjectID]"), m_SubjectID);
	RFX_Int(pFX, _T("[SchoolID]"), m_SchoolID);
	RFX_Text(pFX, _T("[ClassName]"), m_ClassName);
	RFX_Byte(pFX, _T("[Type]"), m_Type);

}
/////////////////////////////////////////////////////////////////////////////
// CvAVGClass 诊断

#ifdef _DEBUG
void CvAVGClass::AssertValid() const
{
	CRecordset::AssertValid();
}

void CvAVGClass::Dump(CDumpContext& dc) const
{
	CRecordset::Dump(dc);
}
#endif //_DEBUG


