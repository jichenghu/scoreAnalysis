// setSelectID.h : CSetSelectID 类的实现



// CSetSelectID 实现

// 代码生成在 2005年8月31日, 20:19

#include "stdafx.h"
#include "setSelectID.h"
IMPLEMENT_DYNAMIC(CSetSelectID, CRecordset)

CSetSelectID::CSetSelectID(CDatabase* pdb)
	: CRecordset(pdb)
{
	m_iID		= 0	;
	m_strName	= "";
	m_nFields	= 2	;
	m_nDefaultType = snapshot;
}

CString CSetSelectID::GetDefaultSQL()
{
	return _T("[dbo].[Region]");
}

void CSetSelectID::DoFieldExchange(CFieldExchange* pFX)
{
	pFX->SetFieldType(CFieldExchange::outputColumn);
// RFX_Text() 和 RFX_Int() 这类宏依赖的是
// 成员变量的类型，而不是数据库字段的类型。
// ODBC 尝试自动将列值转换为所请求的类型
	RFX_Int(pFX, _T("[RegionID]"), m_iID);
	RFX_Text(pFX, _T("[RegionName]"), m_strName);

}
/////////////////////////////////////////////////////////////////////////////
// CSetSelectID 诊断

#ifdef _DEBUG
void CSetSelectID::AssertValid() const
{
	CRecordset::AssertValid();
}

void CSetSelectID::Dump(CDumpContext& dc) const
{
	CRecordset::Dump(dc);
}
#endif //_DEBUG


