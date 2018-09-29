// setSelectID.h : CSetSelectID ���ʵ��



// CSetSelectID ʵ��

// ���������� 2005��8��31��, 20:19

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
// RFX_Text() �� RFX_Int() �������������
// ��Ա���������ͣ����������ݿ��ֶε����͡�
// ODBC �����Զ�����ֵת��Ϊ�����������
	RFX_Int(pFX, _T("[RegionID]"), m_iID);
	RFX_Text(pFX, _T("[RegionName]"), m_strName);

}
/////////////////////////////////////////////////////////////////////////////
// CSetSelectID ���

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


