// setCutSheetScoreEx.h : CSetCutSheetScoreEx ���ʵ��



// CSetCutSheetScoreEx ʵ��

// ���������� 2005��8��31��, 10:34

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
//// �������ַ����п��ܰ�������
//// ����������ַ����п��ܰ������������/��
//// ������Ҫ��Ϣ�����ڲ鿴��
//// �������ַ������ҵ������밲ȫ�йص�������Ƴ� #error��������Ҫ
//// ��������洢Ϊ������ʽ��ʹ���������û������֤��
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
// RFX_Text() �� RFX_Int() �������������
// ��Ա���������ͣ����������ݿ��ֶε����͡�
// ODBC �����Զ�����ֵת��Ϊ�����������
	RFX_Long(pFX, _T("[ID]"), m_ID);
	RFX_Long(pFX, _T("[CutSheetID]"), m_CutSheetID);
	RFX_Long(pFX, _T("[TeacherID]"), m_TeacherID);
	RFX_Byte(pFX, _T("[Step]"), m_Step);
	RFX_Byte(pFX, _T("[Score]"), m_Score);
	RFX_Byte(pFX, _T("[Flag]"), m_Flag);

}
/////////////////////////////////////////////////////////////////////////////
// CSetCutSheetScoreEx ���

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


