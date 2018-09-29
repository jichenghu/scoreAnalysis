// viewAVGRegionStep.h : CViewAVGRegionStep ���ʵ��



// CViewAVGRegionStep ʵ��

// ���������� 2005��9��2��, 10:01

#include "stdafx.h"
#include "viewAVGRegionStep.h"
IMPLEMENT_DYNAMIC(CViewAVGRegionStep, CRecordset)

CViewAVGRegionStep::CViewAVGRegionStep(CDatabase* pdb)
	: CRecordset(pdb)
{
	m_QuestionID = 0;
	m_Step = 0;
	m_avg = 0.0;
	m_RegionID = 0;
	m_SubjectID = 0;
	m_Type = 0;
	m_nFields = 6;
	m_nDefaultType = snapshot;
}
//#error Security Issue: The connection string may contain a password
//// �������ַ����п��ܰ�������
//// ����������ַ����п��ܰ������������/��
//// ������Ҫ��Ϣ�����ڲ鿴��
//// �������ַ������ҵ������밲ȫ�йص�������Ƴ� #error��������Ҫ
//// ��������洢Ϊ������ʽ��ʹ���������û������֤��
//CString CViewAVGRegionStep::GetDefaultConnect()
//{
//	return _T("DSN=eTest;APP=Microsoft\x00ae Visual Studio .NET;WSID=SKY;DATABASE=eTest;LANGUAGE=\x7b80\x4f53\x4e2d\x6587;Trusted_Connection=Yes");
//}

CString CViewAVGRegionStep::GetDefaultSQL()
{
	return _T("[dbo].[vAVGRegionStep]");
}

void CViewAVGRegionStep::DoFieldExchange(CFieldExchange* pFX)
{
	pFX->SetFieldType(CFieldExchange::outputColumn);
// RFX_Text() �� RFX_Int() �������������
// ��Ա���������ͣ����������ݿ��ֶε����͡�
// ODBC �����Զ�����ֵת��Ϊ�����������
	RFX_Long(pFX, _T("[QuestionID]"), m_QuestionID);
	RFX_Byte(pFX, _T("[Step]"), m_Step);
	RFX_Single(pFX, _T("[avg]"), m_avg);
	RFX_Int(pFX, _T("[RegionID]"), m_RegionID);
	RFX_Int(pFX, _T("[SubjectID]"), m_SubjectID);
	RFX_Byte(pFX, _T("[Type]"), m_Type);

}
/////////////////////////////////////////////////////////////////////////////
// CViewAVGRegionStep ���

#ifdef _DEBUG
void CViewAVGRegionStep::AssertValid() const
{
	CRecordset::AssertValid();
}

void CViewAVGRegionStep::Dump(CDumpContext& dc) const
{
	CRecordset::Dump(dc);
}
#endif //_DEBUG


