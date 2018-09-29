// vCnt.h : CvCnt ���ʵ��



// CvCnt ʵ��

// ���������� 2005��9��28��, 16:56

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
// �������ַ����п��ܰ�������
// ����������ַ����п��ܰ������������/��
// ������Ҫ��Ϣ�����ڲ鿴��
// �������ַ������ҵ������밲ȫ�йص�������Ƴ� #error��������Ҫ
// ��������洢Ϊ������ʽ��ʹ���������û������֤��
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
// RFX_Text() �� RFX_Int() �������������
// ��Ա���������ͣ����������ݿ��ֶε����͡�
// ODBC �����Զ�����ֵת��Ϊ�����������
	RFX_Long(pFX, _T("[Cnt]"), m_Cnt);
//	RFX_Long(pFX, _T("[TeacherID]"), m_TeacherID);

}
/////////////////////////////////////////////////////////////////////////////
// CvCnt ���

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


