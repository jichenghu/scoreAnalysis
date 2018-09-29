// setChoiceCnt.h : CSetChoiceCnt ���ʵ��



// CSetChoiceCnt ʵ��

// ���������� 2006��2��5��, 11:24

#include "stdafx.h"
#include "setChoiceCnt.h"
IMPLEMENT_DYNAMIC(CSetChoiceCnt, CRecordset)

CSetChoiceCnt::CSetChoiceCnt(CDatabase* pdb)
	: CRecordset(pdb)
{
	m_ClassID		= 0	;
	m_QuestionID	= 0	;
	m_Step			= 0	;
	m_Choice		= 0	;
	m_Cnt			= 0	;
	m_CntStudent	= 0	;

	m_nFields		= 6			;
	m_nDefaultType	= snapshot	;
}
//#error Security Issue: The connection string may contain a password
//// �������ַ����п��ܰ�������
//// ����������ַ����п��ܰ������������/��
//// ������Ҫ��Ϣ�����ڲ鿴��
//// �������ַ������ҵ������밲ȫ�йص�������Ƴ� #error��������Ҫ
//// ��������洢Ϊ������ʽ��ʹ���������û������֤��
//CString CSetChoiceCnt::GetDefaultConnect()
//{
//	return _T("DSN=eTest;APP=Microsoft\x00ae Visual Studio .NET;WSID=SKY;DATABASE=eTest;LANGUAGE=\x7b80\x4f53\x4e2d\x6587;Trusted_Connection=Yes");
//}

CString CSetChoiceCnt::GetDefaultSQL()
{
	return _T("[dbo].[ChoiceCnt]");
}

void CSetChoiceCnt::DoFieldExchange(CFieldExchange* pFX)
{
	pFX->SetFieldType(CFieldExchange::outputColumn);
// RFX_Text() �� RFX_Int() �������������
// ��Ա���������ͣ����������ݿ��ֶε����͡�
// ODBC �����Զ�����ֵת��Ϊ�����������
	RFX_Long(pFX, _T("[ClassID]"), m_ClassID);
	RFX_Long(pFX, _T("[QuestionID]"), m_QuestionID);
	RFX_Byte(pFX, _T("[Step]"), m_Step);
	RFX_Byte(pFX, _T("[Choice]"), m_Choice);
	RFX_Long(pFX, _T("[Cnt]"), m_Cnt);
	RFX_Double(pFX, _T("[CntStudent]"), m_CntStudent);

}
/////////////////////////////////////////////////////////////////////////////
// CSetChoiceCnt ���

#ifdef _DEBUG
void CSetChoiceCnt::AssertValid() const
{
	CRecordset::AssertValid();
}

void CSetChoiceCnt::Dump(CDumpContext& dc) const
{
	CRecordset::Dump(dc);
}
#endif //_DEBUG


