// vChoicePercentage.h : CvChoicePercentage ���ʵ��



// CvChoicePercentage ʵ��

// ���������� 2006��2��5��, 15:32

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
//// �������ַ����п��ܰ�������
//// ����������ַ����п��ܰ������������/��
//// ������Ҫ��Ϣ�����ڲ鿴��
//// �������ַ������ҵ������밲ȫ�йص�������Ƴ� #error��������Ҫ
//// ��������洢Ϊ������ʽ��ʹ���������û������֤��
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
// RFX_Text() �� RFX_Int() �������������
// ��Ա���������ͣ����������ݿ��ֶε����͡�
// ODBC �����Զ�����ֵת��Ϊ�����������
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
// CvChoicePercentage ���

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


