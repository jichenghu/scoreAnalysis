// setCutSheetEx.h : CSetCutSheetEx ���ʵ��



// CSetCutSheetEx ʵ��

// ���������� 2005��8��31��, 10:29

#include "stdafx.h"
#include "setCutSheetEx.h"
IMPLEMENT_DYNAMIC(CSetCutSheetEx, CRecordset)

CSetCutSheetEx::CSetCutSheetEx(CDatabase* pdb)
	: CRecordset(pdb)
{
	m_CutSheetID = 0;
	m_SubjectID = 0;
	m_RoomID = 0;
	m_SheetID = 0;
	m_QuestionID = 0;
	m_Flag = 0;
	m_nFields = 6;
	m_nDefaultType = snapshot;
}
//#error Security Issue: The connection string may contain a password
//// �������ַ����п��ܰ�������
//// ����������ַ����п��ܰ������������/��
//// ������Ҫ��Ϣ�����ڲ鿴��
//// �������ַ������ҵ������밲ȫ�йص�������Ƴ� #error��������Ҫ
//// ��������洢Ϊ������ʽ��ʹ���������û������֤��
//CString CSetCutSheetEx::GetDefaultConnect()
//{
//	return _T("DSN=eTest;UID=sa;PWD=87211237,;APP=Microsoft\x00ae Visual Studio .NET;WSID=SKY;DATABASE=eTest;LANGUAGE=\x7b80\x4f53\x4e2d\x6587");
//}

CString CSetCutSheetEx::GetDefaultSQL()
{
	return _T("[dbo].[CutSheetEx]");
}

void CSetCutSheetEx::DoFieldExchange(CFieldExchange* pFX)
{
	pFX->SetFieldType(CFieldExchange::outputColumn);
// RFX_Text() �� RFX_Int() �������������
// ��Ա���������ͣ����������ݿ��ֶε����͡�
// ODBC �����Զ�����ֵת��Ϊ�����������
	RFX_Long(pFX, _T("[CutSheetID]"), m_CutSheetID);
	RFX_Int(pFX, _T("[SubjectID]"), m_SubjectID);
	RFX_Long(pFX, _T("[RoomID]"), m_RoomID);
	RFX_Long(pFX, _T("[SheetID]"), m_SheetID);
	RFX_Long(pFX, _T("[QuestionID]"), m_QuestionID);
	RFX_Byte(pFX, _T("[Flag]"), m_Flag);

}
/////////////////////////////////////////////////////////////////////////////
// CSetCutSheetEx ���

#ifdef _DEBUG
void CSetCutSheetEx::AssertValid() const
{
	CRecordset::AssertValid();
}

void CSetCutSheetEx::Dump(CDumpContext& dc) const
{
	CRecordset::Dump(dc);
}
#endif //_DEBUG


