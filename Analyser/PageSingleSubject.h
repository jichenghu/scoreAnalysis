#pragma once

#include "DialogResize.h"

// CPageSingleSubject �Ի���
class CAnalyserDlg	;

class CPageSingleSubject : public CDialogResize
{
	DECLARE_DYNAMIC(CPageSingleSubject)

public:
	CMapStringToPtr	m_mapSubjectID	;	// ʹ��MAP(subjectName->subjectID)��ѡ���Ŀ
	CMapStringToPtr	m_mapTypeID		;	// ʹ��MAP(typeName->typeID)��ѡ���Ŀ

public:
	CComboBox		m_ctlSubject	;
	CComboBox		m_ctlType		;
	CComboBox		m_ctlSubject2	;
	//CComboBox		m_ctlSubject3	;

	int				m_iSubjectID	;
	int				m_iTypeID		;
	int				m_iSubjectID2	;
	//int				m_iSubjectID3	;
	CString			m_strSubject	;
	CString			m_strType		;
	CString			m_strSubject2	;
	//CString			m_strSubject3	;
	CString			m_strPathOutput	;

	void	SetPathOutput ( CString strPath ) {m_strPathOutput=strPath;};

	CPageSingleSubject(CWnd* pParent = NULL);   // ��׼���캯��
	virtual ~CPageSingleSubject();

	CAnalyserDlg *	m_pParentDlg;

// �Ի�������
	enum { IDD = IDD_PROPPAGE_SINGLE_SUBJECT };

public:
	void	UpdateUI ( )	;
	void	UpdateData ( )	;	// ���¿ؼ��е�����

protected:
	virtual void DoDataExchange(CDataExchange* pDX)	;    // DDX/DDV ֧��

	afx_msg void OnSize(UINT nType, int cx, int cy)	;
	afx_msg void OnCbnSelchangeComboSubject()		;
	afx_msg void OnCbnSelchangeComboType()			;
	afx_msg void OnCbnSelchangeComboSubject2()		;
	//afx_msg void OnCbnSelchangeComboSubject3()		;

	DECLARE_MESSAGE_MAP()

public:
	virtual BOOL OnInitDialog();
	afx_msg void OnDestroy();
	afx_msg void OnBnClickedBtnStepanalyse();
	virtual BOOL Create(UINT nIDTemplate, CWnd* pParentWnd = NULL);
	afx_msg void OnBnClickedBtnExportScores();
	afx_msg void OnBnClickedBtnCompositive();
	afx_msg void OnBnClickedBtnExportScores2();
};
