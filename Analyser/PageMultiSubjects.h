#pragma once


#include "DialogResize.h"

// CPageMultiSubjects �Ի���

class CPageMultiSubjects : public CDialogResize
{
	DECLARE_DYNAMIC(CPageMultiSubjects)

protected:
	void	InputSingleSubject ( CFile *pFile )	;
	void	InputCompositive ( CFile *pFile )	;

public:
	CPageMultiSubjects(CWnd* pParent = NULL);   // ��׼���캯��
	virtual ~CPageMultiSubjects();

// �Ի�������
	enum { IDD = IDD_PROPPAGE_MULTI_SUBJECTS };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedBtnInput();
	virtual BOOL OnInitDialog();
	afx_msg void OnBnClickedBtnGenerate();
};
