#pragma once


#include "DialogResize.h"

// CPageMultiSubjects 对话框

class CPageMultiSubjects : public CDialogResize
{
	DECLARE_DYNAMIC(CPageMultiSubjects)

protected:
	void	InputSingleSubject ( CFile *pFile )	;
	void	InputCompositive ( CFile *pFile )	;

public:
	CPageMultiSubjects(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~CPageMultiSubjects();

// 对话框数据
	enum { IDD = IDD_PROPPAGE_MULTI_SUBJECTS };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedBtnInput();
	virtual BOOL OnInitDialog();
	afx_msg void OnBnClickedBtnGenerate();
};
