#if !defined(AFX_PROGRESS_H__67A1A4E1_2D09_11D5_908B_0010B568E0CE__INCLUDED_)
#define AFX_PROGRESS_H__67A1A4E1_2D09_11D5_908B_0010B568E0CE__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// Progress.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CProgress dialog
#include "GradientProgressCtrl.h"
#include "Resource.h"

class CProgress : public CDialog
{
// Construction
public:
	CProgress(CWnd* pParent = NULL);   // standard constructor

	void SetRange1(int iStart, int iEnd );
	void SetRange2(int iStart, int iEnd );
	void SetPos1(int site);
	void SetPos2(int site);
	void Step1();
	void Step2();

	void SetTitle1 ( CString strTitle )	;
	void SetTitle2 ( CString strTitle )	;

	// Dialog Data
	//{{AFX_DATA(CProgress)
	enum { IDD = IDD_PROGRESS };
	CGradientProgressCtrl	m_Progress2;
	CGradientProgressCtrl	m_Progress1;
	CEdit	m_Title2;
	CEdit	m_Title1;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CProgress)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual void PostNcDestroy();
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CProgress)
	virtual BOOL OnInitDialog();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_PROGRESS_H__67A1A4E1_2D09_11D5_908B_0010B568E0CE__INCLUDED_)
