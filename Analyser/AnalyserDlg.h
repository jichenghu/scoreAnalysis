// AnalyserDlg.h : ͷ�ļ�
//

#pragma once
#include "afxwin.h"
#include "DialogResize.h"

#include "PageSingleSubject.h"
#include "PageMultiSubjects.h"
#include "PageConfig.h"

#include "afxcmn.h"

//#import  "D:\\smart\\dtspkg.dll" \
//   no_namespace rename("EOF", "EndOfFile") 

#import  "dtspkg.dll" no_namespace rename("EOF", "EndOfFile") 

class CAnalyzer	;

// CAnalyserDlg �Ի���
class CAnalyserDlg : public CDialogResize
{
// �ⲿ�ӿ�
public:
	void	AddItem ( LPITEM lpItem )	;	// ����һ�� item
	void	DestroyAllItems ( )			;	// ����ȫ���� item
	
	LPITEM	GetItemList ( )	{return m_lstItemHead;};

	int		GetRegionID ( ) {return m_pageConfig.m_iRegionID;};
	CString	GetRegionName ( ) {return m_pageConfig.m_strRegion;};
	CString	GetPathOutput ( ) {return m_strPathOutput;};

	void	SetPathOutput ( CString strPath );

	void	SetPage ( int nIndex )	;

// �ؼ�����
public:
	CAnalyzer *	m_pAnalyzer	;
   
protected:
	BOOL	m_bInitialized	;

	LPITEM	m_lstItemHead	;
	LPITEM	m_lstItemTail	;

	int		m_iCurrentItem	;	// ��ǰ�Ļpage

	CString	m_strPathOutput	;	// �����·��

// �ڲ�����
protected:
	void	ActivatePage(int nIndex);
	void	MoveChilds()			;

// ����
public:
	CAnalyserDlg(CWnd* pParent = NULL);	// ��׼���캯��

	CPageConfig			m_pageConfig;
	CPageSingleSubject	m_pageSingle;
	CPageMultiSubjects	m_pageMulti	;

// �Ի�������
	enum { IDD = IDD_ANALYSER_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��

// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
	virtual BOOL OnInitDialog();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnDestroy();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
	DECLARE_DLGRESIZE_MAP;

public:
	afx_msg void OnTcnSelchangeTab(NMHDR *pNMHDR, LRESULT *pResult);
	
	CTabCtrl		m_tabCtrl	;
};
