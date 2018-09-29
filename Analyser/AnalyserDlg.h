// AnalyserDlg.h : 头文件
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

// CAnalyserDlg 对话框
class CAnalyserDlg : public CDialogResize
{
// 外部接口
public:
	void	AddItem ( LPITEM lpItem )	;	// 增加一个 item
	void	DestroyAllItems ( )			;	// 消除全部的 item
	
	LPITEM	GetItemList ( )	{return m_lstItemHead;};

	int		GetRegionID ( ) {return m_pageConfig.m_iRegionID;};
	CString	GetRegionName ( ) {return m_pageConfig.m_strRegion;};
	CString	GetPathOutput ( ) {return m_strPathOutput;};

	void	SetPathOutput ( CString strPath );

	void	SetPage ( int nIndex )	;

// 关键对象
public:
	CAnalyzer *	m_pAnalyzer	;
   
protected:
	BOOL	m_bInitialized	;

	LPITEM	m_lstItemHead	;
	LPITEM	m_lstItemTail	;

	int		m_iCurrentItem	;	// 当前的活动page

	CString	m_strPathOutput	;	// 输出的路径

// 内部动作
protected:
	void	ActivatePage(int nIndex);
	void	MoveChilds()			;

// 构造
public:
	CAnalyserDlg(CWnd* pParent = NULL);	// 标准构造函数

	CPageConfig			m_pageConfig;
	CPageSingleSubject	m_pageSingle;
	CPageMultiSubjects	m_pageMulti	;

// 对话框数据
	enum { IDD = IDD_ANALYSER_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持

// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
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
