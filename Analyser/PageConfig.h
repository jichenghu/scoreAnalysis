#pragma once

#include "DialogResize.h"
#include "afxwin.h"


// CPageConfig �Ի���
class CAnalyserDlg	;

class CPageConfig : public CDialogResize
{
	DECLARE_DYNAMIC(CPageConfig)

public:
	CComboBox		m_ctlRegion	;
	int				m_iRegionID	;
	CString			m_strRegion	;

	CAnalyserDlg *	m_pParentDlg;

//	�ڲ�����
protected:
	BOOL	ReadConfig ( CString strPath )	;
	BOOL	ReadItem ( CStdioFile * pFile,
				BYTE yItemNumber )			;
	BOOL	ReadSub ( CStdioFile * pFile,
				LPITEM lpItem )				;
	BOOL	ReadSel ( CStdioFile * pFile,
				paraStruct ** lppPara,
				CWordAction ** lppAction, 
				LPSUBITEM lpSubItem )		;	// ���붨λ��
	BOOL	ReadFunction ( CStdioFile * pFile,
				LPSUBITEM lpSubItem,
				CString strFuncName)		;	// ���뺯��
	BOOL	ReadMaxRowOfBlock (
				CStdioFile *pFile,
				LPSUBITEM lpSubItem )		;	// ����ÿ Block �е������/����Ŀ
//	BOOL	ReadBody ( CStdioFile * pFile,
//				LPSUBITEM lpSubItem )		;	// �����嶨λ��
	void	ResetConfig ( )					;	// ������������

	BOOL	GenerateMarks ( )	;	// ���ɴ���÷ֵ���ر�
	BOOL	GenerateAvg ( )		;	// ����ƽ���ַ�������ر�

public:
	CPageConfig(CWnd* pParent = NULL);   // ��׼���캯��
	virtual ~CPageConfig();

// �Ի�������
	enum { IDD = IDD_PROPPAGE_CONFIG };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

	DECLARE_MESSAGE_MAP()

public:
	afx_msg void OnCbnSelchangeComboRegion()	;

	CButton		m_chkForm1	;
	CButton		m_chkForm2	;
	CButton		m_chkForm3	;
	CButton		m_chkForm4	;
	CButton		m_chkForm5	;
	CButton		m_chkForm6	;
	CButton		m_chkForm7	;
	CButton		m_chkForm8	;

	CButton		m_chkReverse;
	BOOL		m_bReverse	;
	
	CButton		m_btnPathConfig	;
	CButton		m_btnPathResults;
	CEdit		m_edPathConfig	;
	CEdit		m_edPathResults	;

	CMapStringToPtr	m_mapRegionID	;	// ʹ��MAP(regionName->regionID)��ѡ����
	virtual BOOL	OnInitDialog()	;
	afx_msg void	OnDestroy()		;
	afx_msg void	OnBnClickedBtnPathConfig()	;
	afx_msg void	OnBnClickedBtnPathResults()	;
	virtual BOOL	Create(UINT nIDTemplate, CWnd* pParentWnd = NULL);
	
	void			OnBnClickedCheck(BYTE yItem);
	afx_msg void	OnBnClickedCheck1()	;
	afx_msg void	OnBnClickedCheck2()	;
	afx_msg void	OnBnClickedCheck3()	;
	afx_msg void	OnBnClickedCheck4()	;
	afx_msg void	OnBnClickedCheck5()	;
	afx_msg void	OnBnClickedCheck6()	;
	afx_msg void	OnBnClickedCheck7()	;
	afx_msg void	OnBnClickedCheck8()	;
	afx_msg void	OnBnClickedReverse()	;
	afx_msg void	OnBnClickedBtnStepscore()	;
	afx_msg void	OnBnClickedBtnCheckCnt()	;
	afx_msg void	OnBnClickedBtnStepscore3()	;
	afx_msg void	OnPaint();
};
