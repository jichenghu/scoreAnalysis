// Analyser.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error �ڰ������� PCH �Ĵ��ļ�֮ǰ������stdafx.h��
#endif

#include "resource.h"		// ������

#include "stdafx.h"        // ODBC

// CAnalyserApp:
// �йش����ʵ�֣������ Analyser.cpp
//

//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

//	SKLSE 2007-01-15, jicheng, �����ܷ�, begin

//-------------------------------------------------------------------------------

class CAnalyserApp : public CWinApp
{
public:
	CAnalyserApp();

	CDatabase	m_database;

// ��д
	public:
	virtual BOOL InitInstance();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CAnalyserApp theApp;
