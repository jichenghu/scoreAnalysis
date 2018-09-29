
#include "stdafx.h"
#include "Bind.h"
#include "setSelectID.h"

//////////////////////////////////////////////////////////////
// BindSQL()
// 绑定SQL数据源到列表框或组合框
// 参数 wnd:     列表控件或者组合框控件对象，要求支持ResetContent()和AddString()函数
//      sql:     绑定的sql语句(注：该函数不进行sql语句有效性校验，请自行确保其正确性)
//               sql语句必须类似"select distinct [field] from tablename where ..."的形式
//      attach:  附加功能开关(添加首行'-')
//////////////////////////////////////////////////////////////
template <class T> bool BindSQL(T &wnd, CString sql, CMapStringToPtr* pMapID, bool bAttach)
{
	CString			str		;
	CSetSelectID	setID	;

	setID.m_pDatabase	= &theApp.m_database;

	wnd.ResetContent ( )	;

	POSITION	pos		;
	CString		key		;
	int *		piID	;

	for( pos = pMapID->GetStartPosition(); pos != NULL; )
	{
		pMapID->GetNextAssoc( pos, key, (void*&)piID );
#ifdef _DEBUG
//		afxDump << key << " : " << lpLevel << "\n";
#endif
		delete	piID;
	}
	pMapID->RemoveAll ( )	;

	try
	{
		setID.Open ( CRecordset::snapshot, sql, CRecordset::readOnly )	;

		while ( !setID.IsEOF() )
		{
			wnd.AddString ( setID.m_strName )	;

			piID	= new int	;

			*piID	= setID.m_iID	;

			pMapID->SetAt( setID.m_strName, piID );

			setID.MoveNext ( )	;
		}

		setID.Close ( )	;
	}
	catch ( CDBException * e )
	{
		AfxMessageBox ( e->m_strError )	;

		return	false;
	}

	//选择第一项
	if( wnd.GetCount() > 0 ) 
		wnd.SetCurSel( 0 )	;

	return true;
}

template <class T> bool BindSQL(T &wnd, CString sql, bool bAttach)
{
	CString			str		;
	CSetSelectID	setID	;

	wnd.ResetContent ( )	;

	setID.m_pDatabase	= &theApp.m_database;

	try
	{
		setID.Open ( CRecordset::snapshot, sql, CRecordset::readOnly )	;

		while ( !setID.IsEOF() )
		{
			wnd.AddString ( setID.m_strName )	;

			//int*	piID	= new int	;

			//*piID	= setID.m_iID	;

			//pMapID->SetAt( setID.m_strName, piID );

			setID.MoveNext ( )	;
		}

		setID.Close ( )	;
	}
	catch ( CDBException * e )
	{
		AfxMessageBox ( e->m_strError )	;

		return	false;
	}

	//选择第一项
	if( wnd.GetCount() > 0 ) 
		wnd.SetCurSel( 0 )	;

	return true;
}

