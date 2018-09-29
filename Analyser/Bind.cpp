
#include "stdafx.h"
#include "Bind.h"
#include "setSelectID.h"

//////////////////////////////////////////////////////////////
// BindSQL()
// ��SQL����Դ���б�����Ͽ�
// ���� wnd:     �б�ؼ�������Ͽ�ؼ�����Ҫ��֧��ResetContent()��AddString()����
//      sql:     �󶨵�sql���(ע���ú���������sql�����Ч��У�飬������ȷ������ȷ��)
//               sql����������"select distinct [field] from tablename where ..."����ʽ
//      attach:  ���ӹ��ܿ���(�������'-')
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

	//ѡ���һ��
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

	//ѡ���һ��
	if( wnd.GetCount() > 0 ) 
		wnd.SetCurSel( 0 )	;

	return true;
}

