
template <class T> bool BindSQL(T &wnd, CString sql, CMapStringToPtr* pMapID, bool bAttach);
template <class T> bool BindSQL(T &wnd, CString sql, bool bAttach);
template <class T> bool AddBind(T &wnd, CString sql, CMapStringToPtr* pMapID, bool bAttach);

BOOL TestDir(CString dir,CString& hint)	;
BOOL TestDirectory( CString dir )	;
