// setSchoolReport.h : CSetSchoolReport 的声明

#pragma once

// 代码生成在 2007年1月18日, 14:38

class CSetSchoolReport : public CRecordset
{
public:
	CSetSchoolReport(CDatabase* pDatabase = NULL);
	DECLARE_DYNAMIC(CSetSchoolReport)

// 字段/参数数据

// 以下字符串类型(如果存在)反映数据库字段(ANSI 数据类型的 CStringA 和 Unicode
// 数据类型的 CStringW)的实际数据类型。
//  这是为防止 ODBC 驱动程序执行可能
// 不必要的转换。如果希望，可以将这些成员更改为
// CString 类型，ODBC 驱动程序将执行所有必要的转换。
// (注意: 必须使用 3.5 版或更高版本的 ODBC 驱动程序
// 以同时支持 Unicode 和这些转换)。

	long	m_ID;
	long	m_StudentID;
	float	column1 ;
	float	column2 ;
	float	column3 ;
	float	column4 ;
	float	column5 ;
	float	column6 ;
	float	column7 ;
	float	column8 ;
	float	column9 ;
	float	column10;
	float	column11;
	float	column12;
	float	column13;
	float	column14;
	float	column15;
	float	column16;

	float *	m_pColumn[16];

	void	Set_column_pointer ( );

// 重写
	// 向导生成的虚函数重写
	public:
	//virtual CString GetDefaultConnect();	// 默认连接字符串

	virtual CString GetDefaultSQL(); 	// 记录集的默认 SQL
	virtual void DoFieldExchange(CFieldExchange* pFX);	// RFX 支持

// 实现
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

};


