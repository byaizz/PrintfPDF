#pragma once

#include "afxdisp.h"
#include "CApplication.h"
#include "CWorkbooks.h"
#include "CWorksheets.h"
#include "CWorkbook.h"
#include "CWorksheet.h"
#include "CRange.h"
#include "CPageSetup.h"
#include "CFont0.h"
#include "..\RollScheduleDefinition.h"
#include <comutil.h>


class ExcelManager
{
public:
	ExcelManager(void);
	~ExcelManager(void);
	
	//************************************************************************
	// Method:		CheckVersion	检查excel版本，当前支持2010(14.0),2013(15.0)
	// Returns:		bool
	// Author:		byshi
	// Date:		2018-12-11	
	//************************************************************************
	bool CheckVersion();
	
	//************************************************************************
	// Method:		Close
	// Returns:		void
	// Author:		byshi
	// Date:		2018-12-7	
	//************************************************************************
	void Close();
	
	//************************************************************************
	// Method:		GetCellsValue	获取批量单元格数据
	// Returns:		bool
	// Parameter:	COleVariant startCell	起始单元格
	// Parameter:	COleVariant endCell		结束单元格
	// Parameter:	VARIANT & iData		返回获取到的数据值
	// Author:		byshi
	// Date:		2018-12-10	
	//************************************************************************
	bool GetCellsValue(_variant_t startCell,_variant_t endCell,
		VARIANT &iData);

	//************************************************************************
	// Method:		GetCellValue	获取单元格数据
	// Returns:		bool
	// Parameter:	COleVariant rowIndex		单元格行号,("2",或数字2)
	// Parameter:	COleVariant columnIndex		单元格列号,("C",或数字3)
	// Parameter:	VARIANT & data		返回获取到的数据值
	// Author:		byshi
	// Date:		2018-12-10	
	//************************************************************************
	bool GetCellValue(_variant_t rowIndex, _variant_t columnIndex,VARIANT &data);
	
	//************************************************************************
	// Method:		IsFileExist		判断文件或文件夹是否存在
	// Returns:		bool
	// Parameter:	const CString & fileName	文件路径名
	// Author:		byshi
	// Date:		2018-12-7	
	//************************************************************************
	bool IsFileExist(const CString &fileName);
	
	//************************************************************************
	// Method:		Open	打开文件
	// Returns:		bool
	// Parameter:	const CString & iFileName	文件路径名
	// Parameter:	bool autoCreate		打开模式，默认文件不存在不自动创建
	// Author:		byshi
	// Date:		2018-12-7	
	//************************************************************************
	bool Open(const CString &iFileName, bool autoCreate = false);
	
	//************************************************************************
	// Method:		OpenFromTemplate	根据模板创建新文件
	// Returns:		bool
	// Parameter:	const CString & iFileName	模板文件路径名	
	// Author:		byshi
	// Date:		2018-12-7	
	//************************************************************************
	bool OpenFromTemplate(const CString &iFileName);
	
	//************************************************************************
	// Method:		Save	保存文件，OpenFromTemplate创建的文件不能使用该方法保存
	// Returns:		void
	// Author:		byshi
	// Date:		2018-12-7	
	//************************************************************************
	bool Save();
	
	//************************************************************************
	// Method:		SaveAs	另存为
	// Returns:		bool
	// Parameter:	const CString & iFileName	文件路径名
	// Author:		byshi
	// Date:		2018-12-7	
	//************************************************************************
	bool SaveAs(const CString &iFileName);
	
	//************************************************************************
	// Method:		SaveAsPDF	另存为PDF文件
	// Returns:		bool
	// Parameter:	const CString & iFileName	文件路径名
	// Author:		byshi
	// Date:		2018-12-7	
	//************************************************************************
	bool SaveAsPDF(const CString &iFileName);
	
	//************************************************************************
	// Method:		SetCellsValue	批量写入多个单元格数据
	// Returns:		bool
	// Parameter:	COleVariant startCell	起始单元格(左上角单元格),例如:"B5"或"$B$5"
	// Parameter:	COleVariant endCell		起始单元格(左上角单元格),例如:"C7"或"$C$7"
	// Parameter:	COleSafeArray & iTwoDimArray	二维数组(要写入的数据)
	// Author:		byshi
	// Date:		2018-12-7	
	//************************************************************************
	bool SetCellsValue(_variant_t startCell,_variant_t endCell,
		COleSafeArray &iTwoDimArray);
	
	//************************************************************************
	// Method:		SetCellValue	写入单个单元格数据
	// Returns:		bool
	// Parameter:	COleVariant rowIndex		单元格行号,例如:"8"或数字表示
	// Parameter:	COleVariant columnIndex		单元格列号,例如:"C"或数字表示
	// Parameter:	COleVariant data			要写入的数据
	// Author:		byshi
	// Date:		2018-12-7	
	//************************************************************************
	bool SetCellValue(_variant_t rowIndex, _variant_t columnIndex,_variant_t data);
	
	//************************************************************************
	// Method:		SwitchWorksheet		切换sheet
	// Returns:		bool
	// Parameter:	const CString & sheetName	sheet名称
	// Author:		byshi
	// Date:		2018-12-7	
	//************************************************************************
	bool SwitchWorksheet(const long &sheetIndex);
	
	void Test1();

protected:	
	//************************************************************************
	// Method:		GetRange	获取CRange对象的指针
	// Returns:		CRange *
	// Author:		byshi
	// Date:		2018-12-12	
	//************************************************************************
	CRange *GetRange();
	
	//************************************************************************
	// Method:		GetWorkbook		获取CWorkbook对象的指针
	// Returns:		CWorkbook *
	// Author:		byshi
	// Date:		2018-12-12	
	//************************************************************************
	CWorkbook *GetWorkbook();
	
	//************************************************************************
	// Method:		GetWorkbooks	获取CWorkbooks对象的指针
	// Returns:		CWorkbooks *
	// Author:		byshi
	// Date:		2018-12-12	
	//************************************************************************
	CWorkbooks *GetWorkbooks();
	
	//************************************************************************
	// Method:		GetWorksheet	获取CWorksheet对象的指针
	// Returns:		CWorksheet *
	// Author:		byshi
	// Date:		2018-12-12	
	//************************************************************************
	CWorksheet *GetWorksheet();
	
	//************************************************************************
	// Method:		GetWorksheets	获取CWorksheets对象的指针
	// Returns:		CWorksheets *
	// Author:		byshi
	// Date:		2018-12-12	
	//************************************************************************
	CWorksheets *GetWorksheets();

	//************************************************************************
	// Method:		IsRegionEqual	判断单元格区域是否与二维数组区域大小相同
	// Returns:		bool
	// Parameter:	CRange & iRange		单元格区域
	// Parameter:	COleSafeArray & iTwoDimArray	二维数组
	// Author:		byshi
	// Date:		2018-12-7	
	//************************************************************************
	bool IsRegionEqual(CRange &iRange,COleSafeArray &iTwoDimArray);
	
private:
	CApplication	m_app;
	CRange			m_range;
	CWorkbook		m_workbook;
	CWorkbooks		m_workbooks;
	CWorksheet		m_worksheet;
	CWorksheets		m_worksheets;

	_variant_t		covTrue;
	_variant_t		covFalse;
	_variant_t		covOptional;
};
