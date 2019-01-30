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
	// Method:		CheckVersion	���excel�汾����ǰ֧��2010(14.0),2013(15.0)
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
	// Method:		GetCellsValue	��ȡ������Ԫ������
	// Returns:		bool
	// Parameter:	COleVariant startCell	��ʼ��Ԫ��
	// Parameter:	COleVariant endCell		������Ԫ��
	// Parameter:	VARIANT & iData		���ػ�ȡ��������ֵ
	// Author:		byshi
	// Date:		2018-12-10	
	//************************************************************************
	bool GetCellsValue(_variant_t startCell,_variant_t endCell,
		VARIANT &iData);

	//************************************************************************
	// Method:		GetCellValue	��ȡ��Ԫ������
	// Returns:		bool
	// Parameter:	COleVariant rowIndex		��Ԫ���к�,("2",������2)
	// Parameter:	COleVariant columnIndex		��Ԫ���к�,("C",������3)
	// Parameter:	VARIANT & data		���ػ�ȡ��������ֵ
	// Author:		byshi
	// Date:		2018-12-10	
	//************************************************************************
	bool GetCellValue(_variant_t rowIndex, _variant_t columnIndex,VARIANT &data);
	
	//************************************************************************
	// Method:		IsFileExist		�ж��ļ����ļ����Ƿ����
	// Returns:		bool
	// Parameter:	const CString & fileName	�ļ�·����
	// Author:		byshi
	// Date:		2018-12-7	
	//************************************************************************
	bool IsFileExist(const CString &fileName);
	
	//************************************************************************
	// Method:		Open	���ļ�
	// Returns:		bool
	// Parameter:	const CString & iFileName	�ļ�·����
	// Parameter:	bool autoCreate		��ģʽ��Ĭ���ļ������ڲ��Զ�����
	// Author:		byshi
	// Date:		2018-12-7	
	//************************************************************************
	bool Open(const CString &iFileName, bool autoCreate = false);
	
	//************************************************************************
	// Method:		OpenFromTemplate	����ģ�崴�����ļ�
	// Returns:		bool
	// Parameter:	const CString & iFileName	ģ���ļ�·����	
	// Author:		byshi
	// Date:		2018-12-7	
	//************************************************************************
	bool OpenFromTemplate(const CString &iFileName);
	
	//************************************************************************
	// Method:		Save	�����ļ���OpenFromTemplate�������ļ�����ʹ�ø÷�������
	// Returns:		void
	// Author:		byshi
	// Date:		2018-12-7	
	//************************************************************************
	bool Save();
	
	//************************************************************************
	// Method:		SaveAs	���Ϊ
	// Returns:		bool
	// Parameter:	const CString & iFileName	�ļ�·����
	// Author:		byshi
	// Date:		2018-12-7	
	//************************************************************************
	bool SaveAs(const CString &iFileName);
	
	//************************************************************************
	// Method:		SaveAsPDF	���ΪPDF�ļ�
	// Returns:		bool
	// Parameter:	const CString & iFileName	�ļ�·����
	// Author:		byshi
	// Date:		2018-12-7	
	//************************************************************************
	bool SaveAsPDF(const CString &iFileName);
	
	//************************************************************************
	// Method:		SetCellsValue	����д������Ԫ������
	// Returns:		bool
	// Parameter:	COleVariant startCell	��ʼ��Ԫ��(���Ͻǵ�Ԫ��),����:"B5"��"$B$5"
	// Parameter:	COleVariant endCell		��ʼ��Ԫ��(���Ͻǵ�Ԫ��),����:"C7"��"$C$7"
	// Parameter:	COleSafeArray & iTwoDimArray	��ά����(Ҫд�������)
	// Author:		byshi
	// Date:		2018-12-7	
	//************************************************************************
	bool SetCellsValue(_variant_t startCell,_variant_t endCell,
		COleSafeArray &iTwoDimArray);
	
	//************************************************************************
	// Method:		SetCellValue	д�뵥����Ԫ������
	// Returns:		bool
	// Parameter:	COleVariant rowIndex		��Ԫ���к�,����:"8"�����ֱ�ʾ
	// Parameter:	COleVariant columnIndex		��Ԫ���к�,����:"C"�����ֱ�ʾ
	// Parameter:	COleVariant data			Ҫд�������
	// Author:		byshi
	// Date:		2018-12-7	
	//************************************************************************
	bool SetCellValue(_variant_t rowIndex, _variant_t columnIndex,_variant_t data);
	
	//************************************************************************
	// Method:		SwitchWorksheet		�л�sheet
	// Returns:		bool
	// Parameter:	const CString & sheetName	sheet����
	// Author:		byshi
	// Date:		2018-12-7	
	//************************************************************************
	bool SwitchWorksheet(const long &sheetIndex);
	
	void Test1();

protected:	
	//************************************************************************
	// Method:		GetRange	��ȡCRange�����ָ��
	// Returns:		CRange *
	// Author:		byshi
	// Date:		2018-12-12	
	//************************************************************************
	CRange *GetRange();
	
	//************************************************************************
	// Method:		GetWorkbook		��ȡCWorkbook�����ָ��
	// Returns:		CWorkbook *
	// Author:		byshi
	// Date:		2018-12-12	
	//************************************************************************
	CWorkbook *GetWorkbook();
	
	//************************************************************************
	// Method:		GetWorkbooks	��ȡCWorkbooks�����ָ��
	// Returns:		CWorkbooks *
	// Author:		byshi
	// Date:		2018-12-12	
	//************************************************************************
	CWorkbooks *GetWorkbooks();
	
	//************************************************************************
	// Method:		GetWorksheet	��ȡCWorksheet�����ָ��
	// Returns:		CWorksheet *
	// Author:		byshi
	// Date:		2018-12-12	
	//************************************************************************
	CWorksheet *GetWorksheet();
	
	//************************************************************************
	// Method:		GetWorksheets	��ȡCWorksheets�����ָ��
	// Returns:		CWorksheets *
	// Author:		byshi
	// Date:		2018-12-12	
	//************************************************************************
	CWorksheets *GetWorksheets();

	//************************************************************************
	// Method:		IsRegionEqual	�жϵ�Ԫ�������Ƿ����ά���������С��ͬ
	// Returns:		bool
	// Parameter:	CRange & iRange		��Ԫ������
	// Parameter:	COleSafeArray & iTwoDimArray	��ά����
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
