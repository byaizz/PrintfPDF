#include "ExcelManager.h"
#include <comdef.h>
#include <afxstr.h>
#include <iostream>

using namespace std;

ExcelManager::ExcelManager(void)
:covTrue((short)TRUE)
,covFalse((short)FALSE)
,covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR)
{
}

ExcelManager::~ExcelManager(void)
{
	Close();
}

bool ExcelManager::CopyCurrentSheet()
{
	try
	{
		//复制当前sheet到当前sheet之后
		m_worksheet.Copy(covOptional,_variant_t(m_worksheet.m_lpDispatch));
		return true;
	}
	catch (CException* e)
	{
		cout<<"添加sheet失败."<<endl;
		return false;
	}
	
}

bool ExcelManager::CheckVersion()
{
	try
	{
		CString version = m_app.get_Version();
		//2010版 版本号为"14.0"; 2013版 版本号为"15.0"
		if (version == _T("14.0") || version == _T("15.0"))
		{
			return true;
		}
	}
	catch (CException* e)
	{
		cout<<"版本号检查异常."<<endl;
		return false;
	}
	return false;
}

void ExcelManager::Close()
{
	m_workbook.Close(covFalse,covOptional,covOptional);//默认不保存更改
	m_workbooks.Close();
	m_range.ReleaseDispatch();
 	m_worksheet.ReleaseDispatch();
 	m_worksheets.ReleaseDispatch();
 	m_workbook.ReleaseDispatch();
 	m_workbooks.ReleaseDispatch();
	m_app.Quit();
	m_app.ReleaseDispatch();
}

bool ExcelManager::GetCellsValue(_variant_t startCell,_variant_t endCell, VARIANT &iData)
{
	try
	{
		m_range.AttachDispatch(m_worksheet.get_Range(startCell, endCell));
		iData = m_range.get_Value2();
	}
	catch (CException* e)
	{
		cout<<"获取单元格区域数据异常."<<endl;
		return false;
	}
	return true;
}

bool ExcelManager::GetCellValue(_variant_t rowIndex, _variant_t columnIndex,VARIANT &data)
{
	try
	{
		m_range.AttachDispatch(m_worksheet.get_Cells());
		data = m_range.get_Item(rowIndex, columnIndex);
	}
	catch (CException* e)
	{
		cout<<"获取单元格数据异常."<<endl;
		return false;
	}
	return true;
}

long ExcelManager::GetSheetCount()
{
	return m_worksheets.get_Count();
}

bool ExcelManager::IsFileExist(const CString &fileName)
{
	DWORD dwAttrib = GetFileAttributes(fileName);
	return INVALID_FILE_ATTRIBUTES != dwAttrib;
}

bool ExcelManager::Open(const CString &iFileName, bool autoCreate)
{
	try
	{
		if (!m_app.CreateDispatch(_T("Excel.Application")))//创建Excel服务
		{
			cout<<"创建Excel服务失败."<<endl;
			return false;
		}
		if (!CheckVersion())
		{
			cout<<"Excel版本过低不符合要求,请检查."<<endl;
			return false;
		}
		m_app.put_Visible(FALSE);//默认不显示excel表格

		m_workbooks.AttachDispatch(m_app.get_Workbooks());
		if (IsFileExist(iFileName))
		{
			m_workbook.AttachDispatch(m_workbooks.Open(iFileName,vtMissing,vtMissing,
				vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,
				vtMissing,vtMissing,vtMissing,vtMissing,vtMissing));//获取workbook
		}
		else if (autoCreate)
		{
			m_workbook.AttachDispatch(m_workbooks.Add(_variant_t(covOptional)));//获取workbook
			SaveAs(iFileName);//创建新文件
		}
		else
		{
			cout<<"excel文件不存在."<<endl;
			return false;
		}
		m_worksheets.AttachDispatch(m_workbook.get_Sheets());//获取所有sheet
		m_worksheet.AttachDispatch(m_workbook.get_ActiveSheet());//获取当前活动sheet
	}
	catch (CException* e)
	{
		cout<<"excel文件打开异常."<<endl;
		return false;
	}
	return true;
}

bool ExcelManager::OpenFromTemplate(const CString &iFileName)
{
	try
	{
		if (!m_app.CreateDispatch(_T("Excel.Application")))//创建Excel服务
		{
			cout<<"创建Excel服务失败."<<endl;
			return false;
		}
		if (!CheckVersion())
		{
			cout<<"Excel版本过低."<<endl;
			return false;
		}
		m_app.put_Visible(FALSE);//默认不显示excel表格

		if (!IsFileExist(iFileName))
		{
			cout<<"excel文件不存在."<<endl;
			return false;
		}
		m_workbooks.AttachDispatch(m_app.get_Workbooks());
		//以路径指出的文件为模板创建excel
		m_workbook.AttachDispatch(m_workbooks.Add(_variant_t(iFileName)));//获取workbook
		m_worksheets.AttachDispatch(m_workbook.get_Sheets());//获取所有sheet
		m_worksheet.AttachDispatch(m_workbook.get_ActiveSheet());//获取当前活动sheet
	}
	catch (CException* e)
	{
		cout<<"excel文件打开异常."<<endl;
		return false;
	}
	
	return true;
}

bool ExcelManager::Save()
{
	try
	{
		m_workbook.Save();
	}
	catch (CException* e)
	{
		cout<<"文件保存异常."<<endl;
		return false;
	}
	return true;
}

bool ExcelManager::SaveAs(const CString &iFileName)
{
	try
	{
		long acessMode = 1;//默认值（不更改访问模式）
		m_workbook.SaveAs(_variant_t(iFileName),covOptional,covOptional,covOptional,covOptional,covOptional
			,acessMode,covOptional,covOptional,covOptional,covOptional,covOptional);
	}
	catch (CException* e)
	{
		cout<<"文件另存异常."<<endl;
		return false ;
	}
	return true;
}

bool ExcelManager::SaveAsPDF(const CString &iFileName)
{
	try
	{
		long type = 0;//参数0为.pdf，参数1为.xps
		m_workbook.ExportAsFixedFormat(type,_variant_t(iFileName),covOptional,covOptional,
			covOptional,covOptional,covOptional,covOptional,covOptional);
	}
	catch (CException* e)
	{
		cout<<"文件另存为pdf异常."<<endl;
		return false;
	}
	return true;
}

bool ExcelManager::SetCellsValue(_variant_t startCell,_variant_t endCell,COleSafeArray &iTwoDimArray)
{
	try
	{
		m_range.AttachDispatch(m_worksheet.get_Range(_variant_t(startCell),_variant_t(endCell)));
		if (!IsRegionEqual(m_range,iTwoDimArray))
		{
			return false;
		}
		m_range.put_Value2(iTwoDimArray);
	}
	catch (CException* e)
	{
		cout<<"写入单元格区域数据异常."<<endl;
		return false;
	}
	
	return true;
}

bool ExcelManager::SetCellValue(_variant_t rowIndex, _variant_t columnIndex,_variant_t data)
{
	try
	{
		m_range.AttachDispatch(m_worksheet.get_Cells());
		m_range.put_Item(rowIndex,columnIndex,data);
	}
	catch (CException* e)
	{
		cout<<"写入单元格数据异常."<<endl;
		return false;
	}
	return true;
}

bool ExcelManager::SwitchWorksheet(const long &sheetIndex)
{
	try
	{
		m_worksheet.AttachDispatch(m_worksheets.get_Item(_variant_t(sheetIndex)));//获取sheet
	}
	catch (CException* e)
	{
		cout<<"切换sheet异常."<<endl;
		return false;
	}
	return true;
}

void ExcelManager::Test1()
{
	//获取数据
	COleSafeArray safeArray;
	GetCellsValue(_T("$A$1"),_T("$Z$35"),safeArray);
	long firstLBound = 0;//第一维左边界
	long firstUBound = 0;//第一维右边界
	long secondLBound = 0;//第二维左边界
	long secondUBound = 0;//第二维右边界
	safeArray.GetLBound(1,&firstLBound);
	safeArray.GetUBound(1,&firstUBound);
	safeArray.GetLBound(2,&secondLBound);
	safeArray.GetUBound(2,&secondUBound);

	long index[2] = {0};
	for (int i = firstLBound; i <= firstUBound; ++i)
	{
		index[0] = i;
		for (int j = secondLBound; j <= secondUBound; ++j)
		{
			index[1] = j;
			VARIANT var;
			safeArray.GetElement(index,&var);
			//if (var.vt == VT_BSTR)
			{
				float num = 123.00000;
				var = _variant_t(num);
			}
			safeArray.PutElement(index,&var);
		}
	}

	SetCellsValue(_T("$A$1"),_T("$Z$35"),safeArray);

	CPageSetup pageSetUp;
	pageSetUp.AttachDispatch(m_worksheet.get_PageSetup());
	long xlLandscape = 2;//横向模式
	pageSetUp.put_Orientation(xlLandscape);
	//pageSetUp.put_PrintArea(_T("$I$25:$P$31,$J$35:$M$41,$O$35:$R$41"));
	//pageSetUp.put_PrintArea(_T("A15:G30"));

// 	m_range.AttachDispatch(m_worksheet.get_Cells());
// 	VARIANT var = COleVariant(_T("mill"));
// 
// 	LPDISPATCH lpDisp = m_range.Find(var,vtMissing,vtMissing,
// 		vtMissing,vtMissing,1,vtMissing,vtMissing,vtMissing);
// 	
// 	if (lpDisp != NULL)
// 	{
// 		m_range.AttachDispatch(lpDisp);
// 		CString str = m_range.get_Address(vtMissing,vtMissing,2,vtMissing,vtMissing);
// 		MessageBox(NULL,str,NULL,MB_OK);
// 	}
// 	else
// 	{
// 		MessageBox(NULL,_T("没找到"),NULL,MB_OK);
// 	}
}

bool ExcelManager::IsRegionEqual(CRange &iRange,COleSafeArray &iTwoDimArray)
{
	if ( iTwoDimArray.GetDim() != 2)//数组维数判断
	{
		return false;//数组非二维数组
	}

	//获取二维数组行数和列数
	long firstLBound = 0;//第一维左边界
	long fristUBound = 0;//第一维右边界
	long secondLBound = 0;//第二维左边界
	long secondUBound = 0;//第二维右边界
	long row1 = 0;//行数
	long column1 = 0;//列数
	iTwoDimArray.GetLBound(1,&firstLBound);
	iTwoDimArray.GetUBound(1,&fristUBound);
	iTwoDimArray.GetLBound(2,&secondLBound);
	iTwoDimArray.GetUBound(2,&secondUBound);
	row1 = fristUBound - firstLBound + 1;
	column1 = secondUBound - secondLBound + 1;

	//获取表格区域行数和列数
	CRange range;
	range.AttachDispatch(iRange.get_Rows());
	long row2 = range.get_Count();//获取行数
	range.AttachDispatch(iRange.get_Columns());
	long column2 = range.get_Count();//获取列数
	if (row1 != row2 || column1 != column2)
	{
		return false;//区域大小不一致
	}
	return true;
}

CRange * ExcelManager::GetRange()
{
	return &m_range;
}

CWorkbook * ExcelManager::GetWorkbook()
{
	return &m_workbook;
}

CWorkbooks * ExcelManager::GetWorkbooks()
{
	return &m_workbooks;
}

CWorksheet * ExcelManager::GetWorksheet()
{
	return &m_worksheet;
}

CWorksheets * ExcelManager::GetWorksheets()
{
	return &m_worksheets;
}
