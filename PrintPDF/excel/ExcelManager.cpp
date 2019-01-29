#include "StdAfx.h"
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

bool ExcelManager::CreateTwoDimSafeArray(COleSafeArray &safeArray, DWORD dimElements1, 
										 DWORD dimElements2, long iLBound1, long iLBound2)
{
	try
	{
		const int dim = 2;
		SAFEARRAYBOUND saBound[dim];
		saBound[0].cElements = dimElements1;
		saBound[0].lLbound = iLBound1;
		saBound[1].cElements = dimElements2;
		saBound[1].lLbound = iLBound2;

		safeArray.Create(VT_VARIANT, dim, saBound);
	}
	catch (CException* e)
	{
		cout<<"创建二维数组异常."<<endl;
		return false;//创建二维数组失败
	}
	return true;
}

bool ExcelManager::GetCellsValue(COleVariant startCell,COleVariant endCell, VARIANT &iData)
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

bool ExcelManager::GetCellValue(COleVariant rowIndex, COleVariant columnIndex,VARIANT &data)
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
			m_workbook.AttachDispatch(m_workbooks.Add(COleVariant(covOptional)));//获取workbook
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
		m_workbook.AttachDispatch(m_workbooks.Add(COleVariant(iFileName)));//获取workbook
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
		m_workbook.SaveAs(COleVariant(iFileName),covOptional,covOptional,covOptional,covOptional,covOptional
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
		m_workbook.ExportAsFixedFormat(type,COleVariant(iFileName),covOptional,covOptional,
			covOptional,covOptional,covOptional,covOptional,covOptional);
	}
	catch (CException* e)
	{
		cout<<"文件另存为pdf异常."<<endl;
		return false;
	}
	return true;
}

bool ExcelManager::SetCellsValue(COleVariant startCell,COleVariant endCell,COleSafeArray &iTwoDimArray)
{
	try
	{
		m_range.AttachDispatch(m_worksheet.get_Range(COleVariant(startCell),COleVariant(endCell)));
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

bool ExcelManager::SetCellValue(COleVariant rowIndex, COleVariant columnIndex,COleVariant data)
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

bool ExcelManager::SwitchWorksheet(const CString &sheetName)
{
	try
	{
		m_worksheet.AttachDispatch(m_worksheets.get_Item(COleVariant(sheetName)));//获取sheet
	}
	catch (CException* e)
	{
		cout<<"切换sheet异常."<<endl;
		return false;
	}
	return true;
}

void ExcelManager::SetRollingData(const PDF::ROLLSchedule &rollingData)
{
	//获取数据并根据索引写入数据
	GetCellsValue(_T("$A$1"),_T("$Z$35"),safeArray);
	SetTitle();//设置标题
	SetPDIData(rollingData.pdiData);//设置PDI信息
	SetRollSetup(rollingData.rollSetup);//设置轧机设定信息
	SetMillDataRM(rollingData.millDataRM);//设置粗轧轧机轧辊信息
	SetMillDataFM(rollingData.millDataFM);//设置精轧轧机轧辊信息
	SetPassData(rollingData.passData);//设置道次信息
	SetCellsValue(_T("$A$1"),_T("$Z$35"),safeArray);

// 	CPageSetup pageSetUp;
// 	pageSetUp.AttachDispatch(m_worksheet.get_PageSetup());
// 	long xlLandscape = 2;//横向模式
// 	pageSetUp.put_Orientation(xlLandscape);
// 	pageSetUp.put_PrintArea(_T("$I$25:$P$31,$J$35:$M$41,$O$35:$R$41"));
// 	pageSetUp.put_PrintArea(_T("A15:G30"));
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
				var = COleVariant(num);
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

bool ExcelManager::Find(const CString &text, long (&index)[2])
{
	long firstLBound = 0;//第一维左边界
	long firstUBound = 0;//第一维右边界
	long secondLBound = 0;//第二维左边界
	long secondUBound = 0;//第二维右边界
	safeArray.GetLBound(1,&firstLBound);
	safeArray.GetUBound(1,&firstUBound);
	safeArray.GetLBound(2,&secondLBound);
	safeArray.GetUBound(2,&secondUBound);

	VARIANT var;
	for (int i = firstLBound; i <= firstUBound; ++i)
	{
		index[0] = i;
		for (int j = secondLBound; j <= secondUBound; ++j)
		{
			index[1] = j;	
			safeArray.GetElement(index,&var);
			if (var.vt == VT_BSTR)
			{
				CString str(var.bstrVal);
				if (str == text)
				{
					return true;
				}
			}
		}
	}
	return false;
}

bool ExcelManager::FindByColumn(const CString &text, const long column, long &row)
{
	long firstLBound = 0;//第一维左边界
	long firstUBound = 0;//第一维右边界
	long secondLBound = 0;//第二维左边界
	long secondUBound = 0;//第二维右边界
	safeArray.GetLBound(1,&firstLBound);
	safeArray.GetUBound(1,&firstUBound);
	safeArray.GetLBound(2,&secondLBound);
	safeArray.GetUBound(2,&secondUBound);

	if (column < secondLBound || column > secondUBound)
	{
		return false;
	}

	long index[2] = {0};
	index[1] = column;
	VARIANT var;
	for (int i = firstLBound; i <= firstUBound; ++i)
	{
		index[0] = i;
		safeArray.GetElement(index,&var);
		if (var.vt == VT_BSTR)
		{
			CString str(var.bstrVal);
			if (str == text)
			{
				row = i;
				return true;
			}
		}
	}
	return false;
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
void ExcelManager::SetTitle()
{

}

void ExcelManager::SetPDIData(const PDF::PDIData &pdiData)
{
	long index[2] = {0};
	VARIANT var;
	CString tempStr;
	if (Find(_T("钢种"),index))
	{
		index[1] += 2;
		tempStr.Format(_T("%s"),pdiData.steelGrade);
		var = COleVariant(tempStr);
		safeArray.PutElement(index,&var);
		index[1] += 1;
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("加热炉号"),index))
	{
		index[1] += 3;
		tempStr.Format(_T("%d-%d"),pdiData.furnaceNo,pdiData.furnaceCol);
		var = COleVariant(tempStr);
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("终轧温度 [oC]"),index))
	{
		index[1] += 2;
		tempStr.Format(_T("%d"),pdiData.finalTemp);
		var = COleVariant(tempStr);
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("板坯宽度 [mm]"),index))
	{
		index[1] += 2;
		tempStr.Format(_T("%ld"),pdiData.slabWidthL3);
		var = COleVariant(tempStr);
		safeArray.PutElement(index,&var);
		index[1] += 1;
		tempStr.Format(_T("%ld"),pdiData.slabWidthAct);
		var = COleVariant(tempStr);
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("板坯厚度 [mm]"),index))
	{
		index[1] += 2;
		tempStr.Format(_T("%ld"),pdiData.slabThickL3);
		var = COleVariant(tempStr);
		safeArray.PutElement(index,&var);
		index[1] += 1;
		tempStr.Format(_T("%ld"),pdiData.slabThickAct);
		var = COleVariant(tempStr);
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("钢板长度[mm]"),index))
	{
		index[1] += 2;
		tempStr.Format(_T("%ld"),pdiData.slabLengthL3);
		var = COleVariant(tempStr);
		safeArray.PutElement(index,&var);
		index[1] += 1;
		tempStr.Format(_T("%ld"),pdiData.slabLengthAct);
		var = COleVariant(tempStr);
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("板坯重量"),index))
	{
		index[1] += 2;
		tempStr.Format(_T("%ld"),pdiData.slabWeightL3);
		var = COleVariant(tempStr);
		safeArray.PutElement(index,&var);
		index[1] += 1;
		tempStr.Format(_T("%ld"),pdiData.slabWeightAct);
		var = COleVariant(tempStr);
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("厚度 [mm]"),index))
	{
		index[1] += 1;
		tempStr.Format(_T("%ld"),pdiData.plateThickSet);
		var = COleVariant(tempStr);
		safeArray.PutElement(index,&var);
		index[1] += 1;
		tempStr.Format(_T("%ld"),pdiData.plateThickAct);
		var = COleVariant(tempStr);
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("宽度 [mm]"),index))
	{
		index[1] += 1;
		tempStr.Format(_T("%ld"),pdiData.plateWidthSet);
		var = COleVariant(tempStr);
		safeArray.PutElement(index,&var);
		index[1] += 1;
		tempStr.Format(_T("%ld"),pdiData.plateWidthAct);
		var = COleVariant(tempStr);
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("长度 [mm]"),index))
	{
		index[1] += 1;
		tempStr.Format(_T("%ld"),pdiData.plateLengthSet);
		var = COleVariant(tempStr);
		safeArray.PutElement(index,&var);
		index[1] += 1;
		tempStr.Format(_T("%ld"),pdiData.plateLengthAct);
		var = COleVariant(tempStr);
		safeArray.PutElement(index,&var);
	}
}

void ExcelManager::SetRollSetup(const PDF::RollSetup &rollSetup)
{
	long index[2] = {0};
	VARIANT var;
	CString tempStr;
	if(Find(_T("轧制模式"),index))
	{
		index[1] += 1;
		tempStr.Format(_T("%d"),rollSetup.ctrlMode);
		var = COleVariant(tempStr);
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("总道次数"),index))
	{
		index[1] += 1;
		tempStr.Format(_T("%d"),rollSetup.totalPassNum);
		var = COleVariant(tempStr);
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("转钢模式"),index))
	{
		index[1] += 1;
		tempStr.Format(_T("%d"),rollSetup.turnMode);
		var = COleVariant(tempStr);
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("转钢道次"),index))
	{
		index[1] += 1;
		tempStr.Format(_T("%d"),rollSetup.turnPassNum);
		var = COleVariant(tempStr);
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("转钢厚度[mm]"),index))
	{
		index[1] += 1;
		tempStr.Format(_T("%ld"),rollSetup.turnThick);
		var = COleVariant(tempStr);
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("控轧厚度[mm]"),index))
	{
		index[1] += 1;
		tempStr.Format(_T("%ld"),rollSetup.controlThick);
		var = COleVariant(tempStr);
		safeArray.PutElement(index,&var);
	}
}

void ExcelManager::SetMillDataRM(const PDF::MillData &millDataRM)
{
	long index[2] = {0};	
	if (Find(_T("粗轧机轧辊数据"), index))
	{
		long row = 0;
		long column = index[1];
		if(FindByColumn(_T("上支撑辊"),column,row))
		{
			index[0] = row;
			SetRollData(millDataRM.upBackupRoll,index);
		}
		if(FindByColumn(_T("上工作辊"),column,row))
		{
			index[0] = row;
			SetRollData(millDataRM.upWorkRoll,index);
		}
		if(FindByColumn(_T("下工作辊"),column,row))
		{
			index[0] = row;
			SetRollData(millDataRM.downWorkRoll,index);
		}
		if(FindByColumn(_T("下支撑辊"),column,row))
		{
			index[0] = row;
			SetRollData(millDataRM.downBackupRoll,index);
		}
	}
}

void ExcelManager::SetMillDataFM(const PDF::MillData &millDataFM)
{
	long index[2] = {0};	
	if (Find(_T("精轧机轧辊数据"), index))
	{
		long row = 0;
		long column = index[1];
		if(FindByColumn(_T("上支撑辊"),column,row))
		{
			index[0] = row;
			SetRollData(millDataFM.upBackupRoll,index);
		}
		if(FindByColumn(_T("上工作辊"),column,row))
		{
			index[0] = row;
			SetRollData(millDataFM.upWorkRoll,index);
		}
		if(FindByColumn(_T("下工作辊"),column,row))
		{
			index[0] = row;
			SetRollData(millDataFM.downWorkRoll,index);
		}
		if(FindByColumn(_T("下支撑辊"),column,row))
		{
			index[0] = row;
			SetRollData(millDataFM.downBackupRoll,index);
		}
	}
}

void ExcelManager::SetRollData(const PDF::RollData &rollData, const long (&startIndex)[2])
{
	VARIANT var;
	CString tempStr;
	long index[2] = {0};
	index[0] = startIndex[0];
	index[1] = startIndex[1];

	index[1] += 1;
	tempStr.Format(_T("%s"),rollData.RollID);
	var = COleVariant(tempStr);
	safeArray.PutElement(index,&var);
	index[1] += 1;
	tempStr.Format(_T("%s"),rollData.RollMaterial);
	var = COleVariant(tempStr);
	safeArray.PutElement(index,&var);
	index[1] += 1;
	tempStr.Format(_T("%ld"),rollData.RollDiameter);
	var = COleVariant(tempStr);
	safeArray.PutElement(index,&var);
	index[1] += 1;
	tempStr.Format(_T("%s"),rollData.RollStartTime);
	var = COleVariant(tempStr);
	safeArray.PutElement(index,&var);
}

void ExcelManager::SetPassData(const PDF::PassData (&passData)[PASS_MAX])
{
	long index[2] = {0};
	VARIANT var;
	CString tempStr;

	if(Find(_T("轧机（RM/FM)"),index))
	{
		long row = index[0];
		long column = index[1];
		for (int i = 0; i < PASS_MAX; ++i)
		{
			if (passData[i].passNo <= 0 || passData[i].passNo >= PASS_MAX)
			{
			//	break;
			}
			index[0] = row+i+2;
			index[1] = column+1;
			tempStr.Format(_T("%d"),passData[i].passNo);
			var = COleVariant(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+2;
			tempStr.Format(_T("%ld"),passData[i].passGapSet);
			var = COleVariant(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+3;
			tempStr.Format(_T("%ld"),passData[i].passGapAct);
			var = COleVariant(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+4;
			tempStr.Format(_T("%ld"),passData[i].passThickSet);
			var = COleVariant(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+5;
			tempStr.Format(_T("%ld"),passData[i].passThickAct);
			var = COleVariant(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+6;
			tempStr.Format(_T("%ld"),passData[i].passWidthSet);
			var = COleVariant(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+7;
			tempStr.Format(_T("%ld"),passData[i].passWidthAct);
			var = COleVariant(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+8;
			tempStr.Format(_T("%ld"),passData[i].passForceSet);
			var = COleVariant(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+9;
			tempStr.Format(_T("%ld"),passData[i].passForceAct);
			var = COleVariant(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+10;
			tempStr.Format(_T("%ld"),passData[i].passTorqueSet);
			var = COleVariant(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+11;
			tempStr.Format(_T("%ld"),passData[i].passTorqueAct);
			var = COleVariant(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+12;
			tempStr.Format(_T("%ld"),passData[i].passBendForceSet);
			var = COleVariant(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+13;
			tempStr.Format(_T("%ld"),passData[i].passBendForceAct);
			var = COleVariant(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+14;
			tempStr.Format(_T("%ld"),passData[i].passShiftSet);
			var = COleVariant(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+15;
			tempStr.Format(_T("%ld"),passData[i].passShiftAct);
			var = COleVariant(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+16;
			tempStr.Format(_T("%ld"),passData[i].passVelocity);
			var = COleVariant(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+17;
			tempStr.Format(_T("%ld"),passData[i].passEntryTemp);
			var = COleVariant(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+18;
			tempStr.Format(_T("%ld"),passData[i].passExitTemp);
			var = COleVariant(tempStr);
			safeArray.PutElement(index,&var);
		}
	}
}
