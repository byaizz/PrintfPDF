#include "PDFCreater.h"
#include <iostream>

using namespace std;

PDFCreater::PDFCreater(void)
:firstLBound(0)
,firstUBound(0)
,secondLBound(0)
,secondUBound(0)
{
}

PDFCreater::~PDFCreater(void)
{
}

void PDFCreater::Init()
{
	int nRetGUI = m_comm.InitComm();

	//循环初始化通信
	while (false == nRetGUI)
	{
		Sleep(1000);
		nRetGUI = m_comm.InitComm();
	}

	return;
}

bool PDFCreater::CreateTwoDimSafeArray(COleSafeArray &safeArray, DWORD dimElements1, 
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

void PDFCreater::Run()
{
	m_comm.RecvData();
 	if (m_comm.m_isNewRollData == true)
 	{
		//在此处将数据写入到excel，并输出为pdf文件
		if (m_excel.OpenFromTemplate(_T("E:\\pdf数据表_wq.xlsx")))
		{
			SetRollingData();
			m_excel.SaveAsPDF(_T("E:\\test_by.pdf"));
		}
		m_excel.Close();
		m_comm.m_isNewRollData = false;
 	}
}

void PDFCreater::SetRollingData()
{
	const PDF::ROLLSchedule &rollingData = m_comm.m_newRollData;
	//获取数据并根据索引写入数据
	m_excel.GetCellsValue(_T("$A$1"),_T("$Z$35"),safeArray);//读取数据
	UpdateSafeArrayBound();//更新安全数组的边界信息
	SetTitle();//设置标题
	SetPDIData();//设置PDI信息
	SetRollSetup();//设置轧机设定信息
	SetMillDataRM();//设置粗轧轧机轧辊信息
	SetMillDataFM();//设置精轧轧机轧辊信息
	m_excel.SetCellsValue(_T("$A$1"),_T("$Z$35"),safeArray);//写入数据

	//当前每个sheet可写入20道次数据信息，可根据需要修改
	int sheetNum = rollingData.rollSetup.totalPassNum / 20;
	for (int i =0; i <rollingData.rollSetup.totalPassNum; ++i)
	{
// 		GetCellsValue(_T("$A$1"),_T("$Z$35"),safeArray)
// 			if (passData[i].passNo <= 0 || passData[i].passNo >= PASS_MAX)
// 			{
// 				//	break;
// 			}
// 			index[0] = row+i+2;
// 			index[1] = column+1;
// 			tempStr.Format(_T("%d"),passData[i].passNo);
// 			var = COleVariant(tempStr);
// 			safeArray.PutElement(index,&var);
// 			index[1] = column+2;
// 			tempStr.Format(_T("%ld"),passData[i].passGapSet);
// 			var = COleVariant(tempStr);
// 			safeArray.PutElement(index,&var);
// 			index[1] = column+3;
// 			tempStr.Format(_T("%ld"),passData[i].passGapAct);
// 			var = COleVariant(tempStr);
// 			safeArray.PutElement(index,&var);
// 			index[1] = column+4;
// 			tempStr.Format(_T("%ld"),passData[i].passThickSet);
// 			var = COleVariant(tempStr);
// 			safeArray.PutElement(index,&var);
// 			index[1] = column+5;
// 			tempStr.Format(_T("%ld"),passData[i].passThickAct);
// 			var = COleVariant(tempStr);
// 			safeArray.PutElement(index,&var);
// 			index[1] = column+6;
// 			tempStr.Format(_T("%ld"),passData[i].passWidthSet);
// 			var = COleVariant(tempStr);
// 			safeArray.PutElement(index,&var);
// 			index[1] = column+7;
// 			tempStr.Format(_T("%ld"),passData[i].passWidthAct);
// 			var = COleVariant(tempStr);
// 			safeArray.PutElement(index,&var);
// 			index[1] = column+8;
// 			tempStr.Format(_T("%ld"),passData[i].passForceSet);
// 			var = COleVariant(tempStr);
// 			safeArray.PutElement(index,&var);
// 			index[1] = column+9;
// 			tempStr.Format(_T("%ld"),passData[i].passForceAct);
// 			var = COleVariant(tempStr);
// 			safeArray.PutElement(index,&var);
// 			index[1] = column+10;
// 			tempStr.Format(_T("%ld"),passData[i].passTorqueSet);
// 			var = COleVariant(tempStr);
// 			safeArray.PutElement(index,&var);
// 			index[1] = column+11;
// 			tempStr.Format(_T("%ld"),passData[i].passTorqueAct);
// 			var = COleVariant(tempStr);
// 			safeArray.PutElement(index,&var);
// 			index[1] = column+12;
// 			tempStr.Format(_T("%ld"),passData[i].passBendForceSet);
// 			var = COleVariant(tempStr);
// 			safeArray.PutElement(index,&var);
// 			index[1] = column+13;
// 			tempStr.Format(_T("%ld"),passData[i].passBendForceAct);
// 			var = COleVariant(tempStr);
// 			safeArray.PutElement(index,&var);
// 			index[1] = column+14;
// 			tempStr.Format(_T("%ld"),passData[i].passShiftSet);
// 			var = COleVariant(tempStr);
// 			safeArray.PutElement(index,&var);
// 			index[1] = column+15;
// 			tempStr.Format(_T("%ld"),passData[i].passShiftAct);
// 			var = COleVariant(tempStr);
// 			safeArray.PutElement(index,&var);
// 			index[1] = column+16;
// 			tempStr.Format(_T("%ld"),passData[i].passVelocity);
// 			var = COleVariant(tempStr);
// 			safeArray.PutElement(index,&var);
// 			index[1] = column+17;
// 			tempStr.Format(_T("%ld"),passData[i].passEntryTemp);
// 			var = COleVariant(tempStr);
// 			safeArray.PutElement(index,&var);
// 			index[1] = column+18;
// 			tempStr.Format(_T("%ld"),passData[i].passExitTemp);
// 			var = COleVariant(tempStr);
// 			safeArray.PutElement(index,&var);
	}
	SetPassData(rollingData.passData);//设置道次信息

	// 	CPageSetup pageSetUp;
	// 	pageSetUp.AttachDispatch(m_worksheet.get_PageSetup());
	// 	long xlLandscape = 2;//横向模式
	// 	pageSetUp.put_Orientation(xlLandscape);
	// 	pageSetUp.put_PrintArea(_T("$I$25:$P$31,$J$35:$M$41,$O$35:$R$41"));
	// 	pageSetUp.put_PrintArea(_T("A15:G30"));
}

void PDFCreater::UpdateSafeArrayBound()
{
	safeArray.GetLBound(1,&firstLBound);
	safeArray.GetUBound(1,&firstUBound);
	safeArray.GetLBound(2,&secondLBound);
	safeArray.GetUBound(2,&secondUBound);
}

bool PDFCreater::Find(const CString &text, long (&index)[2])
{
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

bool PDFCreater::FindByColumn(const CString &text, const long column, long &row)
{
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

void PDFCreater::SetTitle()
{

}

void PDFCreater::SetPDIData()
{
	const PDF::PDIData &pdiData = m_comm.m_newRollData.pdiData;
	long index[2] = {0};
	VARIANT var;
	CString tempStr;
	if (Find(_T("钢种"),index))
	{
		index[1] += 2;
		tempStr.Format(_T("%s"),pdiData.steelGrade);
		var = _variant_t(tempStr);
		safeArray.PutElement(index,&var);
		index[1] += 1;
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("加热炉号"),index))
	{
		index[1] += 3;
		tempStr.Format(_T("%d-%d"),pdiData.furnaceNo,pdiData.furnaceCol);
		var = _variant_t(tempStr);
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("终轧温度 [oC]"),index))
	{
		index[1] += 2;
		tempStr.Format(_T("%d"),pdiData.finalTemp);
		var = _variant_t(tempStr);
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("板坯宽度 [mm]"),index))
	{
		index[1] += 2;
		tempStr.Format(_T("%ld"),pdiData.slabWidthL3);
		var = _variant_t(tempStr);
		safeArray.PutElement(index,&var);
		index[1] += 1;
		tempStr.Format(_T("%ld"),pdiData.slabWidthAct);
		var = _variant_t(tempStr);
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("板坯厚度 [mm]"),index))
	{
		index[1] += 2;
		tempStr.Format(_T("%ld"),pdiData.slabThickL3);
		var = _variant_t(tempStr);
		safeArray.PutElement(index,&var);
		index[1] += 1;
		tempStr.Format(_T("%ld"),pdiData.slabThickAct);
		var = _variant_t(tempStr);
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("钢板长度[mm]"),index))
	{
		index[1] += 2;
		tempStr.Format(_T("%ld"),pdiData.slabLengthL3);
		var = _variant_t(tempStr);
		safeArray.PutElement(index,&var);
		index[1] += 1;
		tempStr.Format(_T("%ld"),pdiData.slabLengthAct);
		var = _variant_t(tempStr);
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("板坯重量"),index))
	{
		index[1] += 2;
		tempStr.Format(_T("%ld"),pdiData.slabWeightL3);
		var = _variant_t(tempStr);
		safeArray.PutElement(index,&var);
		index[1] += 1;
		tempStr.Format(_T("%ld"),pdiData.slabWeightAct);
		var = _variant_t(tempStr);
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("厚度 [mm]"),index))
	{
		index[1] += 1;
		tempStr.Format(_T("%ld"),pdiData.plateThickSet);
		var = _variant_t(tempStr);
		safeArray.PutElement(index,&var);
		index[1] += 1;
		tempStr.Format(_T("%ld"),pdiData.plateThickAct);
		var = _variant_t(tempStr);
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("宽度 [mm]"),index))
	{
		index[1] += 1;
		tempStr.Format(_T("%ld"),pdiData.plateWidthSet);
		var = _variant_t(tempStr);
		safeArray.PutElement(index,&var);
		index[1] += 1;
		tempStr.Format(_T("%ld"),pdiData.plateWidthAct);
		var = _variant_t(tempStr);
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("长度 [mm]"),index))
	{
		index[1] += 1;
		tempStr.Format(_T("%ld"),pdiData.plateLengthSet);
		var = _variant_t(tempStr);
		safeArray.PutElement(index,&var);
		index[1] += 1;
		tempStr.Format(_T("%ld"),pdiData.plateLengthAct);
		var = _variant_t(tempStr);
		safeArray.PutElement(index,&var);
	}
}

void PDFCreater::SetRollSetup()
{
	const PDF::RollSetup &rollSetup = m_comm.m_newRollData.rollSetup;
	long index[2] = {0};
	VARIANT var;
	CString tempStr;
	if(Find(_T("轧制模式"),index))
	{
		index[1] += 1;
		tempStr.Format(_T("%d"),rollSetup.ctrlMode);
		var = _variant_t(tempStr);
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("总道次数"),index))
	{
		index[1] += 1;
		tempStr.Format(_T("%d"),rollSetup.totalPassNum);
		var = _variant_t(tempStr);
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("转钢模式"),index))
	{
		index[1] += 1;
		tempStr.Format(_T("%d"),rollSetup.turnMode);
		var = _variant_t(tempStr);
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("转钢道次"),index))
	{
		index[1] += 1;
		tempStr.Format(_T("%d"),rollSetup.turnPassNum);
		var = _variant_t(tempStr);
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("转钢厚度[mm]"),index))
	{
		index[1] += 1;
		tempStr.Format(_T("%ld"),rollSetup.turnThick);
		var = _variant_t(tempStr);
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("控轧厚度[mm]"),index))
	{
		index[1] += 1;
		tempStr.Format(_T("%ld"),rollSetup.controlThick);
		var = _variant_t(tempStr);
		safeArray.PutElement(index,&var);
	}
}

void PDFCreater::SetMillDataRM()
{
	const PDF::MillData &millDataRM = m_comm.m_newRollData.millDataRM;
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

void PDFCreater::SetMillDataFM()
{
	const PDF::MillData &millDataFM = m_comm.m_newRollData.millDataFM;
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

void PDFCreater::SetRollData(const PDF::RollData &rollData, const long (&startIndex)[2])
{
	VARIANT var;
	CString tempStr;
	long index[2] = {0};
	index[0] = startIndex[0];
	index[1] = startIndex[1];

	index[1] += 1;
	tempStr.Format(_T("%s"),rollData.RollID);
	var = _variant_t(tempStr);
	safeArray.PutElement(index,&var);
	index[1] += 1;
	tempStr.Format(_T("%s"),rollData.RollMaterial);
	var = _variant_t(tempStr);
	safeArray.PutElement(index,&var);
	index[1] += 1;
	tempStr.Format(_T("%ld"),rollData.RollDiameter);
	var = _variant_t(tempStr);
	safeArray.PutElement(index,&var);
	index[1] += 1;
	tempStr.Format(_T("%s"),rollData.RollStartTime);
	var = _variant_t(tempStr);
	safeArray.PutElement(index,&var);
}

void PDFCreater::SetPassData(const PDF::PassData (&passData)[PASS_MAX])
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
			var = _variant_t(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+2;
			tempStr.Format(_T("%ld"),passData[i].passGapSet);
			var = _variant_t(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+3;
			tempStr.Format(_T("%ld"),passData[i].passGapAct);
			var = _variant_t(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+4;
			tempStr.Format(_T("%ld"),passData[i].passThickSet);
			var = _variant_t(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+5;
			tempStr.Format(_T("%ld"),passData[i].passThickAct);
			var = _variant_t(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+6;
			tempStr.Format(_T("%ld"),passData[i].passWidthSet);
			var = _variant_t(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+7;
			tempStr.Format(_T("%ld"),passData[i].passWidthAct);
			var = _variant_t(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+8;
			tempStr.Format(_T("%ld"),passData[i].passForceSet);
			var = _variant_t(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+9;
			tempStr.Format(_T("%ld"),passData[i].passForceAct);
			var = _variant_t(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+10;
			tempStr.Format(_T("%ld"),passData[i].passTorqueSet);
			var = _variant_t(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+11;
			tempStr.Format(_T("%ld"),passData[i].passTorqueAct);
			var = _variant_t(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+12;
			tempStr.Format(_T("%ld"),passData[i].passBendForceSet);
			var = _variant_t(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+13;
			tempStr.Format(_T("%ld"),passData[i].passBendForceAct);
			var = _variant_t(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+14;
			tempStr.Format(_T("%ld"),passData[i].passShiftSet);
			var = _variant_t(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+15;
			tempStr.Format(_T("%ld"),passData[i].passShiftAct);
			var = _variant_t(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+16;
			tempStr.Format(_T("%ld"),passData[i].passVelocity);
			var = _variant_t(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+17;
			tempStr.Format(_T("%ld"),passData[i].passEntryTemp);
			var = _variant_t(tempStr);
			safeArray.PutElement(index,&var);
			index[1] = column+18;
			tempStr.Format(_T("%ld"),passData[i].passExitTemp);
			var = _variant_t(tempStr);
			safeArray.PutElement(index,&var);
		}
	}
}
