#include "PDFCreater.h"
#include <iostream>

using namespace std;

static PDF::ROLLSchedule rollSchedu;
void CreateTestData()
{
	rollSchedu.millDataFM.downBackupRoll.RollDiameter = 1024;
	strcpy(rollSchedu.millDataFM.downBackupRoll.RollID,"RollID");
	strcpy(rollSchedu.millDataFM.downBackupRoll.RollMaterial,"RollMaterial");
	strcpy(rollSchedu.millDataFM.downBackupRoll.RollStartTime,"20190212");
	memcpy(&rollSchedu.millDataFM.upBackupRoll,&rollSchedu.millDataFM.downBackupRoll,sizeof(PDF::RollData));
	memcpy(&rollSchedu.millDataFM.downWorkRoll,&rollSchedu.millDataFM.downBackupRoll,sizeof(PDF::RollData));
	memcpy(&rollSchedu.millDataFM.upWorkRoll,&rollSchedu.millDataFM.downBackupRoll,sizeof(PDF::RollData));

	memcpy(&rollSchedu.millDataRM,&rollSchedu.millDataFM,sizeof(PDF::MillData));

	rollSchedu.rollSetup.totalPassNum = 25;

	strcpy(rollSchedu.pdiData.slabID,"slabID");
	strcpy(rollSchedu.pdiData.steelGrade,"steelGrade");

	for (int i = 0; i <= PASS_MAX; ++i)
	{
		rollSchedu.passData[i].millType = i % 2;
		rollSchedu.passData[i].passNo = i + 1;
		rollSchedu.passData[i].passGapSet = i + 1000;
		rollSchedu.passData[i].passGapAct = i + 1000;
		rollSchedu.passData[i].passThickSet = i + 1000;
		rollSchedu.passData[i].passThickAct = i + 1000;
		rollSchedu.passData[i].passWidthSet = i + 1000;
		rollSchedu.passData[i].passWidthAct = i + 1000;
		rollSchedu.passData[i].passForceSet = i + 1000;
		rollSchedu.passData[i].passForceAct = i + 1000;
		rollSchedu.passData[i].passTorqueSet = i + 1000;
		rollSchedu.passData[i].passTorqueAct = i + 1000;
		rollSchedu.passData[i].passBendForceSet = i + 1000;
		rollSchedu.passData[i].passBendForceAct = i + 1000;
		rollSchedu.passData[i].passShiftSet = i + 1;
		rollSchedu.passData[i].passShiftAct = i + 1;
		rollSchedu.passData[i].passVelocity = i + 1;
		rollSchedu.passData[i].passEntryTemp = i + 1000;
		rollSchedu.passData[i].passExitTemp = i + 1000;
	}
}

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

	//ѭ����ʼ��ͨ��
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
		cout<<"������ά�����쳣."<<endl;
		return false;//������ά����ʧ��
	}
	return true;
}

CString PDFCreater::GetTime()
{
	time_t t;
	time(&t);
	tm *local = localtime(&t);
	CString timeStr;
	timeStr.Format(_T("%d��%d��%d��%dʱ%d��%d��"), 
		local->tm_year + 1990, 
		local->tm_mon + 1, 
		local->tm_mday, 
		local->tm_hour, 
		local->tm_min, 
		local->tm_sec);
	return timeStr;
}

void PDFCreater::Run()
{
	//m_comm.RecvData();
 	//if (m_comm.m_isNewRollData == true)
 	{
		//�ڴ˴�������д�뵽excel�������Ϊpdf�ļ�
		if (m_excel.OpenFromTemplate(_T("E:\\pdf���ݱ�_wq.xlsx")))
		{
			CreateTestData();
			memcpy(&m_comm.m_newRollData, &rollSchedu, sizeof(PDF::ROLLSchedule));
			SetRollingData();
			CString fileName;//�ļ�����·��
			fileName.Format(_T("E:\\%s.pdf"),m_comm.m_newRollData.pdiData.slabID);
			m_excel.SaveAsPDF(fileName);
		}
		m_excel.Close();
		m_comm.m_isNewRollData = false;
 	}
}

void PDFCreater::SetRollingData()
{
	//��ȡ���ݲ���������д������
	m_excel.GetCellsValue(_T("$A$1"),_T("$Z$35"),safeArray);//��ȡ����
	UpdateSafeArrayBound();//���°�ȫ����ı߽���Ϣ
	SetTitle();//���ñ���
	SetPDIData();//����PDI��Ϣ
	SetRollSetup();//���������趨��Ϣ
	SetMillDataRM();//���ô�������������Ϣ
	SetMillDataFM();//���þ�������������Ϣ
	m_excel.SetCellsValue(_T("$A$1"),_T("$Z$35"),safeArray);//д������

	//��ǰÿ��sheet��д��20����������Ϣ���ɸ�����Ҫ�޸�
	int totalPassNum = m_comm.m_newRollData.rollSetup.totalPassNum;
	int sheetNum = totalPassNum / PASS_COUNT;//��ҳ��
	if (totalPassNum % PASS_COUNT != 0)
	{
		++sheetNum;
	}
	for (int i =0; i < sheetNum - 1; ++i)
	{
		if (m_excel.CopyCurrentSheet())
		{
			cout<<"���Ƶ�ǰsheet�ɹ�."<<endl;
		}
	}
	SetPassDatas();//���õ�����Ϣ

	// 	CPageSetup pageSetUp;
	// 	pageSetUp.AttachDispatch(m_worksheet.get_PageSetup());
	// 	long xlLandscape = 2;//����ģʽ
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
	long index[2] = {0};
	VARIANT var;
	CString tempStr;

	if (Find(_T("**�������޹�˾�к��������"),index))
	{
		tempStr = _T("�Ͼ��������޹�˾�к��������");
		var = _variant_t(tempStr);
		safeArray.PutElement(index,&var);//
	}
	if (Find(_T("���ƹ�̱� �ֺ�:   ������ʱ�䣺  "),index))
	{
		tempStr.Format(_T("���ƹ�̱� �ֺ�:%s, ����ʱ�䣺"), 
			m_comm.m_newRollData.pdiData.slabID);	
		CString timeStr = GetTime();
		tempStr += timeStr;
		var = _variant_t(tempStr);
		safeArray.PutElement(index,&var);//
	}
}

void PDFCreater::SetPDIData()
{
	const PDF::PDIData &pdiData = m_comm.m_newRollData.pdiData;
	long index[2] = {0};
	VARIANT var;
	CString tempStr;
	if (Find(_T("����"),index))
	{
		index[1] += 2;
		tempStr.Format(_T("%s"),pdiData.steelGrade);
		var = _variant_t(tempStr);
		safeArray.PutElement(index,&var);
		index[1] += 1;
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("����¯��"),index))
	{
		index[1] += 3;
		tempStr.Format(_T("%d-%d"),pdiData.furnaceNo,pdiData.furnaceCol);
		var = _variant_t(tempStr);
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("�����¶� [oC]"),index))
	{
		index[1] += 2;
		tempStr.Format(_T("%d"),pdiData.finalTemp);
		var = _variant_t(tempStr);
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("������� [mm]"),index))
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
	if(Find(_T("������� [mm]"),index))
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
	if(Find(_T("�ְ峤��[mm]"),index))
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
	if(Find(_T("��������"),index))
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
	if(Find(_T("��� [mm]"),index))
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
	if(Find(_T("��� [mm]"),index))
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
	if(Find(_T("���� [mm]"),index))
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
	if(Find(_T("����ģʽ"),index))
	{
		index[1] += 1;
		tempStr.Format(_T("%d"),rollSetup.ctrlMode);
		var = _variant_t(tempStr);
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("�ܵ�����"),index))
	{
		index[1] += 1;
		tempStr.Format(_T("%d"),rollSetup.totalPassNum);
		var = _variant_t(tempStr);
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("ת��ģʽ"),index))
	{
		index[1] += 1;
		tempStr.Format(_T("%d"),rollSetup.turnMode);
		var = _variant_t(tempStr);
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("ת�ֵ���"),index))
	{
		index[1] += 1;
		tempStr.Format(_T("%d"),rollSetup.turnPassNum);
		var = _variant_t(tempStr);
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("ת�ֺ��[mm]"),index))
	{
		index[1] += 1;
		tempStr.Format(_T("%ld"),rollSetup.turnThick);
		var = _variant_t(tempStr);
		safeArray.PutElement(index,&var);
	}
	if(Find(_T("�������[mm]"),index))
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
	if (Find(_T("��������������"), index))
	{
		long row = 0;
		long column = index[1];
		if(FindByColumn(_T("��֧�Ź�"),column,row))
		{
			index[0] = row;
			SetRollData(millDataRM.upBackupRoll,index);
		}
		if(FindByColumn(_T("�Ϲ�����"),column,row))
		{
			index[0] = row;
			SetRollData(millDataRM.upWorkRoll,index);
		}
		if(FindByColumn(_T("�¹�����"),column,row))
		{
			index[0] = row;
			SetRollData(millDataRM.downWorkRoll,index);
		}
		if(FindByColumn(_T("��֧�Ź�"),column,row))
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
	if (Find(_T("��������������"), index))
	{
		long row = 0;
		long column = index[1];
		if(FindByColumn(_T("��֧�Ź�"),column,row))
		{
			index[0] = row;
			SetRollData(millDataFM.upBackupRoll,index);
		}
		if(FindByColumn(_T("�Ϲ�����"),column,row))
		{
			index[0] = row;
			SetRollData(millDataFM.upWorkRoll,index);
		}
		if(FindByColumn(_T("�¹�����"),column,row))
		{
			index[0] = row;
			SetRollData(millDataFM.downWorkRoll,index);
		}
		if(FindByColumn(_T("��֧�Ź�"),column,row))
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

void PDFCreater::SetPassDatas()
{
	const PDF::PassData *passDatas = m_comm.m_newRollData.passData;
	long index[2] = {0};
	long sheetCount = m_excel.GetSheetCount();//��ȡsheet����

	//Ϊÿ��sheetд�����������Ϣ
	for (long sheetIndex = 1; sheetIndex <= sheetCount; ++sheetIndex)
	{
		m_excel.SwitchWorksheet(sheetIndex);//�л�sheet
		m_excel.GetCellsValue(_T("$A$13"),_T("$Z$35"),safeArray);//��ȡ����
		UpdateSafeArrayBound();//���°�ȫ����ı߽���Ϣ

		if(Find(_T("������RM/FM)"),index))
		{
			index[0] += 2;
			for (int i = 0; i < PASS_COUNT; ++i)
			{
				int passIndex = i + (sheetIndex-1)*PASS_COUNT;
				//����У��
				if (passIndex >= PASS_MAX || passDatas[passIndex].passNo <= 0)
				{
					break;
				}
				//���õ�����Ϣ������
				SetPassData(passDatas[passIndex],index);
				index[0] += 1;
			}
		}
		m_excel.SetCellsValue(_T("$A$13"),_T("$Z$35"),safeArray);//д�����ݵ����
	}
}

void PDFCreater::SetPassData(const PDF::PassData &passData, const long (&startIndex)[2])
{
	VARIANT var;//��ʱ�洢����
	CString tempStr;//��ʱ�洢����

	long index[2] = {0};
	index[0] = startIndex[0];
	index[1] = startIndex[1];

	if (passData.millType == PDF::PassData::MILL_TYPE::RM)
	{
		tempStr = _T("RM");
	}
	else if (passData.millType == PDF::PassData::MILL_TYPE::FM)
	{
		tempStr = _T("FM");
	}
	var = _variant_t(tempStr);
	safeArray.PutElement(index,&var);//��������
	index[1] += 1;
	tempStr.Format(_T("%d"),passData.passNo);
	var = _variant_t(tempStr);
	safeArray.PutElement(index,&var);//���κ�
	index[1] += 1;
	tempStr.Format(_T("%.2lf"),passData.passGapSet);
	var = _variant_t(tempStr);
	safeArray.PutElement(index,&var);//���ι����趨ֵ
	index[1] += 1;
	tempStr.Format(_T("%.2lf"),passData.passGapAct);
	var = _variant_t(tempStr);
	safeArray.PutElement(index,&var);//���ι���ʵ��ֵ
	index[1] += 1;
	tempStr.Format(_T("%.2lf"),passData.passThickSet);
	var = _variant_t(tempStr);
	safeArray.PutElement(index,&var);//���κ���趨ֵ
	index[1] += 1;
	tempStr.Format(_T("%.2lf"),passData.passThickAct);
	var = _variant_t(tempStr);
	safeArray.PutElement(index,&var);//���κ��ʵ��ֵ
	index[1] += 1;
	tempStr.Format(_T("%.2lf"),passData.passWidthSet);
	var = _variant_t(tempStr);
	safeArray.PutElement(index,&var);//���ο���趨ֵ
	index[1] += 1;
	tempStr.Format(_T("%.2lf"),passData.passWidthAct);
	var = _variant_t(tempStr);
	safeArray.PutElement(index,&var);//���ο��ʵ��ֵ
	index[1] += 1;
	tempStr.Format(_T("%.2lf"),passData.passForceSet);
	var = _variant_t(tempStr);
	safeArray.PutElement(index,&var);//�����������趨ֵ
	index[1] += 1;
	tempStr.Format(_T("%.2lf"),passData.passForceAct);
	var = _variant_t(tempStr);
	safeArray.PutElement(index,&var);//����������ʵ��ֵ
	index[1] += 1;
	tempStr.Format(_T("%.2lf"),passData.passTorqueSet);
	var = _variant_t(tempStr);
	safeArray.PutElement(index,&var);//���������趨ֵ
	index[1] += 1;
	tempStr.Format(_T("%.2lf"),passData.passTorqueAct);
	var = _variant_t(tempStr);
	safeArray.PutElement(index,&var);//��������ʵ��ֵ
	index[1] += 1;
	tempStr.Format(_T("%.2lf"),passData.passBendForceSet);
	var = _variant_t(tempStr);
	safeArray.PutElement(index,&var);//����������趨ֵ
	index[1] += 1;
	tempStr.Format(_T("%.2lf"),passData.passBendForceAct);
	var = _variant_t(tempStr);
	safeArray.PutElement(index,&var);//���������ʵ��ֵ
	index[1] += 1;
	tempStr.Format(_T("%.2lf"),passData.passShiftSet);
	var = _variant_t(tempStr);
	safeArray.PutElement(index,&var);//���δܹ��趨ֵ
	index[1] += 1;
	tempStr.Format(_T("%.2lf"),passData.passShiftAct);
	var = _variant_t(tempStr);
	safeArray.PutElement(index,&var);//���δܹ�ʵ��ֵ
	index[1] += 1;
	tempStr.Format(_T("%.2lf"),passData.passVelocity);
	var = _variant_t(tempStr);
	safeArray.PutElement(index,&var);//�����ٶ�
	index[1] += 1;
	tempStr.Format(_T("%d"),passData.passEntryTemp);
	var = _variant_t(tempStr);
	safeArray.PutElement(index,&var);//��������¶�
	index[1] += 1;
	tempStr.Format(_T("%d"),passData.passExitTemp);
	var = _variant_t(tempStr);
	safeArray.PutElement(index,&var);//���γ����¶�
}
