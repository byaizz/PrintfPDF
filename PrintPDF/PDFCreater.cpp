#include "PDFCreater.h"

PDFCreater::PDFCreater(void)
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

void PDFCreater::Run()
{
	m_comm.RecvData();
 	if (m_comm.m_isNewRollData == true)
 	{
		//�ڴ˴�������д�뵽excel�������Ϊpdf�ļ�
// 		if (m_excel.OpenFromTemplate(_T("E:\\test_by.xlsx")))
// 		{
// 			m_excel.Test1();
// 			m_excel.SaveAsPDF(_T("E:\\test_by.pdf"));
// 			m_comm.m_isNewRollData = false;
// 		}
// 		m_excel.Close();
		m_comm.m_isNewRollData = false;
 	}
}
