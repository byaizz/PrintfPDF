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

	//循环初始化通信
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
		//在此处将数据写入到excel，并输出为pdf文件
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
