#pragma once
#include "ServerComm.h"
#include "ExcelManager.h" 

class PDFCreater
{
public:
	PDFCreater(void);
	~PDFCreater(void);
	void Init();//��ʼ��
	void Run();//run
private:
	ExcelManager m_excel;
	ServerComm m_comm;
};
