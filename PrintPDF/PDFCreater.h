#pragma once
#include "PRASGuiSerComm.h"
#include "ExcelManager.h" 

class PDFCreater
{
public:
	PDFCreater(void);
	~PDFCreater(void);
	void Init();
	void Run();
private:
	ExcelManager m_excel;
	CPRASGuiSerComm m_comm;
};
