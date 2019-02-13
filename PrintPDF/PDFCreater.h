#pragma once
#include "ServerComm.h"
#include "ExcelManager.h" 

class PDFCreater
{
public:
	PDFCreater(void);
	~PDFCreater(void);
	void Init();//初始化
	void Run();//run

	//************************************************************************
	// Method:		CreateTwoDimSafeArray	创建二维安全数组
	// Returns:		bool
	// Parameter:	COleSafeArray & safeArray	数组名
	// Parameter:	DWORD dimElements1	一维的元素数量
	// Parameter:	DWORD dimElements2	二维的元素数量
	// Parameter:	long iLBound1	一维的下边界，默认值:1
	// Parameter:	long iLBound2	二维的下边界，默认值:1
	// Author:		byshi
	// Date:		2018-12-10	
	//************************************************************************
	bool CreateTwoDimSafeArray(COleSafeArray &safeArray, DWORD dimElements1, 
		DWORD dimElements2, long iLBound1 = 1, long iLBound2 = 1);
	
	//************************************************************************
	// Method:		GetTime		获取当前系统时间
	// Returns:		CString		返回当前系统时间
	// Author:		byshi
	// Date:		2019-2-13	
	//************************************************************************
	CString GetTime();
	
	//************************************************************************
	// Method:		SetRollData		设置轧制数据
	// Returns:		bool
	// Parameter:	const ROLLDATA & rollData	轧制数据结构体
	// Author:		byshi
	// Date:		2019-1-25	
	//************************************************************************
	void SetRollingData();
	
	//************************************************************************
	// Method:		UpdateSafeArrayBound	更新安全数组的边界
	// Returns:		void
	// Author:		byshi
	// Date:		2019-1-30	
	//************************************************************************
	void UpdateSafeArrayBound();

protected:
		//************************************************************************
	// Method:		Find	查找指定文本的位置
	// Returns:		bool
	// Parameter:	const CString & text	需要查询的文本内容
	// Parameter:	int 
	// Parameter:	& index		若文本存在，保存文本位置索引
	// Parameter:	[2]
	// Author:		byshi
	// Date:		2019-1-28	
	//************************************************************************
	bool Find(const CString &text, long (&index)[2]);

	//************************************************************************
	// Method:		FindByColumn	从指定列查找指定文本的索引位置
	// Returns:		bool
	// Parameter:	const CString & text	需要查询的文本内容
	// Parameter:	const long column	要查询的列的列索引
	// Parameter:	long & row	返回查询结果的行索引
	// Author:		byshi
	// Date:		2019-1-29	
	//************************************************************************
	bool FindByColumn(const CString &text, const long column, long &row);

	//************************************************************************
	// Method:		SetTitle	设置标题
	// Returns:		void
	// Author:		byshi
	// Date:		2019-1-29	
	//************************************************************************
	void SetTitle();
	
	//************************************************************************
	// Method:		SetPDIData	设置PDI数据
	// Returns:		bool
	// Parameter:	const PDF::PDIData & pdiData	PDI数据结构体
	// Author:		byshi
	// Date:		2019-1-28	
	//************************************************************************
	void SetPDIData();
	
	//************************************************************************
	// Method:		SetRollSetup	设置轧制设定信息
	// Returns:		bool
	// Parameter:	const PDF::RollSetup & rollSetup	轧制设定信息数据结构体
	// Author:		byshi
	// Date:		2019-1-28	
	//************************************************************************
	void SetRollSetup();
	
	//************************************************************************
	// Method:		SetMillDataRM	设置粗轧轧机轧辊信息
	// Returns:		bool
	// Parameter:	const PDF::MillData & millDataRM	粗轧轧机轧辊信息数据结构体
	// Author:		byshi
	// Date:		2019-1-28	
	//************************************************************************
	void SetMillDataRM();
	
	//************************************************************************
	// Method:		SetMillDataFM	设置精轧轧机轧辊信息
	// Returns:		bool
	// Parameter:	const PDF::MillData & millDataFM	精轧轧机轧辊信息数据结构体
	// Author:		byshi
	// Date:		2019-1-28	
	//************************************************************************
	void SetMillDataFM();
	
	//************************************************************************
	// Method:		SetRollData		设置轧辊信息
	// Returns:		void
	// Parameter:	const PDF::RollData & rollData
	// Author:		byshi
	// Date:		2019-1-29	
	//************************************************************************
	void SetRollData(const PDF::RollData &rollData, const long (&startIndex)[2]);
	
	//************************************************************************
	// Method:		SetPassDatas		设置道次数据
	// Returns:		bool
	// Parameter:	const PDF::PassData & passData	道次数据结构体
	// Author:		byshi
	// Date:		2019-1-28	
	//************************************************************************
	void SetPassDatas();
	
	//************************************************************************
	// Method:		SetPassData		设置单个道次数据
	// Returns:		void
	// Parameter:	const PDF::PassData & passData		道次数据信息
	// Parameter:	const long & startIndex[2]		起始索引位置
	// Author:		byshi
	// Date:		2019-2-12	
	//************************************************************************
	void SetPassData(const PDF::PassData &passData, const long (&startIndex)[2]);

private:
	ExcelManager m_excel;
	ServerComm m_comm;

	COleSafeArray	safeArray;//安全数组，用于保存excel数据
	long			firstLBound;//第一维左边界
	long			firstUBound;//第一维右边界
	long			secondLBound;//第二维左边界
	long			secondUBound;//第二维右边界
};
