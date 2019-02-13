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

	//************************************************************************
	// Method:		CreateTwoDimSafeArray	������ά��ȫ����
	// Returns:		bool
	// Parameter:	COleSafeArray & safeArray	������
	// Parameter:	DWORD dimElements1	һά��Ԫ������
	// Parameter:	DWORD dimElements2	��ά��Ԫ������
	// Parameter:	long iLBound1	һά���±߽磬Ĭ��ֵ:1
	// Parameter:	long iLBound2	��ά���±߽磬Ĭ��ֵ:1
	// Author:		byshi
	// Date:		2018-12-10	
	//************************************************************************
	bool CreateTwoDimSafeArray(COleSafeArray &safeArray, DWORD dimElements1, 
		DWORD dimElements2, long iLBound1 = 1, long iLBound2 = 1);
	
	//************************************************************************
	// Method:		GetTime		��ȡ��ǰϵͳʱ��
	// Returns:		CString		���ص�ǰϵͳʱ��
	// Author:		byshi
	// Date:		2019-2-13	
	//************************************************************************
	CString GetTime();
	
	//************************************************************************
	// Method:		SetRollData		������������
	// Returns:		bool
	// Parameter:	const ROLLDATA & rollData	�������ݽṹ��
	// Author:		byshi
	// Date:		2019-1-25	
	//************************************************************************
	void SetRollingData();
	
	//************************************************************************
	// Method:		UpdateSafeArrayBound	���°�ȫ����ı߽�
	// Returns:		void
	// Author:		byshi
	// Date:		2019-1-30	
	//************************************************************************
	void UpdateSafeArrayBound();

protected:
		//************************************************************************
	// Method:		Find	����ָ���ı���λ��
	// Returns:		bool
	// Parameter:	const CString & text	��Ҫ��ѯ���ı�����
	// Parameter:	int 
	// Parameter:	& index		���ı����ڣ������ı�λ������
	// Parameter:	[2]
	// Author:		byshi
	// Date:		2019-1-28	
	//************************************************************************
	bool Find(const CString &text, long (&index)[2]);

	//************************************************************************
	// Method:		FindByColumn	��ָ���в���ָ���ı�������λ��
	// Returns:		bool
	// Parameter:	const CString & text	��Ҫ��ѯ���ı�����
	// Parameter:	const long column	Ҫ��ѯ���е�������
	// Parameter:	long & row	���ز�ѯ�����������
	// Author:		byshi
	// Date:		2019-1-29	
	//************************************************************************
	bool FindByColumn(const CString &text, const long column, long &row);

	//************************************************************************
	// Method:		SetTitle	���ñ���
	// Returns:		void
	// Author:		byshi
	// Date:		2019-1-29	
	//************************************************************************
	void SetTitle();
	
	//************************************************************************
	// Method:		SetPDIData	����PDI����
	// Returns:		bool
	// Parameter:	const PDF::PDIData & pdiData	PDI���ݽṹ��
	// Author:		byshi
	// Date:		2019-1-28	
	//************************************************************************
	void SetPDIData();
	
	//************************************************************************
	// Method:		SetRollSetup	���������趨��Ϣ
	// Returns:		bool
	// Parameter:	const PDF::RollSetup & rollSetup	�����趨��Ϣ���ݽṹ��
	// Author:		byshi
	// Date:		2019-1-28	
	//************************************************************************
	void SetRollSetup();
	
	//************************************************************************
	// Method:		SetMillDataRM	���ô�������������Ϣ
	// Returns:		bool
	// Parameter:	const PDF::MillData & millDataRM	��������������Ϣ���ݽṹ��
	// Author:		byshi
	// Date:		2019-1-28	
	//************************************************************************
	void SetMillDataRM();
	
	//************************************************************************
	// Method:		SetMillDataFM	���þ�������������Ϣ
	// Returns:		bool
	// Parameter:	const PDF::MillData & millDataFM	��������������Ϣ���ݽṹ��
	// Author:		byshi
	// Date:		2019-1-28	
	//************************************************************************
	void SetMillDataFM();
	
	//************************************************************************
	// Method:		SetRollData		����������Ϣ
	// Returns:		void
	// Parameter:	const PDF::RollData & rollData
	// Author:		byshi
	// Date:		2019-1-29	
	//************************************************************************
	void SetRollData(const PDF::RollData &rollData, const long (&startIndex)[2]);
	
	//************************************************************************
	// Method:		SetPassDatas		���õ�������
	// Returns:		bool
	// Parameter:	const PDF::PassData & passData	�������ݽṹ��
	// Author:		byshi
	// Date:		2019-1-28	
	//************************************************************************
	void SetPassDatas();
	
	//************************************************************************
	// Method:		SetPassData		���õ�����������
	// Returns:		void
	// Parameter:	const PDF::PassData & passData		����������Ϣ
	// Parameter:	const long & startIndex[2]		��ʼ����λ��
	// Author:		byshi
	// Date:		2019-2-12	
	//************************************************************************
	void SetPassData(const PDF::PassData &passData, const long (&startIndex)[2]);

private:
	ExcelManager m_excel;
	ServerComm m_comm;

	COleSafeArray	safeArray;//��ȫ���飬���ڱ���excel����
	long			firstLBound;//��һά��߽�
	long			firstUBound;//��һά�ұ߽�
	long			secondLBound;//�ڶ�ά��߽�
	long			secondUBound;//�ڶ�ά�ұ߽�
};
