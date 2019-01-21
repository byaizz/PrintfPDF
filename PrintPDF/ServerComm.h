#pragma once

#include "NG_malloc.h"
#include "NG_errno.h"
#include "NG_types.h"
#include "NG_Handle.h"
#include "RollDataDefinition.h"

struct COMM_DATA_HEAD{
	short sTelID;//电文识别号
	short sTelHeaderLength;//电文头部长度字节数
	short sTelDataLength;//电文总长度字节数
	short sCounter;//1~32766
	char  szSpare[22];//备用
	short sHeaderSum;//电文头部字节校验和
};

class ServerComm
{
public:
	ServerComm(void);
public:
	~ServerComm(void);
private:
	enum DataLength{
		BUFF_MAX = 15000		//缓存最大长度
	};
private:
	HANDLE m_hComm;//通信句柄
	int m_iCount;//通道数
public:
	bool m_isNewRollData;//是否是新轧制信息
	ROLLDATA m_newRollData;//轧制信息
	void *m_pBuff;//缓存区

public:
	bool IsCommInit();//是否初始化通信
	bool InitComm();//初始化通信
	int CloseComm();//关闭通信
	int GetCommState();//获取通信状态
	void RecvData();//循环接受数据

	int SendData(void *pData, int iSize);//发送数据
	void RecvDataSpec(int iCount);//接受指定通道数据
};
