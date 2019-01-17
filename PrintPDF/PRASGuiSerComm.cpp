#include "PRASGuiSerComm.h"
#include "NG_Path.h"
#include "MP_Comm.h"
#include "NG_sock.h"
#include <fstream>
#include <iostream>

#define GUI_SERVICE_COMM_CFG	"PRASGUIServiceComm.cfg"
#define BUFF_MAX	5000

CPRASGuiSerComm::CPRASGuiSerComm(void)
:m_hComm(NULL)
,m_iCount(0)
,m_isNewRollData(false)
{
	m_pBuff = NG_malloc(BUFF_MAX);
	memset(&m_newRollData,0,sizeof(m_newRollData));
}

CPRASGuiSerComm::~CPRASGuiSerComm(void)
{
	CloseComm();
	if (NULL != m_pBuff)
	{
		NG_free(m_pBuff);
	}
}

bool CPRASGuiSerComm::InitComm()
{
	int ret = 0;
	char path[512] = {0};

	//获取配置文件全路径名
	ret = NG_Path_GetRunPath(path,sizeof(path)-1);
	if (ret == ERR_FAILED)
	{
		return false;
	}
	strcat(path,GUI_SERVICE_COMM_CFG);
	std::cout<<path<<std::endl;

	//创建通信连接
	m_hComm = MP_Comm_Create(path);
	if (m_hComm == NULL)
	{
		return false;
	}
	

	m_iCount = MP_Comm_GetTunnelCount(m_hComm);
	std::cout<<m_iCount<<std::endl;
	return true;
}

int CPRASGuiSerComm::CloseComm()
{
	if (m_hComm != NULL)
	{
		NG_CloseHandle(m_hComm);
		m_hComm = NULL;
	}
	return ERR_SUCCESS;
}

int CPRASGuiSerComm::GetCommState()
{
	if (m_hComm == NULL)
	{
		return ERR_FAILED;
	}

	for (int i = 0; i < m_iCount; ++i)
	{
		if (0 != MP_Comm_GetStatus(m_hComm,i))
		{
			return ERR_FAILED;
		}
	}
	return ERR_SUCCESS;
}

void CPRASGuiSerComm::RecvData()
{
	for (int i = 0; i < m_iCount; ++i)
	{
		RecvDataSpec(i);
	}
}

void CPRASGuiSerComm::RecvDataSpec(int iCount)
{
	if (!IsCommInit())
	{
		return;
	}
	if (!m_isNewRollData)
	{
		int iSize = 0;
		iSize = MP_Comm_RecvNoBlock(m_hComm,iCount,m_pBuff,BUFF_MAX);
		if (iSize == sizeof(m_newRollData) && iSize <= BUFF_MAX)
		{
			m_isNewRollData = true;
			memcpy(&m_newRollData, m_pBuff, sizeof(m_newRollData));
			printf("接收到轧钢信息!\n");
		}
		printf("---%d---\n",iSize);
	}
}

void CPRASGuiSerComm::CommHeadInit(COMM_DATA_HEAD& head, int iID, int iSize)
{
	head.sTelID = iID;
	head.sTelHeaderLength = sizeof(COMM_DATA_HEAD);
	head.sTelDataLength = iSize;
	head.sTelHeaderLength = 0;//等待修改
}

void CPRASGuiSerComm::SendHeart()
{
	int i = 0;
	SendData(&i, sizeof(int));
}

int CPRASGuiSerComm::SendData(void *pData, int iSize)
{
	if(pData == NULL)
	{
		return ERR_FAILED;
	}
	if (iSize <= 0)
	{
		return ERR_FAILED;
	}
	if (!IsCommInit())
	{
		return ERR_FAILED;
	}
	int res = ERR_FAILED;
	for (int i = 0; i < m_iCount; ++i)
	{
		int ret = MP_Comm_SendNoBlock(m_hComm,i,pData,iSize);
		if (ERR_SUCCESS != ret)
		{
			res = ERR_FAILED;
		}
	}
	return res;
}

bool CPRASGuiSerComm::IsCommInit()
{
	return (m_hComm != NULL);
}