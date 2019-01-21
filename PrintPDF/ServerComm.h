#pragma once

#include "NG_malloc.h"
#include "NG_errno.h"
#include "NG_types.h"
#include "NG_Handle.h"
#include "RollDataDefinition.h"

struct COMM_DATA_HEAD{
	short sTelID;//����ʶ���
	short sTelHeaderLength;//����ͷ�������ֽ���
	short sTelDataLength;//�����ܳ����ֽ���
	short sCounter;//1~32766
	char  szSpare[22];//����
	short sHeaderSum;//����ͷ���ֽ�У���
};

class ServerComm
{
public:
	ServerComm(void);
public:
	~ServerComm(void);
private:
	enum DataLength{
		BUFF_MAX = 15000		//������󳤶�
	};
private:
	HANDLE m_hComm;//ͨ�ž��
	int m_iCount;//ͨ����
public:
	bool m_isNewRollData;//�Ƿ�����������Ϣ
	ROLLDATA m_newRollData;//������Ϣ
	void *m_pBuff;//������

public:
	bool IsCommInit();//�Ƿ��ʼ��ͨ��
	bool InitComm();//��ʼ��ͨ��
	int CloseComm();//�ر�ͨ��
	int GetCommState();//��ȡͨ��״̬
	void RecvData();//ѭ����������

	int SendData(void *pData, int iSize);//��������
	void RecvDataSpec(int iCount);//����ָ��ͨ������
};
