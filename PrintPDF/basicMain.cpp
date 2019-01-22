#include <iostream>
#include <process.h>
#include "PDFCreater.h"
#include "basicMain.h"
#include "NG_sock.h"

bool g_exitFlag = true;//�˳���ǣ�falseΪ�˳�

unsigned int _stdcall service_thread_func(void *exitFlag)
{
	//init
	NGSock_SysInit();

	if (S_OK != CoInitialize(NULL))//init
	{
		printf("��ʼ��oleʧ��\n");
		return -1;
	}

	PDFCreater creater;
	creater.Init();//init
	while (*(bool*)exitFlag)
	{
		creater.Run();//run,����ִ�к���
		Sleep(1000);
	}

	CoUninitialize();//uninitialize
	NGSock_SysClose();//close
	return 0;
}

BasicMain::BasicMain(const char* serviceName, const char* displayName)
	:ServiceBase(serviceName, displayName)
	,m_threadHandle(NULL)
{
}

BasicMain::~BasicMain()
{
}

void BasicMain::OnStart(DWORD argc, TCHAR * argv[])
{
	// Create thread.
	m_threadHandle = (HANDLE)_beginthreadex(NULL, 0, service_thread_func, (void *)&g_exitFlag, 0, NULL);
}

void BasicMain::OnStop()
{
	g_exitFlag = false;//�˳����
	WaitForSingleObject(m_threadHandle, INFINITE);//�ȴ��߳̽���
	CloseHandle(m_threadHandle);//�ر��߳̾��
}
