#include <iostream>
#include <process.h>
#include "PDFCreater.h"
#include "basicMain.h"
#include "NG_sock.h"

bool g_exitFlag = true;//退出标记，false为退出

unsigned int _stdcall service_thread_func(void *exitFlag)
{
	//init
	NGSock_SysInit();

	if (S_OK != CoInitialize(NULL))//init
	{
		printf("初始化ole失败\n");
		return -1;
	}

	PDFCreater creater;
	creater.Init();//init
	while (*(bool*)exitFlag)
	{
		creater.Run();//run,任务执行函数
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
	g_exitFlag = false;//退出标记
	WaitForSingleObject(m_threadHandle, INFINITE);//等待线程结束
	CloseHandle(m_threadHandle);//关闭线程句柄
}
