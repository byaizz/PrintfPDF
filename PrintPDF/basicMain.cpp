#include <iostream>
#include <process.h>
#include "PDFCreater.h"
#include "basicMain.h"
#include "NG_sock.h"

bool g_exitFlag = true;

unsigned int _stdcall service_thread_func(void *exitFlag)
{
	NGSock_SysInit();
	PDFCreater creater;
	creater.Init();
	while (*(bool*)exitFlag)
	{
		creater.Run();
		Sleep(1000);
	}

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
	g_exitFlag = false;
	WaitForSingleObject(m_threadHandle, INFINITE);
	CloseHandle(m_threadHandle);
	NGSock_SysClose();
}
