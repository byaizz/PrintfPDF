#include "basicMain.h"
#include <windows.h>
#include <iostream>
#include <process.h>

bool g_exitFlag = true;

unsigned int _stdcall service_thread_func(void *exitFlag)
{
	HANDLE hThread;

	printf("Creating second thread...\n");

	// Create thread.
	hThread = (HANDLE)_beginthreadex(NULL, 0, receive_func, exitFlag, 0, NULL);


	while (*(bool*)exitFlag)
	{
		printf("I'm in service_thread_func.\n");
		Sleep(1000);
	}

	printf("waiting for receive_func finished.\n");
	WaitForSingleObject(hThread, INFINITE);
	// Destroy the thread object.
	CloseHandle(hThread);
	return 0;
}

unsigned int _stdcall receive_func(void *exitFlag)
{
	while(*(bool*)exitFlag)
	{
		printf("I'm in receive_func.\n");
		Sleep(1000);
	}
	printf("receive_func will be finished\n");
	Sleep(1000);;
	printf("receive_func is finished\n");
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
}
