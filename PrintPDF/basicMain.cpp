#include "basicMain.h"
#include <windows.h>
#include <thread>
#include <fstream>
#include <iostream>

BasicMain::BasicMain(const char* serviceName, const char* displayName)
	:ServiceBase(serviceName, displayName)
{
}

BasicMain::~BasicMain()
{
}

void BasicMain::service_thread_func()
{
	
}

void BasicMain::OnStart(DWORD argc, TCHAR * argv[])
{
	std::thread service_thread(service_thread_func);
	service_thread.detach();
}
