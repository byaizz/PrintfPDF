#ifndef BASICMAIN_H
#define BASICMAIN_H

#include "serviceBase.h"

extern bool g_exitFlag;
unsigned int _stdcall service_thread_func(void *exitFlag);

class BasicMain : public ServiceBase{
public:
	BasicMain(const char* serviceName, const char* displayName);
	~BasicMain();

protected:
	void OnStart(DWORD argc, TCHAR* argv[]);
	void OnStop();
private:
	HANDLE m_threadHandle;
};

#endif	//BASICMAIN_H