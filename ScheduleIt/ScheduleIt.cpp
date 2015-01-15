// ScheduleIt.cpp : Defines the entry point for the console application.

#include "stdafx.h"

#include <Windows.h>
#include <iostream>
#include <stdio.h>
#include <comdef.h>
#include <wincred.h>

#include <taskschd.h>
# pragma comment(lib, "Taskschd.lib")
# pragma comment(lib, "comsupp.lib")
# pragma comment(lib, "credui.lib")

using namespace std;

//Based on example found at http://goo.gl/dht51R (MSDN Task scheduler example)
//Takes a unicode string indicating the path to the executable to schedule.
int scheduleNotepad(wstring *pwstrExecutablePath, wstring *pwstrCommandArguments){
	
	//Recommend only changing these when integrating into other projects
	/*****************************************************************************/
	// hardcoded values that can become function arguments if we have time
	_bstr_t taskStartTime = _bstr_t(L"2005-01-15T16:11:00");
	LPCWSTR wszTaskName = L"SELF SCHEDULED TASK WITH ARGS TEST";
	
	//path to the executable to schedule.
	wstring wstrExecutablePath = *pwstrExecutablePath;
	//Command line arguments
	wstring wstrCommandArguments = *pwstrCommandArguments;
	// END hardcoded values ...
	/*****************************************************************************/

	HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
	if (FAILED(hr)){
		printf("\nCoInitializeEx failed: %x", hr);
		return 1;
	}

	hr = CoInitializeSecurity(
		NULL,
		-1,
		NULL,
		NULL,
		RPC_C_AUTHN_LEVEL_PKT_PRIVACY,
		RPC_C_IMP_LEVEL_IMPERSONATE,
		NULL,
		0,
		NULL);

	if (FAILED(hr)){
		printf("\nCoInitializeSecurity failed: %x", hr);
		CoUninitialize();
		return 1;
	}


	// -----------------------------------------------------------
	// Create an instance of the Task Service
	ITaskService *pService = NULL;
	hr = CoCreateInstance(
		CLSID_TaskScheduler,
		NULL,
		CLSCTX_INPROC_SERVER,
		IID_ITaskService,
		(void**)&pService);
	if (FAILED(hr)){
		printf("Failed to create an instance of ITaskService: %x", hr);
		CoUninitialize();
		return 1;
	}

	// connect to the task service
	hr = pService->Connect(_variant_t(), _variant_t(), _variant_t(), _variant_t());
	if (FAILED(hr)){
		printf("ITakService::Connect failed: %x", hr);
		pService->Release();
		CoUninitialize();
		return 1;
	}
	//  ------------------------------------------------------
	//  Get the pointer to the root task folder.  This folder will hold the
	//  new task that is registered.
	ITaskFolder *pRootFolder = NULL;
	hr = pService->GetFolder(_bstr_t(L"\\"), &pRootFolder);
	if (FAILED(hr))
	{
		printf("Cannot get Root folder pointer: %x", hr);
		pService->Release();
		CoUninitialize();
		return 1;
	}

	// if the same task exists, remove it.
	// Must do this or creating the task will fail if that name already exists
	pRootFolder->DeleteTask(_bstr_t(wszTaskName), 0);

	//create the task definition object to create the task
	ITaskDefinition *pTask = NULL;
	hr = pService->NewTask(0, &pTask);

	pService->Release();	//COM cleanup. pointer is no longer used.
	if (FAILED(hr)){
		printf("Failed to CoCreate an instance of the TaskService class: %x", hr);
		pRootFolder->Release();
		CoUninitialize();
		return 1;
	}

	// -----------------------------------------------------------
	// Get the registration info for setting the identification
	IRegistrationInfo *pRegInfo = NULL;
	hr = pTask->get_RegistrationInfo(&pRegInfo);
	if (FAILED(hr)){
		printf("\nCannot get identification pointer: %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	hr = pRegInfo->put_Author(L"Author Name");
	pRegInfo->Release();
	if (FAILED(hr)){
		printf("\nCannot put identification info: %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	// ---------------------------------------------------------
	// Create the principal for the task - these credentials are
	// overwritten with the credentials passed to RegisterTaskDefinition
	IPrincipal *pPrincipal = NULL;
	hr = pTask->get_Principal(&pPrincipal);
	if (FAILED(hr)){
		printf("\nCannot get principal pointer: %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	//  Set up principal logon type to interactive logon
	hr = pPrincipal->put_LogonType(TASK_LOGON_INTERACTIVE_TOKEN);
	pPrincipal->Release();
	if (FAILED(hr))
	{
		printf("\nCannot put principal info: %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	//  ------------------------------------------------------
	//  Create the settings for the task
	ITaskSettings *pSettings = NULL;
	hr = pTask->get_Settings(&pSettings);
	if (FAILED(hr))
	{
		printf("\nCannot get settings pointer: %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	//  Set setting values for the task.  
	hr = pSettings->put_StartWhenAvailable(VARIANT_TRUE);
	if (FAILED(hr))
	{
		printf("\nCannot put setting information: %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	hr = pSettings->put_DisallowStartIfOnBatteries((VARIANT_BOOL)false);
	if (FAILED(hr))
	{
		printf("\nCannot put DisallowStartIfOnBatteries information: %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	hr = pSettings->put_MultipleInstances(TASK_INSTANCES_PARALLEL);
	pSettings->Release();
	if (FAILED(hr))
	{
		printf("\nCannot put parallel instance information: %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	// Set the idle settings for the task.
	IIdleSettings *pIdleSettings = NULL;
	hr = pSettings->get_IdleSettings(&pIdleSettings);
	if (FAILED(hr))
	{
		printf("\nCannot get idle setting information: %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	hr = pIdleSettings->put_WaitTimeout(L"PT5M");
	pIdleSettings->Release();
	if (FAILED(hr))
	{
		printf("\nCannot put idle setting information: %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}


	//  ------------------------------------------------------
	//  Get the trigger collection to insert the time trigger.
	ITriggerCollection *pTriggerCollection = NULL;
	hr = pTask->get_Triggers(&pTriggerCollection);
	if (FAILED(hr))
	{
		printf("\nCannot get trigger collection: %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	//  Add the time trigger to the task.
	ITrigger *pTrigger = NULL;
	hr = pTriggerCollection->Create(TASK_TRIGGER_TIME, &pTrigger);
	pTriggerCollection->Release();
	if (FAILED(hr))
	{
		printf("\nCannot create trigger: %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	ITimeTrigger *pTimeTrigger = NULL;
	hr = pTrigger->QueryInterface(
		IID_ITimeTrigger, (void**)&pTimeTrigger);
	pTrigger->Release();
	if (FAILED(hr))
	{
		printf("\nQueryInterface call failed for ITimeTrigger: %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	hr = pTimeTrigger->put_Id(_bstr_t(L"Trigger1"));
	if (FAILED(hr))
		printf("\nCannot put trigger ID: %x", hr);

	IRepetitionPattern *pRepetitionPattern = NULL;
	hr = pTimeTrigger->get_Repetition(&pRepetitionPattern);
	if (FAILED(hr)){
		printf("\nFailed to get repetition pattern: %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}
	//run every 20 minutes
	pRepetitionPattern->put_Interval(_bstr_t(L"PT20M"));
	hr = pTimeTrigger->put_Repetition(pRepetitionPattern);
	if (FAILED(hr)){
		printf("\nFailed to put repetition pattern: %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}
	/*	hr = pTimeTrigger->put_EndBoundary(_bstr_t(L"2015-05-02T08:00:00"));
	//	if (FAILED(hr))
	//		printf("\nCannot put end boundary on trigger: %x", hr);
	*/
	//  Set the task to start at a certain time. The time 
	//  format should be YYYY-MM-DDTHH:MM:SS(+-)(timezone).
	//  For example, the start boundary below
	//  is January 1st 2005 at 12:05
	hr = pTimeTrigger->put_StartBoundary(taskStartTime);
	pTimeTrigger->Release();
	if (FAILED(hr))
	{
		printf("\nCannot add start boundary to trigger: %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}


	//  ------------------------------------------------------
	//  Add an action to the task. This task will execute notepad.exe.     
	IActionCollection *pActionCollection = NULL;

	//  Get the task action collection pointer.
	hr = pTask->get_Actions(&pActionCollection);
	if (FAILED(hr))
	{
		printf("\nCannot get Task collection pointer: %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	//  Create the action, specifying that it is an executable action.
	IAction *pAction = NULL;
	hr = pActionCollection->Create(TASK_ACTION_EXEC, &pAction);
	pActionCollection->Release();
	if (FAILED(hr))
	{
		printf("\nCannot create the action: %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	IExecAction *pExecAction = NULL;
	//  QI for the executable task pointer.
	hr = pAction->QueryInterface(
		IID_IExecAction, (void**)&pExecAction);
	pAction->Release();
	if (FAILED(hr))
	{
		printf("\nQueryInterface call failed for IExecAction: %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	//  Set the path of the executable
	hr = pExecAction->put_Path(_bstr_t(wstrExecutablePath.c_str()));
	if (FAILED(hr))
	{
		printf("\nCannot put action path: %x", hr);
		pExecAction->Release();
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	//Set the arguments for the executable
	hr = pExecAction->put_Arguments(_bstr_t(wstrCommandArguments.c_str()));
	pExecAction->Release();
	if (FAILED(hr)){
		printf("\nCannot put arguments: %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	//  ------------------------------------------------------
	//  Save the task in the root folder.
	IRegisteredTask *pRegisteredTask = NULL;
	hr = pRootFolder->RegisterTaskDefinition(
		_bstr_t(wszTaskName),
		pTask,
		TASK_CREATE_OR_UPDATE,
		_variant_t(),
		_variant_t(),
		TASK_LOGON_INTERACTIVE_TOKEN,
		_variant_t(L""),
		&pRegisteredTask);
	if (FAILED(hr))
	{
		printf("\nError saving the Task : %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	printf("\n Success! Task successfully registered. ");

	//  Clean up.
	pRootFolder->Release();
	pTask->Release();
	pRegisteredTask->Release();
	CoUninitialize();
	return 0;
}

int _tmain(int argc, _TCHAR* argv[]){
	
	for (int i = 0; i < argc; i++){
		printf("\n Arg %d:\t%s", i, argv[i]);
	}

	// Get the windows directory and set the path to the program to
	// execute (currently just using notepad.exe)
	//Replace this with the path to this program
	//wchar_t *wchrExePathBuffer;
	//size_t bufferSize = 500;
	//_wdupenv_s(&wchrExePathBuffer, &bufferSize, L"WINDIR");
	//wstring wstrExecutablePath = (wstring)wchrExePathBuffer;
	//wstrExecutablePath += L"\\SYSTEM32\\NOTEPAD.EXE";
	wchar_t lpFileName[MAX_PATH];
	GetModuleFileName(NULL, lpFileName, MAX_PATH);
	wstring wstrExecutablePath = (wstring)lpFileName;
	wstrExecutablePath += L" /someoption";
	return scheduleNotepad(&wstrExecutablePath, &wstrCommandArguments);
}
