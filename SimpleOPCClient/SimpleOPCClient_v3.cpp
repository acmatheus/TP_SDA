// Simple OPC Client
//
// This is a modified version of the "Simple OPC Client" originally
// developed by Philippe Gras (CERN) for demonstrating the basic techniques
// involved in the development of an OPC DA client.
//
// The modifications are the introduction of two C++ classes to allow the
// the client to ask for callback notifications from the OPC server, and
// the corresponding introduction of a message comsumption loop in the
// main program to allow the client to process those notifications. The
// C++ classes implement the OPC DA 1.0 IAdviseSink and the OPC DA 2.0
// IOPCDataCallback client interfaces, and in turn were adapted from the
// KEPWARE큦  OPC client sample code. A few wrapper functions to initiate
// and to cancel the notifications were also developed.
//
// The original Simple OPC Client code can still be found (as of this date)
// in
//        http://pgras.home.cern.ch/pgras/OPCClientTutorial/
//
//
// Luiz T. S. Mendes - DELT/UFMG - 15 Sept 2011
// luizt at cpdee.ufmg.br
//

#include <atlbase.h>    // required for using the "_T" macro
#include <iostream>
#include <ObjIdl.h>
#include <stdio.h>
#include <thread>
#include <process.h>
#include <windows.h>

#include "opcda.h"
#include "opcerror.h"
#include "SimpleOPCClient_v3.h"
#include "SOCAdviseSink.h"
#include "SOCDataCallback.h"
#include "SOCWrapperFunctions.h"

using namespace std;

#define OPC_SERVER_NAME L"Matrikon.OPC.Simulation.1"
//#define OPC_SERVER_NAME L"TCCPUCMG.OPCDAServer20"
//#define OPC_SERVER_NAME L"ECA.OPCDAServer215"
//#define VT VT_R4


//#define REMOTE_SERVER_NAME L"your_path"

// Global variables

// The OPC DA Spec requires that some constants be registered in order to use
// them. The one below refers to the OPC DA 1.0 IDataObject interface.
UINT OPC_DATA_TIME = RegisterClipboardFormat (_T("OPCSTMFORMATDATATIME"));

//wchar_t ITEM_ID[]=L"Saw-toothed Waves.Real4";
wchar_t* ITEM_ID;
IOPCItemMgt* pIOPCItemMgt = NULL; //pointer to IOPCItemMgt interface
OPCHANDLE hServerItem[4];  // server handle to the item
OPCHANDLE hServerGroup; // server handle to the group
HANDLE gDoneEvent;

////////////////////////////////////////////////////
//////////// VARIAVEIS ECA.OPCDASERVER /////////////
////////////////////////////////////////////////////
/*wchar_t aux0[] = L"Leitura.Analog1";	//Leitura
wchar_t aux1[] = L"Leitura.Analog2";	//Leitura
wchar_t aux2[] = L"Leitura.Digit1";		//Leitura
wchar_t aux3[] = L"Leitura.Digit2";		//Leitura
wchar_t aux4[] = L"Leitura.Digit3";		//Leitura
wchar_t aux5[] = L"Leitura.Digit4";		//Leitura
wchar_t aux6[] = L"Leitura.Digit5";		//Leitura
wchar_t aux7[] = L"Escrita.Analog1";	//Escrita
wchar_t aux8[] = L"Escrita.Analog2";	//Escrita
wchar_t aux9[] = L"Escrita.Digit1";		//Escrita
wchar_t aux10[] = L"Escrita.Digit2";	//Escrita
wchar_t aux11[] = L"Escrita.Digit3";	//Escrita
wchar_t aux12[] = L"Escrita.Digit4";	//Escrita
wchar_t aux13[] = L"Escrita.Digit5";	//Escrita
wchar_t aux14[] = L"Escrita.Digit6";	//Escrita
wchar_t aux15[] = L"Escrita.Digit7";	//Escrita
wchar_t aux16[] = L"Escrita.Digit8";	//Escrita*/

wchar_t aux0[] = L"Random.Boolean";		//Leitura
wchar_t aux1[] = L"Random.Int1";		//Leitura
wchar_t aux2[] = L"Random.Boolean";		//Escrita
wchar_t aux3[] = L"Random.Int1";		//Escrita

//////////////////////////////////////////////////////////////////////
// Read the value of an item on an OPC server. 
//
void main()
{

	IOPCServer* pIOPCServer = NULL;   //pointer to IOPServer interface
	//IOPCItemMgt* pIOPCItemMgt = NULL; //pointer to IOPCItemMgt interface

	//OPCHANDLE hServerGroup; // server handle to the group
	//OPCHANDLE hServerItem;  // server handle to the item
	//OPCHANDLE hServerItem[4];  // server handle to the item

	int i;
	char buf[100];

	// Have to be done before using microsoft COM library:
	printf("Initializing the COM environment...\n");
	//CoInitialize(NULL);
	CoInitializeEx(NULL, COINIT_MULTITHREADED);

	// Let's instantiante the IOPCServer interface and get a pointer of it:
	printf("Intantiating the MATRIKON OPC Server for Simulation...\n");
	pIOPCServer = InstantiateServer(OPC_SERVER_NAME);
	
	// Add the OPC group the OPC server and get an handle to the IOPCItemMgt
	//interface:
	printf("Adding a group in the INACTIVE state for the moment...\n");
	AddTheGroup(pIOPCServer, pIOPCItemMgt, hServerGroup);
	
	// Add the OPC item. First we have to convert from wchar_t* to char*
	// in order to print the item name in the console.
    size_t m;
	int j;
	for (j = 0; j < 4; j++) {
		if (j == 0) {
			ITEM_ID = aux0;
			wcstombs_s(&m, buf, 100, ITEM_ID, _TRUNCATE);//<<<<<<<<<<<-----------add todos itens
			printf("Adding the item %s to the group...\n", buf);
			AddTheItem(pIOPCItemMgt, hServerItem[j], j, VT_BOOL);
		}
		else if (j == 1) {
			ITEM_ID = aux1;
			wcstombs_s(&m, buf, 100, ITEM_ID, _TRUNCATE);
			printf("Adding the item %s to the group...\n", buf);
			AddTheItem(pIOPCItemMgt, hServerItem[j], j, VT_I1);
		}
		else if (j == 2) {
			ITEM_ID = aux2;
			wcstombs_s(&m, buf, 100, ITEM_ID, _TRUNCATE);
			printf("Adding the item %s to the group...\n", buf);
			AddTheItem(pIOPCItemMgt, hServerItem[j], j, VT_BOOL);
		}
		else if (j == 3) {
			ITEM_ID = aux3;
			wcstombs_s(&m, buf, 100, ITEM_ID, _TRUNCATE);
			printf("Adding the item %s to the group...\n", buf);
			AddTheItem(pIOPCItemMgt, hServerItem[j], j, VT_I1);
		}
	}
	/*wcstombs_s(&m, buf, 100, ITEM_ID, _TRUNCATE);
	printf("Adding the item %s to the group...\n", buf);
    AddTheItem(pIOPCItemMgt, hServerItem);*/
	
	//Synchronous read of the device큦 item value.
	//VARIANT varValue; //to store the read value
	//VariantInit(&varValue);

	TimerQueue();

	/*printf ("Reading synchronously on main during 3 seconds...\n");
	for (i=0; i<3; i++) {

	  //ReadItem(pIOPCItemMgt, hServerItem, varValue);
	  //ReadItem(pIOPCItemMgt, hServerItem[0], varValue);
	  ReadItem(hServerItem[1], varValue);
	  // print the read value:
	  printf("mRead value: %d\n", varValue.intVal);

	  //ReadItem(pIOPCSyncIO, hServerItem[2], varValue);
	  ReadItem(hServerItem[2], varValue);
	  printf("mRead value: ");
	  printf(varValue.boolVal ? "true" : "false"); printf("\n");

	  // wait 1 second
	  Sleep(1200);
	}*/
	
	// Establish a callback asynchronous read by means of the old IAdviseSink()
	// (OPC DA 1.0) method. We first instantiate a new SOCAdviseSink object and
	// adjusts its reference count, and then call a wrapper function to
	// setup the callback.
	/*IDataObject* pIDataObject = NULL; //pointer to IDataObject interface
	DWORD tkAsyncConnection = 0;
	SOCAdviseSink* pSOCAdviseSink = new SOCAdviseSink ();
	pSOCAdviseSink->AddRef();
    printf("Setting up the IAdviseSink callback connection...\n");
    SetAdviseSink(pIOPCItemMgt, pSOCAdviseSink, pIDataObject, &tkAsyncConnection);

	// Change the group to the ACTIVE state so that we can receive the
	// server큦 callback notification
	printf("Changing the group state to ACTIVE...\n");
    SetGroupActive(pIOPCItemMgt); 

	// Enters a message pump in order to process the server큦 callback
	// notifications. This is needed because the CoInitialize() function
	// forces the COM threading model to STA (Single Threaded Apartment),
	// in which, according to the MSDN, "all method calls to a COM object
	// (...) are synchronized with the windows message queue for the
	// single-threaded apartment's thread." So, even being a console
	// application, the OPC client must process messages (which in this case
	// are only useless WM_USER [0x0400] messages) in order to process
	// incoming callbacks from a OPC server.
	//
	// A better alternative could be to use the CoInitializeEx() function,
	// which allows one to  specifiy the desired COM threading model;
	// in particular, calling
	//        CoInitializeEx(NULL, COINIT_MULTITHREADED)
	// sets the model to MTA (MultiThreaded Apartments) in which a message
	// loop is __not required__ since objects in this model are able to
	// receive method calls from other threads at any time. However, in the
	// MTA model the user is required to handle any aspects regarding
	// concurrency, since asynchronous, multiple calls to the object methods
	// can occur.
	//
	int bRet;
	MSG msg;
	DWORD ticks1, ticks2;
	ticks1 = GetTickCount();
	printf("Waiting for IAdviseSink callback notifications during 10 seconds...\n");
	do {
		bRet = GetMessage( &msg, NULL, 0, 0 );
		if (!bRet){
			printf ("Failed to get windows message! Error code = %d\n", GetLastError());
			exit(0);
		}
		TranslateMessage(&msg); // This call is not really needed ...
		DispatchMessage(&msg);  // ... but this one is!
        ticks2 = GetTickCount();
	}
	while ((ticks2 - ticks1) < 10000);

	// Cancel the callback and release its reference
	printf("Cancelling the IAdviseSink callback...\n");
    CancelAdviseSink(pIDataObject, tkAsyncConnection);
	pSOCAdviseSink->Release();
    
	// Establish a callback asynchronous read by means of the IOPCDaraCallback
	// (OPC DA 2.0) method. We first instantiate a new SOCDataCallback object and
	// adjusts its reference count, and then call a wrapper function to
	// setup the callback.
	IConnectionPoint* pIConnectionPoint = NULL; //pointer to IConnectionPoint Interface
	DWORD dwCookie = 0;
	SOCDataCallback* pSOCDataCallback = new SOCDataCallback();
	pSOCDataCallback->AddRef();

	printf("Setting up the IConnectionPoint callback connection...\n");
	SetDataCallback(pIOPCItemMgt, pSOCDataCallback, pIConnectionPoint, &dwCookie);

	// Change the group to the ACTIVE state so that we can receive the
	// server큦 callback notification
	printf("Changing the group state to ACTIVE...\n");
    SetGroupActive(pIOPCItemMgt); 

	// Enter again a message pump in order to process the server큦 callback
	// notifications, for the same reason explained before.
		
	ticks1 = GetTickCount();
	printf("Waiting for IOPCDataCallback notifications during 10 seconds...\n");
	do {
		bRet = GetMessage( &msg, NULL, 0, 0 );
		if (!bRet){
			printf ("Failed to get windows message! Error code = %d\n", GetLastError());
			exit(0);
		}
		TranslateMessage(&msg); // This call is not really needed ...
		DispatchMessage(&msg);  // ... but this one is!
        ticks2 = GetTickCount();
	}
	while ((ticks2 - ticks1) < 10000);

	// Cancel the callback and release its reference
	printf("Cancelling the IOPCDataCallback notifications...\n");
    CancelDataCallback(pIConnectionPoint, dwCookie);
	pIConnectionPoint->Release();
	pSOCDataCallback->Release();*/
system("pause");

	// Remove the OPC item:
	printf("Removing the OPC item...\n");
	//RemoveItem(pIOPCItemMgt, hServerItem);
	RemoveItem(pIOPCItemMgt, hServerItem[0]);

	// Remove the OPC group:
	printf("Removing the OPC group object...\n");
    pIOPCItemMgt->Release();
	RemoveGroup(pIOPCServer, hServerGroup);

	// release the interface references:
	printf("Removing the OPC server object...\n");
	pIOPCServer->Release();

	//close the COM library:
	printf ("Releasing the COM environment...\n");
	CoUninitialize();
}

void TimerQueue() {

	HANDLE hTimer = NULL;
	HANDLE hTimerQueue = NULL;
	int arg = 123;

	// Use an event object to track the TimerRoutine execution
	gDoneEvent = CreateEvent(NULL, TRUE, FALSE, NULL);
	if (NULL == gDoneEvent)
	{
		printf("CreateEvent failed (%d)\n", GetLastError());
		//return 1;
	}

	// Create the timer queue.
	hTimerQueue = CreateTimerQueue();
	if (NULL == hTimerQueue)
	{
		printf("CreateTimerQueue failed (%d)\n", GetLastError());
		//return 2;
	}

	// Set a timer to call the timer routine in 0.1 seconds.
	if (!CreateTimerQueueTimer(&hTimer, hTimerQueue,(WAITORTIMERCALLBACK)LeituraSinc, &arg, 0, 500, 0)){
		printf("CreateTimerQueueTimer failed (%d)\n", GetLastError());
		//return 3;
	}

}

//unsigned __stdcall LeituraSinc(void *a) {
VOID CALLBACK LeituraSinc(PVOID lpParam, BOOLEAN TimerOrWaitFired){
	VARIANT varValue; //to store the read value
	VariantInit(&varValue);
	
	//IOPCSyncIO* pIOPCSyncIO;
	//pIOPCItemMgt->QueryInterface(__uuidof(pIOPCSyncIO), (void**)&pIOPCSyncIO);

		ReadItem(pIOPCItemMgt,hServerItem[0], varValue);
		printf("Read value: ");
		printf(varValue.boolVal ? "true" : "false"); printf("\n");

		ReadItem(pIOPCItemMgt,hServerItem[1], varValue);
		printf("Read value: %d\n\n", varValue.intVal);

	//pIOPCSyncIO->Release();

	//system("pause");
	//_endthreadex(0);

	//return 0;
}

////////////////////////////////////////////////////////////////////
// Instantiate the IOPCServer interface of the OPCServer
// having the name ServerName. Return a pointer to this interface
//
IOPCServer* InstantiateServer(wchar_t ServerName[])
{
	CLSID CLSID_OPCServer;
	HRESULT hr;

	// get the CLSID from the OPC Server Name:
	hr = CLSIDFromString(ServerName, &CLSID_OPCServer);
	_ASSERT(!FAILED(hr));


	//queue of the class instances to create
	LONG cmq = 1; // nbr of class instance to create.
	MULTI_QI queue[1] =
		{{&IID_IOPCServer,
		NULL,
		0}};

	//Server info:
	//COSERVERINFO CoServerInfo =
    //{
	//	/*dwReserved1*/ 0,
	//	/*pwszName*/ REMOTE_SERVER_NAME,
	//	/*COAUTHINFO*/  NULL,
	//	/*dwReserved2*/ 0
    //}; 

	// create an instance of the IOPCServer
	hr = CoCreateInstanceEx(CLSID_OPCServer, NULL, CLSCTX_SERVER,
		/*&CoServerInfo*/NULL, cmq, queue);
	_ASSERT(!hr);

	// return a pointer to the IOPCServer interface:
	return(IOPCServer*) queue[0].pItf;
}


/////////////////////////////////////////////////////////////////////
// Add group "Group1" to the Server whose IOPCServer interface
// is pointed by pIOPCServer. 
// Returns a pointer to the IOPCItemMgt interface of the added group
// and a server opc handle to the added group.
//
void AddTheGroup(IOPCServer* pIOPCServer, IOPCItemMgt* &pIOPCItemMgt, 
				 OPCHANDLE& hServerGroup)
{
	DWORD dwUpdateRate = 0;
	OPCHANDLE hClientGroup = 0;

	// Add an OPC group and get a pointer to the IUnknown I/F:
    HRESULT hr = pIOPCServer->AddGroup(/*szName*/ L"Group1",
		/*bActive*/ FALSE,
		/*dwRequestedUpdateRate*/ 1000,
		/*hClientGroup*/ hClientGroup,
		/*pTimeBias*/ 0,
		/*pPercentDeadband*/ 0,
		/*dwLCID*/0,
		/*phServerGroup*/&hServerGroup,
		&dwUpdateRate,
		/*riid*/ IID_IOPCItemMgt,
		/*ppUnk*/ (IUnknown**) &pIOPCItemMgt);
	_ASSERT(!FAILED(hr));
}



//////////////////////////////////////////////////////////////////
// Add the Item ITEM_ID to the group whose IOPCItemMgt interface
// is pointed by pIOPCItemMgt pointer. Return a server opc handle
// to the item.
 
void AddTheItem(IOPCItemMgt* pIOPCItemMgt, OPCHANDLE& hServerItem, int id, int datatype)
{
	HRESULT hr;

	// Array of items to add:
	OPCITEMDEF ItemArray[1] =
	{{
	/*szAccessPath*/ L"",
	/*szItemID*/ ITEM_ID,
	/*bActive*/ TRUE,
	/*hClient*/ id,
	/*dwBlobSize*/ 0,
	/*pBlob*/ NULL,
	/*vtRequestedDataType*/ datatype,
	/*wReserved*/0
	}};

	//Add Result:
	OPCITEMRESULT* pAddResult=NULL;
	HRESULT* pErrors = NULL;
	
	// Add an Item to the previous Group:
	hr = pIOPCItemMgt->AddItems(1, ItemArray, &pAddResult, &pErrors);
	if (hr != S_OK){
		printf("Failed call to AddItems function. Error code = %x\n", hr);
		exit(0);
	}
	
	// Server handle for the added item:
	hServerItem = pAddResult[0].hServer;

	// release memory allocated by the server:
	CoTaskMemFree(pAddResult->pBlob);

	CoTaskMemFree(pAddResult);
	pAddResult = NULL;

	CoTaskMemFree(pErrors);
	pErrors = NULL;
}

///////////////////////////////////////////////////////////////////////////////
// Read from device the value of the item having the "hServerItem" server 
// handle and belonging to the group whose one interface is pointed by
// pGroupIUnknown. The value is put in varValue. 
//
void ReadItem(IUnknown* pGroupIUnknown, OPCHANDLE hServerItem, VARIANT& varValue)
{
	// value of the item:
	OPCITEMSTATE* pValue = NULL;

	//get a pointer to the IOPCSyncIOInterface:
	IOPCSyncIO* pIOPCSyncIO;
	pGroupIUnknown->QueryInterface(__uuidof(pIOPCSyncIO), (void**) &pIOPCSyncIO);

	// read the item value from the device:
	HRESULT* pErrors = NULL; //to store error code(s)
	HRESULT hr = pIOPCSyncIO->Read(OPC_DS_DEVICE, 1, &hServerItem, &pValue, &pErrors);
	_ASSERT(!hr);
	_ASSERT(pValue!=NULL);

	varValue = pValue[0].vDataValue;

	//Release memory allocated by the OPC server:
	CoTaskMemFree(pErrors);
	pErrors = NULL;

	CoTaskMemFree(pValue);
	pValue = NULL;

	// release the reference to the IOPCSyncIO interface:
	pIOPCSyncIO->Release();
}

///////////////////////////////////////////////////////////////////////////
// Remove the item whose server handle is hServerItem from the group
// whose IOPCItemMgt interface is pointed by pIOPCItemMgt
//
void RemoveItem(IOPCItemMgt* pIOPCItemMgt, OPCHANDLE hServerItem)
{
	// server handle of items to remove:
	OPCHANDLE hServerArray[1];
	hServerArray[0] = hServerItem;
	
	//Remove the item:
	HRESULT* pErrors; // to store error code(s)
	HRESULT hr = pIOPCItemMgt->RemoveItems(1, hServerArray, &pErrors);
	_ASSERT(!hr);

	//release memory allocated by the server:
	CoTaskMemFree(pErrors);
	pErrors = NULL;
}

////////////////////////////////////////////////////////////////////////
// Remove the Group whose server handle is hServerGroup from the server
// whose IOPCServer interface is pointed by pIOPCServer
//
void RemoveGroup (IOPCServer* pIOPCServer, OPCHANDLE hServerGroup)
{
	// Remove the group:
	HRESULT hr = pIOPCServer->RemoveGroup(hServerGroup, FALSE);
	if (hr != S_OK){
		if (hr == OPC_S_INUSE)
			printf ("Failed to remove OPC group: object still has references to it.\n");
		else printf ("Failed to remove OPC group. Error code = %x\n", hr);
		exit(0);
	}
}
