#include "CaptureVideo.h"

#include <regex>
#include <cfgmgr32.h>
#include <setupapi.h>
#include <atlconv.h>

#pragma comment (lib, "setupapi.lib")

using namespace std;

SampleGrabberCallback g_sampleGrabberCB;  //CallBack

CaptureVideo::CaptureVideo()
{
	//COM Library Initialize
	if (FAILED(CoInitialize(NULL)))
	{
		Msg(m_App,TEXT("CoInitialize Failed!\r\n"));
		return;
	}
	//initialize member variable
	m_nCaptureDeviceNumber = 0;
	m_pDevFilter = NULL;
	m_pCaptureGB = NULL;
	m_pGraphBuilder = NULL;
	m_pMediaControl = NULL;
	m_pMediaEvent = NULL;
	m_pSampGrabber = NULL;
	m_pKsControl = NULL;
	m_pVideoWindow = NULL;
	m_App = NULL;
	m_bGetOneShot = FALSE;
	m_bConnect = FALSE;
	InitializeEnv();
}

CaptureVideo::~CaptureVideo()
{
	CloseInterface();
	CoUninitialize();
}

HRESULT CaptureVideo::EnumAllDevices(HWND hCombox)
{
	if (!hCombox)
		return S_FALSE;
	ICreateDevEnum *pDevEnum;
	IEnumMoniker   *pEnumMon;
	IMoniker	   *pMoniker;
	HRESULT hr = CoCreateInstance(CLSID_SystemDeviceEnum,NULL,CLSCTX_INPROC_SERVER,
			IID_ICreateDevEnum,(LPVOID*)&pDevEnum);
	if (SUCCEEDED(hr))
	{
		hr = pDevEnum->CreateClassEnumerator(CLSID_VideoInputDeviceCategory,&pEnumMon, 0);
		if (hr == S_FALSE)
		{
			hr = VFW_E_NOT_FOUND;
			return hr;
		}
		pEnumMon->Reset();
		ULONG cFetched;
		while(hr=pEnumMon->Next(1,&pMoniker,&cFetched),hr == S_OK)
		{
			IPropertyBag *pProBag;
			hr = pMoniker->BindToStorage(0,0,IID_IPropertyBag,(LPVOID*)&pProBag);
			if (SUCCEEDED(hr))
			{
				VARIANT varTemp;
				varTemp.vt = VT_BSTR;
				hr = pProBag->Read(L"FriendlyName",&varTemp,NULL);
				if (SUCCEEDED(hr))
				{
					//int nStrSize = WideCharToMultiByte(CP_ACP,0,varTemp.bstrVal,-1,0,0,NULL,FALSE);
					//char *strName = new char[nStrSize];
					//WideCharToMultiByte(CP_ACP,0,varTemp.bstrVal,-1,strName,nStrSize,NULL,FALSE);
					//m_vecCapDeviceName.push_back(string(strName)); //push the device name to vector
					
					StringCchCopy(m_pCapDeviceName[m_nCaptureDeviceNumber],MAX_PATH,varTemp.bstrVal);
					m_nCaptureDeviceNumber++; //Device number add 1
					::SendMessage(hCombox, CB_ADDSTRING, 0,(LPARAM)varTemp.bstrVal);
					SysFreeString(varTemp.bstrVal);
				}
				pProBag->Release();
			}
			pMoniker->Release();
		}
		pEnumMon->Release();
	}
	return hr;
}

HRESULT CaptureVideo::InitializeEnv()
{
	HRESULT hr;

	//Create the filter graph
	hr = CoCreateInstance(CLSID_FilterGraph,NULL,CLSCTX_INPROC_SERVER,
						  IID_IGraphBuilder,(LPVOID*)&m_pGraphBuilder);
	if(FAILED(hr))
		return hr;
	
	//Create the capture graph builder
	hr = CoCreateInstance(CLSID_CaptureGraphBuilder2,NULL,CLSCTX_INPROC_SERVER,
						  IID_ICaptureGraphBuilder2,(LPVOID*)&m_pCaptureGB);
	if(FAILED(hr))
		return hr;

	//Obtain interfaces for media control and Video Window
	hr = m_pGraphBuilder->QueryInterface(IID_IMediaControl,(LPVOID*)&m_pMediaControl);
	if(FAILED(hr))
		return hr;

	hr = m_pGraphBuilder->QueryInterface(IID_IVideoWindow,(LPVOID*)&m_pVideoWindow);
	if(FAILED(hr))
		return hr;

	hr = m_pGraphBuilder->QueryInterface(IID_IMediaEventEx,(LPVOID*)&m_pMediaEvent);
	if(FAILED(hr))
		return hr;

	hr = m_pMediaEvent->SetNotifyWindow((OAHWND)m_App,WM_GRAPHNOTIFY,0);

	if(FAILED(hr))
		return hr;
	m_pCaptureGB->SetFiltergraph(m_pGraphBuilder);
	if(FAILED(hr))
		return hr;
	return hr;
}

void CaptureVideo::CloseInterface()
{
	m_bGetOneShot = FALSE;

	if (m_pMediaControl)
		m_pMediaControl->Stop();
	if(m_pVideoWindow)
	{
		m_pVideoWindow->get_Visible(FALSE);
		m_pVideoWindow->put_Owner(NULL);
	}

	if(m_pMediaEvent)
		m_pMediaEvent->SetNotifyWindow(NULL,WM_GRAPHNOTIFY,0);

	//release interface
	ReleaseInterface(m_pKsControl);
	ReleaseInterface(m_pDevFilter);
	ReleaseInterface(m_pCaptureGB);
	ReleaseInterface(m_pGraphBuilder);
	ReleaseInterface(m_pMediaControl);
	ReleaseInterface(m_pMediaEvent);
	ReleaseInterface(m_pSampGrabber);
	ReleaseInterface(m_pVideoWindow);
}

void DisplayDeviceInformation(IEnumMoniker *pEnum)
{
	IMoniker *pMoniker = NULL;
	while (pEnum->Next(1, &pMoniker, NULL) == S_OK)
	{
		IPropertyBag *pPropBag;
		HRESULT hr = pMoniker->BindToStorage(0, 0, IID_PPV_ARGS(&pPropBag));
		if (FAILED(hr))
		{
			pMoniker->Release();
			continue;
		} 
		VARIANT var;
		VariantInit(&var);
		// Get description or friendly name.
		hr = pPropBag->Read(L"Description", &var, 0);
		if (FAILED(hr))
		{
			hr = pPropBag->Read(L"FriendlyName", &var, 0);
		} 
		if(SUCCEEDED(hr))
		{
			logout("%S\n", var.bstrVal);
			VariantClear(&var);
		}
		hr = pPropBag->Write(L"FriendlyName", &var);
		// WaveInID applies only to audio capture devices.
		hr = pPropBag->Read(L"WaveInID", &var, 0);
		if (SUCCEEDED(hr))
		{
			logout("WaveIn ID: %d\n", var.lVal);
			VariantClear(&var);
		} 
		hr = pPropBag->Read(L"DevicePath", &var, 0);
		if (SUCCEEDED(hr))
		{
			// The device path is not intended for display.
			logout("Device path: %S\n", var.bstrVal);
			VariantClear(&var);
		} 
		pPropBag->Release();
		pMoniker->Release();
	}
}

#define GUID_CAMERA_STRING L"{65e8773d-8f56-11d0-a3b9-00a0c9223196}"
static HRESULT GetDeviceInstance(string szDevPath, DEVINST *pDeviceInstance)
{
    GUID guid;
    HDEVINFO devInfo;
    DWORD devIndex;
    SP_DEVINFO_DATA devInfoData;

    CLSIDFromString(GUID_CAMERA_STRING, &guid);
    devInfo = SetupDiGetClassDevs(&guid, NULL, NULL, DIGCF_PRESENT | DIGCF_DEVICEINTERFACE);
    if (devInfo == INVALID_HANDLE_VALUE)
    {
        logout("SetupDiGetClassDevs failed! errorno: %d\n", GetLastError());
        return E_FAIL;
    }

    devInfoData.cbSize = sizeof(SP_DEVINFO_DATA);
    devIndex = 0;
    while (SetupDiEnumDeviceInfo(devInfo, devIndex, &devInfoData))
    {
        WCHAR szDeviceInstanceId[MAX_DEVICE_ID_LEN];

        memset(szDeviceInstanceId, 0, MAX_DEVICE_ID_LEN);
        SetupDiGetDeviceInstanceIdW(devInfo, &devInfoData, szDeviceInstanceId, MAX_PATH, 0);
        USES_CONVERSION;
        char *szDeviceInstId = W2A(szDeviceInstanceId);
        transform(szDevPath.begin(), szDevPath.end(), szDevPath.begin(), ::toupper);
        if (szDevPath.compare(szDeviceInstId) == 0)
        {
            *pDeviceInstance = devInfoData.DevInst;
            logout("Get DeviceInstance[%d]: %s\n", devIndex, szDeviceInstId);
            break;
        }
        devIndex++;
    }

    return S_OK;
}

HRESULT CaptureVideo::RestartDevice()
{
    HRESULT hr;
    CONFIGRET result;
    DEVINST childDevInstance;
    DEVINST parentDevInstance;

    hr = GetDeviceInstance(m_szDevicePath, &childDevInstance);
    if (FAILED(hr))
    {
        logout("GetDeviceInstance failed!\n");
        return hr;
    }

    /* Must restart parent device instance, or device will not reset all trbs. */
    result = CM_Get_Parent(&parentDevInstance, childDevInstance, 0);
    if (result != CR_SUCCESS) {
        logout("CM_Get_Parent failed! result: %d\n", result);
        return E_FAIL;
    }

    result = CM_Disable_DevNode(parentDevInstance, 0);
    if (result != CR_SUCCESS) {
        logout("CM_Disable_DevNode failed! result: %d\n", result);
        return E_FAIL;
    }

    result = CM_Enable_DevNode(parentDevInstance, 0);
    if (result != CR_SUCCESS) {
        logout("CM_Disable_DevNode failed! result: %d\n", result);
        return E_FAIL;
    }

    return S_OK;
}

HRESULT CaptureVideo::BindFilter(int deviceID, IBaseFilter **pBaseFilter)
{
	ICreateDevEnum *pDevEnum;
	IEnumMoniker   *pEnumMon;
	IMoniker	   *pMoniker;
	HRESULT hr = CoCreateInstance(CLSID_SystemDeviceEnum,NULL,CLSCTX_INPROC_SERVER,
		IID_ICreateDevEnum,(LPVOID*)&pDevEnum);
	if (SUCCEEDED(hr))
	{
		hr = pDevEnum->CreateClassEnumerator(CLSID_VideoInputDeviceCategory,&pEnumMon, 0);
		if (hr == S_FALSE)
		{
			hr = VFW_E_NOT_FOUND;
			return hr;
		}
		DisplayDeviceInformation(pEnumMon);
		pEnumMon->Reset();
		ULONG cFetched;
		int index = 0;
		while(hr=pEnumMon->Next(1,&pMoniker,&cFetched),hr == S_OK, index<= deviceID)
		{
			IPropertyBag *pProBag;
			hr = pMoniker->BindToStorage(0,0,IID_IPropertyBag,(LPVOID*)&pProBag);
			if (SUCCEEDED(hr))
			{
				if (index == deviceID)
				{
					pMoniker->BindToObject(0,0,IID_IBaseFilter,(LPVOID*)pBaseFilter);
                    
                    VARIANT var;
                    string szDevPath;

                    VariantInit(&var);
                    hr = pProBag->Read(L"DevicePath", &var, 0);
                    if (SUCCEEDED(hr))
                    {
                        USES_CONVERSION;
                        szDevPath = W2A(var.bstrVal);
                        // Convert format
                        while (szDevPath[0] != 'u')
                            szDevPath.erase(0, 1);
                        int k = 0;
                        while (szDevPath[k] != '{')
                        {
                            if (szDevPath[k] == '#')
                            {
                                szDevPath.erase(k, 1);
                                szDevPath.insert(k, "\\");
                            }
                            k++;
                        }
                        szDevPath.erase(k - 1);
                    }
                    VariantClear(&var);

                    m_szDevicePath = szDevPath;
				}
			}
			pMoniker->Release();
			index++;
		}
		pEnumMon->Release();
	}
	return hr;
}

HRESULT CaptureVideo::SetupVideoWindow(LONG nLeft, LONG nTop, LONG nWidth, LONG nHeight)
{
	HRESULT hr;
	hr = m_pVideoWindow->put_Owner((OAHWND)m_App);
	if (FAILED(hr))
		return hr;

	hr = m_pVideoWindow->put_WindowStyle(WS_CHILD | WS_CLIPCHILDREN);
	if(FAILED(hr))
		return hr;

	ResizeVideoWindow(nLeft,nTop,nWidth,nHeight);

	hr = m_pVideoWindow->put_Visible(OATRUE);
	if(FAILED(hr))
		return hr;

	return hr;
}

void CaptureVideo::ResizeVideoWindow(LONG nLeft, LONG nTop, LONG nWidth, LONG nHeight)
{
	if(m_pVideoWindow)
	{
		m_pVideoWindow->SetWindowPosition(nLeft,nTop,nWidth,nHeight);
	}
}

HRESULT CaptureVideo::OpenDevice(int deviceID, LONG nLeft, LONG nTop, LONG nWide, LONG nHeight)
{
	HRESULT hr;
	IBaseFilter *pSampleGrabberFilter;
	if (m_bConnect)
	{
		CloseInterface();
		InitializeEnv();
	}

	hr = CoCreateInstance(CLSID_SampleGrabber,NULL,CLSCTX_INPROC_SERVER,
						  IID_IBaseFilter, (LPVOID*)&pSampleGrabberFilter);
	if(FAILED(hr))
		return hr;
	//bind device filter
	hr = BindFilter(deviceID,&m_pDevFilter);
	if (FAILED(hr))
		return hr;

	hr = m_pGraphBuilder->AddFilter(m_pDevFilter,L"Video Filter");
	if (FAILED(hr))
		return hr;

	hr = m_pGraphBuilder->AddFilter(pSampleGrabberFilter,L"Sample Grabber");
	if (FAILED(hr))
		return hr;

	hr = pSampleGrabberFilter->QueryInterface(IID_ISampleGrabber,(LPVOID*)&m_pSampGrabber);
	if(FAILED(hr))
		return hr;

	//set media type
	AM_MEDIA_TYPE mediaType;
	ZeroMemory(&mediaType,sizeof(AM_MEDIA_TYPE));

#if 0
	//Find the current bit depth
	HDC hdc=GetDC(NULL);
	int iBitDepth=GetDeviceCaps(hdc, BITSPIXEL);
	g_sampleGrabberCB.m_iBitCount = iBitDepth;
	ReleaseDC(NULL,hdc);
	//Set the media type
	mediaType.majortype = MEDIATYPE_Video;
	switch(iBitDepth)
	{
	case  8:
		mediaType.subtype=MEDIASUBTYPE_RGB8;
		break;
	case 16:
		mediaType.subtype=MEDIASUBTYPE_RGB555;
		break;
	case 24:
		mediaType.subtype=MEDIASUBTYPE_RGB24;
		break;
	case 32:
		mediaType.subtype=MEDIASUBTYPE_RGB32;
		break;
	default:
		return E_FAIL;
	}
	mediaType.formattype = FORMAT_VideoInfo;
	hr = m_pSampGrabber->SetMediaType(&mediaType);

	hr = m_pCaptureGB->RenderStream(&PIN_CATEGORY_PREVIEW,&MEDIATYPE_Video,
		m_pDevFilter,pSampleGrabberFilter,NULL);
	if(FAILED(hr))
		return hr;
#else

#if 1
	//Set the media type
	mediaType.majortype = MEDIATYPE_Video;
	mediaType.subtype = MEDIASUBTYPE_YUY2;
	mediaType.formattype = FORMAT_VideoInfo;
	hr = m_pSampGrabber->SetMediaType(&mediaType);

	hr = m_pCaptureGB->RenderStream(&PIN_CATEGORY_CAPTURE, &MEDIATYPE_Video,
		m_pDevFilter, pSampleGrabberFilter, NULL);
	if (FAILED(hr))
	{
		logout("RenderStream failed. hr=%d\n", hr);
		return hr;
	}
#else
	hr = m_pCaptureGB->RenderStream(NULL, NULL,
		m_pDevFilter, pSampleGrabberFilter, NULL);
	if (FAILED(hr))
		return hr;
#endif

#endif

	hr = m_pSampGrabber->GetConnectedMediaType(&mediaType);
	if(FAILED(hr))
		return hr;

	if (mediaType.subtype == MEDIASUBTYPE_YUY2)
	{
		logout("mediaType.subtype : MEDIASUBTYPE_YUY2\n");
	}
	VIDEOINFOHEADER * vih = (VIDEOINFOHEADER*) mediaType.pbFormat;
	logout("mediaType.pbFormat : %d x %d\n", vih->bmiHeader.biWidth, vih->bmiHeader.biHeight);
	g_sampleGrabberCB.m_lWidth = vih->bmiHeader.biWidth;
	g_sampleGrabberCB.m_lHeight = vih->bmiHeader.biHeight;
	// Configure the Sample Grabber
	hr = m_pSampGrabber->SetOneShot(FALSE);
	if (FAILED(hr))
		return hr;
	hr = m_pSampGrabber->SetBufferSamples(TRUE);
	if (FAILED(hr))
		return hr;
	// 1 = Use the BufferCB callback method.
	hr = m_pSampGrabber->SetCallback(&g_sampleGrabberCB,1);

	//set capture video Window
	SetupVideoWindow(nLeft,nTop,nWide,nHeight);
	hr = m_pMediaControl->Run();
	if(FAILED(hr))
		return hr;

	if (mediaType.cbFormat != 0)
	{
		CoTaskMemFree((PVOID)mediaType.pbFormat);
		mediaType.cbFormat = 0;
		mediaType.pbFormat = NULL;
	}
	if (mediaType.pUnk != NULL)
	{
		mediaType.pUnk->Release();
		mediaType.pUnk = NULL;
	}
	m_bConnect = TRUE;
	return hr;

}

void CaptureVideo::Stop(void)
{
    if (m_pMediaControl)
        m_pMediaControl->Stop();
}

void CaptureVideo::GrabOneFrame(BOOL bGrab)
{
	m_bGetOneShot = bGrab;
	g_sampleGrabberCB.m_bGetPicture = bGrab;
}

static const GUID GUID_EXTENSION_UNIT_DESCRIPTOR =
{ 0x9a1e7291, 0x6843, 0x4683, { 0x6d, 0x92, 0x39, 0xbc, 0x79, 0x06, 0xee, 0x49 } };
#define MSG_BUF_SIZE    512
void CaptureVideo::TestMsg(void)
{
	HRESULT hr = E_FAIL;
	DWORD dwError = 0;
	IKsTopologyInfo *pKsTopologyInfo;

	// pUnkOuter is the unknown associated with the base filter
	hr = m_pDevFilter->QueryInterface(__uuidof(IKsTopologyInfo),
		(void **)&pKsTopologyInfo);
	if (!SUCCEEDED(hr))
	{
		logout("Unable to obtain IKsTopologyInfo %x\n", hr);
		goto errExit;
	}

	DWORD dwExtensionNode;
	hr = FindExtensionNode(pKsTopologyInfo,
		GUID_EXTENSION_UNIT_DESCRIPTOR,
		&dwExtensionNode);
	if (FAILED(hr))
	{
		logout("Unable to find extension node : %x\n", hr);
		goto errExit;
	}


	KSP_NODE ExtensionProp;
	unsigned long ulBytesReturned = 0;
	BYTE pValue[MSG_BUF_SIZE];

	hr = m_pDevFilter->QueryInterface(__uuidof(IKsControl),
		(void **)&m_pKsControl);
	if (FAILED(hr))
	{
		logout("Failed to register for auto-update event : %x\n", hr);
		goto errExit;
	}

	ExtensionProp.Property.Set = GUID_EXTENSION_UNIT_DESCRIPTOR;
	ExtensionProp.Property.Id = 1;
	ExtensionProp.Property.Flags = KSPROPERTY_TYPE_SET | KSPROPERTY_TYPE_TOPOLOGY;
	ExtensionProp.NodeId = dwExtensionNode;
	ExtensionProp.Reserved = 0;

	memset(pValue, 0xAA, MSG_BUF_SIZE);
	pValue[0] = 0xF0;
	pValue[4] = 0xDE;
	pValue[8] = 0xBC;
	pValue[12] = 0x9A;
	pValue[16] = 0x78;
	pValue[20] = 0x56;
	pValue[24] = 0x34;
	pValue[28] = 0x12;
	hr = m_pKsControl->KsProperty((PKSPROPERTY)&ExtensionProp, sizeof(ExtensionProp), pValue, 60, &ulBytesReturned);
	if (FAILED(hr))
	{
		logout("Failed to KsProperty for IKsControl : %x\n", hr);
	}
	else
	{
		logout("KSPROPERTY_TYPE_SET ulBytesReturned : %u\n", ulBytesReturned);
	}

	ExtensionProp.Property.Set = GUID_EXTENSION_UNIT_DESCRIPTOR;
	ExtensionProp.Property.Id = 1;
	ExtensionProp.Property.Flags = KSPROPERTY_TYPE_GET | KSPROPERTY_TYPE_TOPOLOGY;
	ExtensionProp.NodeId = dwExtensionNode;
	ExtensionProp.Reserved = 0;

	memset(pValue, 0x00, MSG_BUF_SIZE);
	hr = m_pKsControl->KsProperty((PKSPROPERTY)&ExtensionProp, sizeof(ExtensionProp), pValue, 32, &ulBytesReturned);
	if (FAILED(hr))
	{
		logout("Failed to KsProperty for IKsControl : %x\n", hr);
	}
	else
	{
		logout("KSPROPERTY_TYPE_GET ulBytesReturned : %u\n", ulBytesReturned);
		for (ULONG j = 0; j < ulBytesReturned; j++)
		{
			logout("%02X \n", pValue[j]);
		}
		logout("\n");
	}

    ExtensionProp.Property.Set = GUID_EXTENSION_UNIT_DESCRIPTOR;
    ExtensionProp.Property.Id = 1;
    ExtensionProp.Property.Flags = KSPROPERTY_TYPE_DEFAULTVALUES | KSPROPERTY_TYPE_TOPOLOGY;
    ExtensionProp.NodeId = dwExtensionNode;
    ExtensionProp.Reserved = 0;

    memset(pValue, 0x00, MSG_BUF_SIZE);
    hr = m_pKsControl->KsProperty((PKSPROPERTY)&ExtensionProp, sizeof(ExtensionProp), pValue, MSG_BUF_SIZE, &ulBytesReturned);
    if (FAILED(hr))
    {
        logout("Failed to KsProperty for IKsControl : 0x%x\n", hr);
    }
    else
    {
        logout("KSPROPERTY_TYPE_DEFAULTVALUES ulBytesReturned : %u\n", ulBytesReturned);
        for (ULONG j = 0; j < ulBytesReturned; j++)
        {
            logout("%02X \n", pValue[j]);
        }
        logout("\n");
    }


	



	HANDLE hEvent = CreateEvent(NULL, FALSE, FALSE, NULL);
	if (!hEvent)
	{
		logout("CreateEvent failed\n");
		goto errExit;
	}
	KSEVENT Event;
	Event.Set = KSEVENTSETID_VIDCAPNotify;
	Event.Id = KSEVENT_VIDCAP_AUTO_UPDATE;
	Event.Flags = KSEVENT_TYPE_ENABLE;

	KSEVENTDATA EventData;
	EventData.NotificationType = KSEVENTF_EVENT_HANDLE;
	EventData.EventHandle.Event = hEvent;
	EventData.EventHandle.Reserved[0] = 0;
	EventData.EventHandle.Reserved[1] = 0;

	hr = m_pKsControl->KsEvent(
		&Event,
		sizeof(KSEVENT),
		&EventData,
		sizeof(KSEVENTDATA),
		&ulBytesReturned);
	if (FAILED(hr))
	{
		logout("Failed to register for auto-update event : %x\n", hr);
		goto errExit;
	}

	// Wait for event for 5 seconds 
	dwError = WaitForSingleObject(hEvent, 5000);

	// cancel further notifications
	hr = m_pKsControl->KsEvent(
		NULL,
		0,
		&EventData,
		sizeof(KSEVENTDATA),
		&ulBytesReturned);
	if (FAILED(hr))  logout("Cancel event returns : %x\n", hr);

	if ((dwError == WAIT_FAILED) ||
		(dwError == WAIT_ABANDONED) ||
		(dwError == WAIT_TIMEOUT))
	{
		logout("Wait failed : %d\n", dwError);
		goto errExit;
	}

errExit:
	logout("Wait returned : %d\n", dwError);
	ReleaseInterface(m_pKsControl);
	ReleaseInterface(pKsTopologyInfo);	

	// handle the autoupdate event..
}

HRESULT CaptureVideo::FindExtensionNode(IKsTopologyInfo *pKsTopologyInfo, GUID guid, DWORD *node)
{
	HRESULT hr = E_FAIL;
	DWORD dwNumNodes = 0;
	GUID guidNodeType;
	IKsControl *pKsControl = NULL;
	ULONG ulBytesReturned = 0;
	KSP_NODE ExtensionProp;
	if (!pKsTopologyInfo || !node)
		return E_POINTER;
	// Retrieve the number of nodes in the filter
	hr = pKsTopologyInfo->get_NumNodes(&dwNumNodes);
	logout("pKsTopologyInfo get_NumNodes ret=%d, dwNumNodes=%u\n", hr, dwNumNodes);
	if (!SUCCEEDED(hr))
		return hr;
	if (dwNumNodes == 0)
		return E_FAIL;
	// Find the extension unit node that corresponds to the given GUID
	logout("Find %08X-%04X-%04X-", guid.Data1, guid.Data2, guid.Data3);
	for (int k = 0; k < 8; k++)
	{
		logout("%X", guid.Data4[k]);
	}
	logout("\n");
	for (unsigned int i = 0; i < dwNumNodes; i++)
	{
		hr = E_FAIL;
		pKsTopologyInfo->get_NodeType(i, &guidNodeType);
		WCHAR buf[32];
		DWORD len;
		pKsTopologyInfo->get_NodeName(i, buf, 32, &len);
		int iSize;
		char* pszMultiByte;
		iSize = WideCharToMultiByte(CP_ACP, 0, buf, -1, NULL, 0, NULL, NULL);
		pszMultiByte = (char*)malloc(iSize * sizeof(char));
		WideCharToMultiByte(CP_ACP, 0, buf, -1, pszMultiByte, iSize, NULL, NULL);
		logout("get_NodeName %s\n", pszMultiByte);
		free(pszMultiByte);

		if (IsEqualGUID(guidNodeType, KSNODETYPE_DEV_SPECIFIC))
		{
			logout("found one xu node\n");

			hr = m_pDevFilter->QueryInterface(__uuidof(IKsControl),
			(void **)&m_pKsControl);
			if (FAILED(hr))
			{
				logout("Failed to QueryInterface for IKsControl : %x\n", hr);
				continue;
			}
			ExtensionProp.Property.Set = guid;
			ExtensionProp.Property.Id = 0;
			ExtensionProp.Property.Flags = KSPROPERTY_TYPE_GET | KSPROPERTY_TYPE_TOPOLOGY;
			ExtensionProp.NodeId = i;
			ExtensionProp.Reserved = 0;
			hr = m_pKsControl->KsProperty((PKSPROPERTY)&ExtensionProp, sizeof(ExtensionProp), NULL, 0, &ulBytesReturned);
			ReleaseInterface(m_pKsControl);
			if (SUCCEEDED(hr))
			{
				*node = i;
				return S_OK;
			}
			else if (hr == HRESULT_FROM_WIN32(ERROR_MORE_DATA))
			{
				logout("HRESULT_FROM_WIN32(ERROR_MORE_DATA) : %u\n", ulBytesReturned);
				*node = i;
				return S_OK;
			}
			else
			{
				logout("Failed to KsProperty for IKsControl : %x\n", hr);
				continue;
			}
		}
	}
	return hr;
}

HRESULT CaptureVideo::HandleGraphCapturePicture(void)
{
	//////////////////////////////////////////////////////////////////////////
	HRESULT hr;
	long evCode = 0;
	long lBufferSize = 0;
	BYTE *p;
	hr = m_pMediaEvent->WaitForCompletion(INFINITE, &evCode); //
	if (SUCCEEDED(hr))
	{
		switch(evCode)
		{
		case EC_COMPLETE:
			m_pSampGrabber->GetCurrentBuffer(&lBufferSize, NULL);
			p = new BYTE[lBufferSize];
			m_pSampGrabber->GetCurrentBuffer(&lBufferSize, (LONG*)p);// get Current buffer
			g_sampleGrabberCB.SaveBitmap(p,lBufferSize); //save bitmap
			delete [] p;
			p = NULL;
			break;
		default:
			break;
		}
	}
	//////////////////////////////////////////////////////////////////////////
	return hr;
}
