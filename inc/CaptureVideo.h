#ifndef __CAPTUREVIDEO_H__
#define __CAPTUREVIDEO_H__
#include "SampleGrabberCallback.h"

#include <ks.h>
#include <ksmedia.h>
#include <ksproxy.h>
#include <vidcap.h>
#include <atlbase.h>
#include <string>

class CaptureVideo
{
public:
	CaptureVideo();
	~CaptureVideo();
	HRESULT InitializeEnv(); //initialize environment
	HRESULT EnumAllDevices(HWND hCombox);  //Enumeration all devices
	void CloseInterface(); //close all interface
	HRESULT OpenDevice(int deviceID, LONG nLeft, LONG nTop, LONG nWide, LONG nHeight);
    void Stop(void);
    HRESULT RestartDevice();
	HRESULT BindFilter(int deviceID, IBaseFilter **pBaseFilter);
	HRESULT SetupVideoWindow(LONG nLeft, LONG nTop, LONG nWidth, LONG nHeight);
	void ResizeVideoWindow(LONG nLeft, LONG nTop, LONG nWidth, LONG nHeight);
	void GrabOneFrame(BOOL bGrab); // use ISampleGrabberCallbackCB get a picture
	void TestMsg(void);
	HRESULT FindExtensionNode(IKsTopologyInfo *pKsTopologyInfo, GUID guid, DWORD *node);
	HRESULT HandleGraphCapturePicture(void); //capture picture  use GetCurrentBuffer

private:
	IGraphBuilder *m_pGraphBuilder;
	ICaptureGraphBuilder2 *m_pCaptureGB;
	IMediaControl *m_pMediaControl;
	IBaseFilter *m_pDevFilter;
	ISampleGrabber *m_pSampGrabber;  
	IMediaEventEx *m_pMediaEvent;
	IKsControl *m_pKsControl;
    std::string m_szDevicePath;

public:
	int m_nCaptureDeviceNumber;      //Device Count
	//vector<string> m_vecCapDeviceName; //the Device name
	TCHAR m_pCapDeviceName[10][MAX_PATH]; //the Device name
	HWND m_App;
	BOOL m_bGetOneShot;
	BOOL m_bConnect;
	IVideoWindow *m_pVideoWindow;
};


#endif  //__CAPTUREVIDEO_H__