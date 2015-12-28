// LoaderDlg.cpp : implementation file
//

#include "stdafx.h"
#include "Loader.h"
#include "LoaderDlg.h"
#include ".\loaderdlg.h"
#include "Shlwapi.h"
#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// Dialog Data
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Implementation
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
END_MESSAGE_MAP()


// CLoaderDlg dialog



CLoaderDlg::CLoaderDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CLoaderDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CLoaderDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CLoaderDlg, CDialog)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDCANCEL, OnBnClickedCancel)
	ON_BN_CLICKED(IDOK, OnBnClickedOk)
END_MESSAGE_MAP()


// CLoaderDlg message handlers

BOOL CLoaderDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	// TODO: Add extra initialization here
		// Demo 1, the simplest use of the class
	
	CSplashScreenEx *pSplash=new CSplashScreenEx();
	pSplash->Create(this,NULL,5000,CSS_FADE | CSS_CENTERSCREEN | CSS_SHADOW);
	pSplash->SetBitmap(IDB_SPLASH,255,255,255);
	pSplash->Show();

	if (PathFileExists("FaenaMRP2004.exe"))
	{
		ShellExecute(NULL,NULL,"FaenaMRP2004.exe", NULL, NULL, SW_SHOWNORMAL);
		Sleep(5000);
	}
	else
	{
		pSplash->Hide();
		AfxMessageBox("Executable file not found! Please reinstall the software.", MB_OK | MB_ICONSTOP, 0);
	}
	/////////////////////////////////////////////////////////////
	/////////////////////////////////////////////////////////////

	// Demo 2
	/*CSplashScreenEx *pSplash=new CSplashScreenEx();
	pSplash->Create(this,"CSplashScreenEx dynamic text:",0,CSS_FADE | CSS_CENTERSCREEN | CSS_SHADOW);
	pSplash->SetBitmap(IDB_SPLASH,255,0,255);
	pSplash->SetTextFont("Impact",100,CSS_TEXT_NORMAL);
	pSplash->SetTextRect(CRect(125,60,291,104));
	pSplash->SetTextColor(RGB(255,255,255));
	pSplash->SetTextFormat(DT_SINGLELINE | DT_CENTER | DT_VCENTER);
	pSplash->Show();

	Sleep(1000);
	pSplash->SetText("You can display infos");

	Sleep(1000);
	pSplash->SetText("While your app is loading");

	Sleep(1000);
	pSplash->SetText("Just call Hide() when loading");
	
	Sleep(1000);
	pSplash->SetText("is finished");
	Sleep(1500);

	pSplash->Hide();*/

	OnOK();
	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CLoaderDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CLoaderDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CLoaderDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CLoaderDlg::OnBnClickedCancel()
{
	// TODO: Add your control notification handler code here
	OnCancel();
}

void CLoaderDlg::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here
	OnOK();
}
