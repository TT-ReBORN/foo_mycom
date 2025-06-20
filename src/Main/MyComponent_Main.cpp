/////////////////////////////////////////////////////////////////////////////////
// * FB2K Component: COM Automation and ActiveX Interface                    * //
// * Description:    MyComponent Main Source File                            * //
// * Author:         TT                                                      * //
// * Website:        https://github.com/TT-ReBORN/foo_mycom                  * //
// * Version:        0.2                                                     * //
// * Dev. started:   12-12-2024                                              * //
// * Last change:    22-12-2024                                              * //
/////////////////////////////////////////////////////////////////////////////////


#include "MyComponent_PCH.h"
#include "MyComponent_Main.h"


//////////////////////////////////
// * CONSTRUCTOR & DESTRUCTOR * //
//////////////////////////////////
#pragma region Constructor & Destructor
MyComponentMain::MyComponentMain(HWND hWnd) :
	mainHwnd(hWnd) {
}

MyComponentMain::~MyComponentMain() = default;
#pragma endregion


////////////////////////
// * PUBLIC METHODS * //
////////////////////////
#pragma region Public Methods
int MyComponentMain::GetWindowWidth() {
	RECT rect;
	return GetWindowRect(mainHwnd, &rect) ? rect.right - rect.left : -1;
}

void MyComponentMain::SetWindowWidth(int width) {
	const int height = GetWindowHeight();
	SetWindowPos(mainHwnd, nullptr, 0, 0, width, height, SWP_NOMOVE | SWP_NOZORDER);
}

int MyComponentMain::GetWindowHeight() {
	RECT rect;
	return GetWindowRect(mainHwnd, &rect) ? rect.bottom - rect.top : -1;
}

void MyComponentMain::SetWindowHeight(int height) {
	const int width = GetWindowWidth();
	SetWindowPos(mainHwnd, nullptr, 0, 0, width, height, SWP_NOMOVE | SWP_NOZORDER);
}

void MyComponentMain::SetWindowSize(int width, int height) {
	SetWindowPos(mainHwnd, nullptr, 0, 0, width, height, SWP_NOMOVE | SWP_NOZORDER);
}
#pragma endregion
