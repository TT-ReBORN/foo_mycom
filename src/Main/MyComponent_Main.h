/////////////////////////////////////////////////////////////////////////////////
// * FB2K Component: COM Automation and ActiveX Interface                    * //
// * Description:    MyComponent Main Header File                            * //
// * Author:         TT                                                      * //
// * Website:        https://github.com/TT-ReBORN/foo_mycom                  * //
// * Version:        0.2                                                     * //
// * Dev. started:   12-12-2024                                              * //
// * Last change:    22-12-2024                                              * //
/////////////////////////////////////////////////////////////////////////////////


#pragma once


//////////////////////////
// * MyCOMPONENT MAIN * //
//////////////////////////
#pragma region MyComponent Main
class MyComponentMain {
public:
	explicit MyComponentMain(HWND hWnd);
	virtual ~MyComponentMain();

	HWND mainHwnd = nullptr;

	int GetWindowWidth();
	void SetWindowWidth(int width);
	int GetWindowHeight();
	void SetWindowHeight(int height);
	void SetWindowSize(int width, int height);
};
#pragma endregion
