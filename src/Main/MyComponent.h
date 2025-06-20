/////////////////////////////////////////////////////////////////////////////////
// * FB2K Component: COM Automation and ActiveX Interface                    * //
// * Description:    MyComponent Header File                                 * //
// * Author:         TT                                                      * //
// * Website:        https://github.com/TT-ReBORN/foo_mycom                  * //
// * Version:        0.2                                                     * //
// * Dev. started:   12-12-2024                                              * //
// * Last change:    22-12-2024                                              * //
/////////////////////////////////////////////////////////////////////////////////


#pragma once
#include "MyComponent_Main.h"


//////////////////////
// * MY COMPONENT * //
//////////////////////
#pragma region My Component
class MyComponent {
public:
	MyComponent() = default;
	~MyComponent() = default;

	static void InitMyComponent();
	static void QuitMyComponent();

	// Public getters for external access
	static HWND MainWindow() { return mainHwnd; }
	static MyComponentMain* Main() { return myComponentMain.get(); }

private:
	static inline HWND mainHwnd = nullptr;
	static inline std::unique_ptr<MyComponentMain> myComponentMain = nullptr;
};
#pragma endregion
