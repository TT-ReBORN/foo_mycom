/////////////////////////////////////////////////////////////////////////////////
// * FB2K Component: COM Automation and ActiveX Interface                    * //
// * Description:    MyComponent Source File                                 * //
// * Author:         TT                                                      * //
// * Website:        https://github.com/TT-ReBORN/foo_mycom                  * //
// * Version:        0.2                                                     * //
// * Dev. started:   12-12-2024                                              * //
// * Last change:    22-12-2024                                              * //
/////////////////////////////////////////////////////////////////////////////////


#include "MyComponent_PCH.h"
#include "MyCOM.h"
#include "MyComponent.h"


/////////////////////////
// * COMPONENT SETUP * //
/////////////////////////
#pragma region Component Setup
// Declaration of your component's version information
DECLARE_COMPONENT_VERSION(
	"MyCOM", "0.2", "COM Automation and ActiveX Interface by TT.\n\n"
	"This component is a template implementation for COM automation\n"
	"with ActiveX. It allows for seamless integration and interaction\n"
	"with COM objects in your scripts using either registration-free COM\n"
	"or registration with a registry entry:\n"
	"MyCOM.cpp: std::wstring g_regMethod = L\"RegFree\";\n"
	"You can use \"RegFree\" or \"RegEntry\"\n\n"
	"Usage:\n"
	"To demonstrate, create a new instance of the ActiveX object\n"
	"and call its method:\n"
	"const myCom = new ActiveXObject('MyCOM');\n"
	"myCom.PrintMessage();\n"
	"console.log('myCom', myCom);\n"
);

// Define the GUID for your component
// {152C13C3-5F06-4314-B1F7-78ED35C62DAD}
DEFINE_GUID(MyCOM, 0xd4e8f8d6, 0x5b1f, 0x4f4d, 0x8a, 0x8f, 0x11, 0x6c, 0x3b, 0x8b, 0x1e, 0x7e);

// Implement the class_guid for component_installation_validator
// This will prevent users from renaming your component around (important for proper troubleshooter behaviors) or loading multiple instances of it.
VALIDATE_COMPONENT_FILENAME("foo_mycom.dll");

// Activate cfg_var downgrade functionality if enabled. Relevant only when cycling from newer FOOBAR2000_TARGET_VERSION to older.
FOOBAR2000_IMPLEMENT_CFG_VAR_DOWNGRADE;
#pragma endregion


/////////////////////////////
// * MAIN INITIALIZATION * //
/////////////////////////////
#pragma region Main Initialization
void MyComponent::InitMyComponent() {
	mainHwnd = core_api::get_main_window();

	if (!mainHwnd) return;

	myComponentMain = std::make_unique<MyComponentMain>(mainHwnd);
	// MyComponentMain init implementation
}

void MyComponent::QuitMyComponent() {
	// MyComponentMain cleanup implementation
	myComponentMain.reset();
	MyCOM::QuitMyCOM();
}

namespace {
	FB2K_ON_INIT_STAGE(MyCOM::InitMyCOM, init_stages::after_config_read);
	FB2K_ON_INIT_STAGE(MyComponent::InitMyComponent, init_stages::after_ui_init);
	FB2K_RUN_ON_QUIT(MyComponent::QuitMyComponent);
}
#pragma endregion
