/////////////////////////////////////////////////////////////////////////////////
// * FB2K Component: COM Automation and ActiveX Interface                    * //
// * Description:    MyCOM Init Source File                                  * //
// * Author:         TT                                                      * //
// * Website:        https://github.com/TT-ReBORN/Georgia-ReBORN             * //
// * Version:        0.1                                                     * //
// * Dev. started:   12-12-2024                                              * //
// * Last change:    22-12-2024                                              * //
/////////////////////////////////////////////////////////////////////////////////


#include "stdafx.h"
#include "MyCOM.h"


namespace {
	class myinitquit : public initquit {
	public:
		void on_init() {
			RegisterMyCOM();
		}

		void on_quit() {
			UnregisterMyCOM();
		}
	};

	static initquit_factory_t<myinitquit> g_myinitquit_factory;
}
