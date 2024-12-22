This component is a template implementation for COM automation with ActiveX.<br />
It allows for seamless integration and interaction with COM objects in your scripts<br />
using either registration-free COM or registration with a registry entry:<br />
`MyCOM.cpp: std::wstring g_regMethod = L"RegFree"; // "RegFree" or "RegEntry"`<br />
<br />
**Usage:**<br />
To demonstrate, create a new instance of the ActiveX object and call its method:
```
const myCom = new ActiveXObject('MyCOM');
myCom.PrintMessage();
console.log('myCom', myCom);
```
