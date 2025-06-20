# MyCOM

MyCOM is a foobar2000 component and a template implementation of a COM/ActiveX interface designed for seamless integration with JavaScript.
It enables efficient COM automation through ActiveX, supporting both registration-free COM and registry-based registration methods.

<br>

## Features
- Provides a robust COM/ActiveX interface for JavaScript interaction.
- Supports both registration-free COM and registry-based registration.
- Exposes getter and setter properties for window dimensions as examples.
- Includes methods for manipulating window size and displaying messages as examples.

<br>

## Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/TT-ReBORN/foo_mycom.git
   ```
2. Build the component using a C++ compiler.
3. Configure the COM registration method in `MyCOM.cpp`, see [Configuration](#configuration).
4. Copy the compiled DLL to the Foobar2000 user-components directory, e.g. `foobar2000\profile\user-components-x64\foo_mycom`

<br>

## Configuration
Specify the COM registration method by setting the `regMethod` parameter in the following static methods in `MyCOM.cpp`:

```h
	static HRESULT RegisterMyCOM(std::wstring_view regMethod = L"RegFree"); // "RegFree" or "RegEntry"
	static HRESULT UnregisterMyCOM(std::wstring_view regMethod = L"RegFree"); // "RegFree" or "RegEntry"
	static HRESULT QuitMyCOM(std::wstring_view regMethod = L"RegFree"); // "RegFree" or "RegEntry"
```

- **`RegFree`**: Enables registration-free COM for simplified deployment through [MinHook](https://github.com/TsudaKageyu/minhook).
- **`RegEntry`**: Uses traditional registry-based COM registration.

<br>

## Usage
Use the `foo_mycom` user-component in Foobar2000's JavaScript environment (e.g. via Spider Monkey Panel or JSplitter).<br>
Instantiate an ActiveX object and interact with its properties and methods as shown below:

```javascript
// Create an instance of the ActiveX object
const myCOM = new ActiveXObject('MyCOM');

// Log the myCOM object
console.log('myCOM', myCOM);

// Getter properties
console.log(myCOM.WindowWidth);  // Retrieves the window width
console.log(myCOM.WindowHeight); // Retrieves the window height

// Setter properties
myCOM.WindowWidth = 1000;  // Sets the window width
myCOM.WindowHeight = 500;  // Sets the window height

// Methods
myCOM.SetWindowSize(1000, 500); // Sets both window width and height
myCOM.PrintMessage();           // Displays a popup dialog via the API
```

<br>

## Requirements
- [foobar2000](https://www.foobar2000.org)
- [Spider Monkey Panel](https://github.com/TheQwertiest/foo_spider_monkey_panel) or
  [JSplitter](https://foobar2000.ru/forum/viewtopic.php?t=6378) for JavaScript and API support
- C++ compiler and the Foobar2000 SDK ( already included in the solution ) for building the component.

<br>

## Contributing
Contributions are welcome! Please submit a pull request or open an issue to discuss improvements or bug fixes.

<br>

## Credits
- Peter Pawlowski for [foobar2000 and SDK](https://www.foobar2000.org).
- Tsuda Kageyu for [MinHook](https://github.com/TsudaKageyu/minhook).
