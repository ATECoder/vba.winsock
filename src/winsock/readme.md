# About

[cc.isr.winsock] is an Excel workbook implementing TCP Client and Server classes with Windows Winsock API and support higher level [ISR] workbooks.

## Workbook references

* [cc.isr.Core] - Core work book.
* [cc.isr.core.io] - Core I/O workbook.

## Object Libraries references

* [Microsoft Scripting Runtime]
* [Microsoft Visual Basic for Applications Extensibility 5.3]
* [Microsoft VBScript Regular Expression 5.5]

# Key Features

* Encapsulates the Windows API to construct the basic objects for Tcp/IP communication.
* Using Windows Winsock32 calls to construct sockets for communicating with the instrument.

# Main Types

The main types provided by this library are:

* _Winsock_ - initiates a Winsock session.
* _IPv4StreamSocket_ - opens an IPv4 streaming socket to the instrument.
* _TcpCllient_ - Encapsulates the _IPv4StreamSocket_.

## [Testing]

* [cc.isr.winsock.demo] Integration Testing
* [cc.isr.winsock.test] Unit Testing

# Feedback

[cc.isr.winsock] is released as open source under the MIT license.
Bug reports and contributions are welcome at the [cc.isr.winsock] repository.

[cc.isr.winsock]: https://github.com/ATECoder/vba.winsock
[cc.isr.winsock.demo]: https://github.com/ATECoder/vba.winsock/src/demo
[cc.isr.winsock.test]: https://github.com/ATECoder/vba.winsock/src/test

[cc.isr.Core]: https://github.com/ATECoder/vba.core
[cc.isr.core.io]: https://github.com/ATECoder/vba.core/src/io
[cc.isr.test.fx]: https://github.com/ATECoder/vba.core/src/testfx

[ISR]: https://www.integratedscientificresources.com

[Microsoft Scripting Runtime]: c:\windows\system32\scrrun.dll
[Microsoft Visual Basic for Applications Extensibility 5.3]: <c:/program&#32;files/common&#32;files/microsoft&#32;shared/vba/vba7.1/vbeui.dll>
[Microsoft VBScript Regular Expression 5.5]: <c:/windows/system32/vbscript.dll/3>
