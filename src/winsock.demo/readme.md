# Demo of the [cc.isr.winsock] Workbook

[cc.isr.winsock.demo] is an Excel workbook that use the [cc.isr.Winsock] workbook.

## Workbook references

* [cc.isr.winsock] - Winsock workbook
* [cc.isr.Core] - Core work book.
* [cc.isr.core.io] - Core I/O workbook.

## Object Libraries references

* [Microsoft Scripting Runtime]
* [Microsoft Visual Basic for Applications Extensibility 5.3]
* [Microsoft VBScript Regular Expression 5.5]

## Worksheets

* Identity - To query the instrument identity using the *IDN? command.

## Identity querying

Follow this procedure for reading the instrument identity string:

* Select the Identity sheet.
* Enter the instrument dotted IP address, such as `192.168.252`;
* Enter the instrument port:
  * `5025` for an LXI instrument or
  * `1234` for a GPIB instrument connected via a GPIB-Lan controller such as the [Prologix GPIB-Lan controller].
* Click ___Read Identity___ to read the instrument identity using the `*IDN?` query command:
  * Check the following options:
	* ___Using Winsock Read Raw___ -- reads one character at a time till the default termination;
	* ___Using Winsock Buffer Read___ -- reads a buffer of up to 1024 characters at a time;
	* ___Using Tcp Client___ -- reads using the TCP Client class.

# Feedback

[cc.isr.winsock.demo] is released as open source under the MIT license.
Bug reports and contributions are welcome at the [cc.isr.winsock] repository.

[cc.isr.winsock]: https://github.com/ATECoder/vba.winsock/src/
[cc.isr.winsock.demo]: https://github.com/ATECoder/vba.winsock/src/demo
[cc.isr.winsock.test]: https://github.com/ATECoder/vba.winsock/src/test

[cc.isr.Core]: https://github.com/ATECoder/vba.core
[cc.isr.core.io]: https://github.com/ATECoder/vba.core/src/io
[cc.isr.test.fx]: https://github.com/ATECoder/vba.core/src/testfx

[ISR]: https://www.integratedscientificresources.com

[Microsoft Scripting Runtime]: c:\windows\system32\scrrun.dll
[Microsoft Visual Basic for Applications Extensibility 5.3]: <c:/program&#32;files/common&#32;files/microsoft&#32;shared/vba/vba7.1/vbeui.dll>
* [Microsoft VBScript Regular Expression 5.5]: <c:/windows/system32/vbscript.dll/3>
[User-Defined Type Not Defined error]: https://stackoverflow.com/questions/19680402/compile-throws-a-user-defined-type-not-defined-error-but-does-not-go-to-the-of#:~:text=So%20the%20solution%20is%20to%20declare%20every%20referenced,objXML%20As%20Variant%20Set%20objXML%20%3D%20CreateObject%20%28%22MSXML2.DOMDocument%22%29


