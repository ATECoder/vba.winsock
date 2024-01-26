# Testing the [cc.isr.winsock] Workbook

[cc.isr.winsock.test] is an Excel workbook for testing the [cc.isr.Winsock] workbook.

## Workbook references

* [cc.isr.winsock] - Winsock workbook
* [cc.isr.Core] - Core work book.
* [cc.isr.core.io] - Core I/O workbook.
* [cc.isr.Test.Fx] - Test framework workbook

## Object Libraries references

* [Microsoft Scripting Runtime]
* [Microsoft Visual Basic for Applications Extensibility 5.3]
* [Microsoft VBScript Regular Expression 5.5]

## Worksheets

The [cc.isr.Winsock.test] workbook includes two worksheets: Identity and TestSheet.

* TestSheet - To run unit tests.

## Scripts

* [unit test]: shortcut to run unit tests.

## Unit Testing

To enable unit testing, the Excel _Trust Center_, which can be found from the _Search_ box, and check _Trust access to the VBA project object model_ from the _Macro Settings_ in the _Trust Center_.  

### Unit testing with the TestSheet Worksheet

Use the following procedure to run unit tests:
1) Click the ___List Tests___ button.
2) The drop down list now includes the list of available test suites;
3) Select a test from the list;
4) Click ___Run Selected Tests___;
   * The list of tests included in the test suite will display.
   * Passed tests display Passed with a green background;
   * Failed tests display Fail with a red background and a message describing the failure.

## Integration Testing

See [cc.isr.winsock.demo]

# Feedback

[cc.isr.winsock.test] is released as open source under the MIT license.
Bug reports and contributions are welcome at the [cc.isr.winsock] repository.

[cc.isr.winsock]: https://github.com/ATECoder/vba.winsock/src/
[cc.isr.winsock.test]: https://github.com/ATECoder/vba.winsock/src/test
[cc.isr.winsock.demo]: https://github.com/ATECoder/vba.winsock/src/demo

[cc.isr.Core]: https://github.com/ATECoder/vba.core
[cc.isr.core.io]: https://github.com/ATECoder/vba.core/src/io
[cc.isr.test.fx]: https://github.com/ATECoder/vba.core/src/testfx

[unit test]: ./cc.isr.winsock.test.unit.test.lnk

[ISR]: https://www.integratedscientificresources.com

[Microsoft Scripting Runtime]: c:\windows\system32\scrrun.dll
[Microsoft Visual Basic for Applications Extensibility 5.3]: <c:/program&#32;files/common&#32;files/microsoft&#32;shared/vba/vba7.1/vbeui.dll>
[Microsoft VBScript Regular Expression 5.5]: <c:/windows/system32/vbscript.dll/3>

