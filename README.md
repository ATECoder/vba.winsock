# [VBA Winsock]

TCP Socket-based control and communication workbooks for LXI-based instruments. 

* [Description](#Description)
* [Issues](#Issues)
* [Supported VBA Releases](#Supported-VBA-Releases)
* Project README files:
  * [winsock](/src/winsock/readme.md)
  * [winsock demo](/src/winsock.demo/readme.md)
  * [winsock test](/src/winsock.test/readme.md)
* [Using Prologix](Prologix.md)
* [Attributions](Attributions.md)
* [Change Log](./CHANGELOG.md)
* [Cloning](Cloning.md)
* [Code of Conduct](code_of_conduct.md)
* [Contributing](contributing.md)
* [Legal Notices](#legal-notices)
* [License](LICENSE)
* [Open Source](Open-Source.md)
* [Repository Owner](#Repository-Owner)
* [Authors](#Authors)
* [Security](security.md)

## Description

The ISR VBA Winsock workbooks provide VBA classes for communicating with LXI instruments in desktop platforms using Winsock.

Using the [Prologix] GPIB to Ethernet interface, Winsock can be used to implement some of the capabilities of VXI-11 such as device clear and serial poll.

Otherwise,  Unlike VXI-11 or HiSlip, using these classes these classes do not implement the bus level method for issuing device clear, reading service requests or responding to instrument initiated event. While  control ports for these methods are available in some Keysight instruments, these ports are not part of the standard LXI framework.

## Issues

### read after write delay is required  for Async methods

A delay of 1 ms is required for implementing the asynchronous query method using the TCP Client write and read asynchronous methods. Neither the console nor unit tests are succeptible to this issue. 

## Supported VBA Releases

* TBA

## Repository Owner

* [ATE Coder]

<a name="Authors"></a>
## Authors

* [ATE Coder]  

<a name="legal-notices"></a>
## Legal Notices

Integrated Scientific Resources, Inc., and any contributors grant you a license to the documentation and other content in this repository under the [Creative Commons Attribution 4.0 International Public License], see the [LICENSE](./LICENSE) file, and grant you a license to any code in the repository under the [MIT License], see the [LICENSE-CODE](./LICENSE-CODE) file.

Integrated Scientific Resources, Inc., and/or other Integrated Scientific Resources, Inc., products and services referenced in the documentation may be either trademarks or registered trademarks of Integrated Scientific Resources, Inc., in the United States and/or other countries. The licenses for this project do not grant you rights to use any Integrated Scientific Resources, Inc., names, logos, or trademarks.

Integrated Scientific Resources, Inc., and any contributors reserve all other rights, whether under their respective copyrights, patents, or trademarks, whether by implication, estoppel or otherwise.

[Creative Commons Attribution 4.0 International Public License]:(https://creativecommons.org/licenses/by/4.0/legalcode)
[MIT License]:(https://opensource.org/licenses/MIT)
 
[ATE Coder]: https://www.IntegratedScientificResources.com

[VBA Winsock]: https://github.com/ATECoder/vba.winsock.git

