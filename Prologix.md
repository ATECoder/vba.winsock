# Using the Prologix GPIB-ETHERNET (GPIB-Lan) Controller

## Required Applications

* [Netfinder]
* [Prologix GPIB Configurer]

Additional resources are located at the [Prologix] web site.

## Setup

1) Connect the PC to the local network and record it IP address;
2) Connect the Prologix to the local network or directly to the computer using a direct or crossover cable.
3) Open Netfinder;
4) Click _Search_;
5) Netfinder will locate the device and display it's default IP as 0.0.0.0;
6) Click _Asign IP_;
7) Assuming the PC IP address is 192.168.0.100, enter a static IP as follows:
	* IP Address: 192.168.0.252
	* Subnet Mask: 255.255.255.0
	* Default Gateway: 192.168.0.1
8) Open the GPIB Configurer;
9) Select the Prologix in the _Select Device_ panel;
10) Enter the instrument GPIB address, e.g., 16.
11) Enter the identity command, *IDN? to the left of the _Send_ button.
12) Click _Send_;
13) The instrument identity is displayed in the _Terminal_ panel.

[Prologix]: https://prologix.biz/resources/
[Prologix GPIB Configurer]: http://www.ke5fx.com/gpib/readme.htm
[Netfinder]: https://prologix.biz/downloads/netfinder.exe