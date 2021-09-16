# RTA600E-DATA-PULL-APPLICATION
Get data from RTA600E device buffer.

[![Build Status](https://travis-ci.org/joemccann/dillinger.svg?branch=master)](https://travis-ci.org/joemccann/dillinger)

### Topic
- Install Driver and setup Com Port.
- Application Core.
- Run Application.
- Device Info

---

### Install Driver and setup Com Port:

Find the converter you are using and the driver setup manual. If you using `EPC-406` and `RS-422/485` connection then you can easily setup driver.

### Application Core:

This application have some core file which is,
- tsmcom32.dll
- Module1.bas
- Module2.bas
- RTA600.INI

### Application Installation
- At first we install the EPC-406 or others EPC driver.
- We install RMS software and check the com port.
- After check com port we search the device by node id and save the device node id in RMS software.
- After saving the node id first we try to pull some data from device using RMS software.
- After successful pull data from device we add this COM port & NODE id in ```RTA600.INI``` file.
- Then we make a folder which name is 'Data' in D:\ drive.
- Then open the application and run it.

### How to write RTA600.INI file
```
[Main]
NID01=1,COM1,9600,50
NID02=2,COM1,9600,50
NID03=5,COM1,9600,50
NID04=6,COM1,9600,50
NID05=7,COM1,9600,50
[Output]
opath=D:\DATA\
ofile650=C:\CARDATA\RTA600.TXT
```
- ```NID01=1,COM1,9600,50``` 'NID01' represent the index number. It means how many device are connected. '1' represent the NODE ID. 'COM1' represent the port name where our driver installed. '9600' this is default buffer size. '50' this is default cursor size.
- If we add a new device in this list then we add a new line between 'NID05' and 'Output'. We write ```NID06=8,COM1,9600,50```.

** Attached library name is ```tsmcom32.dll```

### Programming Module
In this application have total 3 files:
- RSTATUS.frm
- Module1.bas
- Module2.bas

1. ```RSTATUS.frm```  This file useing for output design.
