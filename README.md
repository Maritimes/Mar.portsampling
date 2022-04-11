# portsampling
A suite of tools for Port Samplers

## Installation
To make the "pushButton" reports, run the following code.  This will generate a batch file on your desktop (Windows only) that you can double click to generate the reports on demand.  All reports will be created in "C:\\DFO-MPO\\PORTSAMPLING"


```R
require(devtools)
install_github('Maritimes/Mar.portsampling')
require(portsampling)
makePushButton()
```
