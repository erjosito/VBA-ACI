VBA-ACI

Author: Jose Moreno
Version: 0.10

<h1>Why?</h1>

Initially as a prototype to demonstrate two things:
1. You can code against ACI using any programming language that supports REST calls
2. Automation does not necessarily mean DevOps or Cloud, it typically starts by 
   small optimizations in existing processes (like who configures network ports
   when servers are attached).

<h1>What?</h1>

Essentially this project is a VBA module that implements native REST calls to ACI's API, 
and an example spreadsheet that uses that module to partially configure an ACI fabric.

The core of the spreadsheet handles port configuration, but additional functionality has
been implemented to support configuration of EPGs, BDs, VRFs, etc.

Port configuration is done specifying on which "rack" a server is connected. Switches
(and optionally FEXes) are assigned to racks in pairs. When a server is attached you can
decide whether the connection is to both switches or not, and whether it is in active/
active or active/standby mode.
