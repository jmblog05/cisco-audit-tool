# cisco-audit-tool
Python tool for Cisco CLI log analysis and Excel report generation.

## Overview

Cisco Audit Tool is a Python application that analyzes Cisco CLI log files and automatically generates Excel reports.

This project was created to improve network operation efficiency in environments where API access is unavailable.

## Why I created this tool

Managing hundreds of network devices manually is time-consuming.

Since API access was prohibited in my work environment, I developed this tool to analyze CLI outputs and automatically generate inventory and status reports in Excel.

## Features

- Inventory collection
- IP address summary
- Hardware model detection
- Software version detection
- OSPF neighbor status
- BGP peer status
- Default route check
- NTP synchronization status
- Excel report generation

## Environment

Python 3.x

Windows

openpyxl

tkinter

## How to use
1. Execute commands on Cisco routers.
2. Save CLI logs.
3. Select the log directory.
4. Generate Excel report.

## Sample Output
<img width="773" height="241" alt="result_open" src="https://github.com/user-attachments/assets/8d890410-cbe1-485a-a5fa-927c6e2ab702" />

<img width="1248" height="578" alt="status_1" src="https://github.com/user-attachments/assets/2db55ca7-a26b-48b0-83b3-a095046e7510" />


## Technical Details

Parsing CLI logs

Regular expressions

Excel generation

GUI application

## Future Improvements

- Support more Cisco commands
- Support Juniper devices
- AI-powered log analysis using LangChain
