# cisco-audit-tool
Python tool for Cisco CLI log analysis and Excel report generation.

## Overview

Cisco Audit Tool is a Python application that parses Cisco CLI log files and automatically generates Excel reports for network inventory and health checks.

The tool was developed to improve operational efficiency in environments where API access is unavailable.

## Why I created this tool

This project was inspired by a real operational challenge.

In my work environment, network devices could not be accessed through APIs due to security restrictions. Collecting inventory and operational information manually from CLI outputs was time-consuming and error-prone.

To solve this problem, I developed a Python-based tool that automatically parses CLI outputs and generates Excel reports.

## Features

- Collect IP address information
- Detect hardware models
- Detect software versions
- Detect uptime
- Check interface status
- Check OSPF neighbor status
- Check BGP peer status
- Check default routes
- Check NTP synchronization
- Generate Excel reports

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

- Developed as a Windows GUI application using `tkinter`.
- Distributed as a `.pyw` application so that no console window appears during execution, providing a cleaner user experience.

## Future Improvements

- Support more Cisco commands
- AI-powered log analysis using LangChain

## Author

Created by jmblog05

Feedback and suggestions are always welcome.
