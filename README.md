# VCCLI
Veracode AST and Security Labs Utility in .NET CORE

Includes HMAC AUTHENTICATION as required by Veracode - original C# example includes a library not supported by .NET core.

This small utility demonstrates basic usage of their API for AST and Security Labs. Requires RESTSHARP and NEWTONSOFT JSON libraries.

To run simply download files in folder: https://github.com/michaelhorty/VCCLI/tree/master/bin/Debug/netcoreapp3.1
      VCCLI.exe is the CLI

USAGE: VCCLI.exe action --param1 param1_value --param2 param2_value

Try VCCLI HELP for list of Actions and their associated parameters.

ACTIONS:
--------
- Produce a list of actions and parameters

- Return a summary of Security Labs student & lesson activity

- Return a summary of DAST analyses and scans

- Return a summary of Application Profiles

- Create and starts a DAST scan [req:dast_name & dast_url,optional: linkapp_UUID/NAME]

- Link a DAST scan to an Application Profile [req dast_name & linkapp_UUID/NAME
