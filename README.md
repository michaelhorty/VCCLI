# VCCLI
Veracode AST and Security Labs Utility in .NET CORE

Includes HMAC AUTHENTICATION as required by Veracode - original C# example includes a library not supported by .NET core.

This small utility demonstrates basic usage of their API for AST and Security Labs. Requires RESTSHARP and NEWTONSOFT JSON libraries.

USAGE: VCCLI action --param1 param1_value --param2 param2_value

ACTIONS:
--------

help                     Produces this list of actions and parameters

get_seclabs_summary      Returns a summary of Security Labs student & lesson activity

get_dast                 Returns a summary of DAST analyses and scans

get_app_profiles         Returns a summary of Application Profiles

new_dast                 Creates and starts a DAST scan [req:dast_name & dast_url,optional: linkapp_UUID/NAME]

link_dast                Links a DAST scan to an Application Profile [req dast_name & linkapp_UUID/NAME


PARAMETERS:
-----------

--apiID                  The Veracode API ID (if not inside C:\Users\mhorty

--apiKEY                 The Veracode API KEY (if not inside C:\Users\mhorty

--slToken                The Security Labs API KEY

--dast_name              Name of the DAST scan

--dast_url               URL used for a DAST scan

--linkapp_UUID           Used to link a DAST scan to a Profile using APP UUID

--linkapp_Name           Used to link a DAST scan to a Profile using APP NAME

