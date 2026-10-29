# About this project
It started out as simple excel templates to collapse multiline to single line config and using formulas to display (e.g., VS, pool) configurations.

It eventually grow to a point whereby manual extraction of configuration is no longer feasible as there are too many configuration items in different excel templates and there is a need to correlate the data across different templates.

Today this script extracts the following data and correlate the data across different sheets with the help of some creative excel formulas.

bigip.conf - profile, route, node, pool, SNAT pool, virtual server, virtual address, client SSL profile, server SSL profile, certificate, key, auth, cipher group, cipher rule, monitor data

bigip_base.conf - hostname, management-ip, management-route, self IP, trunk, VLAN, DNS, NTP, provisioning, SNMP, SSHD, HTTPD, Password Policy, CLI global-settings, sys global-settings, syslog, cm device, cm device-group, cm traffic-group, cm trust-domain, net route-domain, net tunnels, sys state-mirroring, security firewall management-ip-rule, auth partition

bigip_user.conf - user

BigDB.dat - UI.logAccess, Config.Auditing

ts_db.data.PLC.PL_POLICIES.cstmp - ASM policy

bigip_gtm.conf - gtm datacenter, gtm prober-pool, gtm region, gtm server, gtm topology, gtm global-settings, gtm monitor, gtm pool, gtm wideip

certificates - device certificates, ssl certificiate and keys, gtm certificates

# Prerequisites:

7-zip installed

Microsoft Excel installed

Create SrcDir (e.g., c:\Temp\ExtractSrc) and DstDir (e.g., c:\Temp\ExtractDst) folder

Place ReportTemplate (e.g., F5 BIG-IP UCS Config Template v1.0.xlsx) in SrcDir

Place one or more UCS files in SrcDir (recommend to test with one UCS file first)

Run Windows Powershell ISE as Administrator

# Steps to run script:
Open "F5 BIG-IP UCS Config Extraction Script v1.0.ps1" from Windows Powershell ISE

Click on Run script button in Windows Powershell ISE

After running the script, the 1st sheet shows you a summary of the system information, statistics on the number of each configuration item and hardening configuration based on CIS standard.
