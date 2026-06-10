# Multi Groups Members Nested

An advanced, enterprise-grade Active Directory audit utility designed to multi-thread deep nested group discovery, crawl multi-domain forest architectures, and compile clear compliance reports. 

The application utilizes optimized cache mechanisms to bypass directory lookup bottlenecks, queries explicit identity security descriptors, and programmatically provisions styled multi-tab spreadsheets via low-level Excel COM automation.

## Features
* **Sliding Window Group Extraction:** Leverages Directory Services range tokens (`member;range=`) to dynamically slice and extract memberships from massive security distribution lists without risking timeout restrictions or query drops.
* **Forest-Wide Tree Navigation:** Targets trusted forest domains dynamically (`Get-ADForest`), matching inputs to their precise schema structures and handling localized property evaluations seamlessly.
* **Integrated Object Caching Matrix:** Stores runtime resolution objects inside an internal script-level hash map memory structure to drop redundant domain query rounds and cut down corporate network overhead.
* **Low-Level Array COM Ingestion:** Spasms isolated, background Microsoft Excel processes headlessly to push multi-dimensional data blocks down into individual worksheet tabs via fast memory array injections rather than cell-by-cell rendering.
* **Advanced Access Descriptor Auditing:** Automatically pulls and decodes low-level Access Control Entries (`Get-Acl`) bound to target security groups, isolating individual identity references explicitly mapped with administrative Full Control rights.

## Prerequisites
* **Windows OS** with PowerShell 5.1+
* **RSAT Active Directory Tools** installed (`ActiveDirectory` module).
* **Microsoft Excel Desktop Application** (relies on local availability of the `Excel.Application` COM matrix).

## Usage
1. Open an elevated administrative console terminal session and execute the script.
2. In the temporary text workspace that loads, paste your target list of security or distribution groups (one per line) and close the scratchpad.
3. Supply a custom output destination folder title and specify the respective home subdomain matching configurations.
4. The extraction pipeline will crawl the forest infrastructure, resolve recursive dependencies, and automatically load a Windows Explorer pane directing to your generated multi-tab Excel workbooks.
