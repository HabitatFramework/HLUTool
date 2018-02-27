HLU GIS Tool
============

HLU GIS Tool 2.4.3

Copyright © 2011 Hampshire Biodiversity Information Centre
Copyright © 2013-2014, 2016 Thames Valley Environmental Records Centre
Copyright © 2014, 2018 Sussex Biodiversity Record Centre

Overview
--------
The HLU GIS Tool provides a user interface for maintaining habitat & Land Use data, including changes to attributes and changes to the spatial data. It also provides an audit trail to indicate when it was edited, why and by whom.

The tool provides an interface that links the spatial and attribute data in multiple software environments - it can link ArcGIS or MapInfo applications
with databases held in Access or SQL Server, and in principle it could link to other database systems such as PostgreSQL and Oracle but these have not been tested.

Features
--------
The HLU GIS Tool provides the following features:
- Ensures that all attributes selected by users are valid and compatible
- Improved data retrieval/update response times over single GIS layers
- Maintains a brief history of all changes made to every habitat polygon
- Stores the data in a relational structure to reduce GIS data volumes
- Enables queries to be produced for a range of purposes using the relational database
- Enables data to be extracted according to user requirements
- Standardises the format of habitat data enabling local, regional or national datasets to be combined

Requirements
------------
The HLU GIS Tool requires the following:

 - Microsoft Windows XP/2003/2008/Vista/Windows 7 or Windows 8
 - 3 GHz or higher processor
 - 2 GB RAM
 - 3 GB of free hard disk space
  (For increased performance a multiple core processor with as much RAM as possible is recommended)

 - Microsoft .NET Framework 3.5 SP1, 4.0 or 4.5 installed.
   (You can download .NET Framework 3.5 and its Service Pack from <http://www.microsoft.com/en-gb/download/details.aspx?id=25150>)

 - Microsoft Access 2000 or later, OR
 - Microsoft SQL Server 2008 Express Edition or later, OR
 - Microsoft SQL Server 2008 or later

 - ArcGIS 10.1 or later (ArcGIS 9.3 is supported by v1.0.1), OR
 - MapInfo 8.0 or later (32bit versions only)

Installation
------------
The HLU GIS Tool setup.exe contains all the required files to install the tool but does not provide any of the GIS layers or relation databases required to run the tool. By default all files are installed in the "Program Files\HLU\HLU GIS Tool" folder (or "Program Files (x86)\HLU\HLU GIS Tool" on 64bit versions of Windows).

Source Code
-----------
The source code for the HLU GIS Tool is open source and can be downloaded from:
<https://github.com/HabitatFramework/HLUTool>

Documentation
-------------
The user guide for the latest release (and for earlier releases) can be viewed online or downloaded as a PDF from:
<https://readthedocs.org/projects/hlugistool-userguide/>

The technical guide for the latest release (and for earlier releases) can be viewed online or downloaded as a PDF from:
<https://readthedocs.org/projects/hlutool-technicalguide/>

Issue Reporting
---------------
To search for existing issues please use:
<https://github.com/HabitatFramework/HLUTool/issues>

To report new issues please use:
<http://forum.lrcs.org.uk/viewforum.php?id=24>

License Information
-------------------
The HLU GIS Tool is free software. You can redistribute it and/or modify it
under the terms of the GNU General Public License as published by the Free
Software Foundation, either version 3 of the License, or (at your option) any
later version.

See the file "License.txt" installed with the tool for information on the
terms & conditions for usage and copying, and a DISCLAIMER OF ALL WARRANTIES
or see <http://www.gnu.org/licenses> for more details of the GNU General Public
License.
