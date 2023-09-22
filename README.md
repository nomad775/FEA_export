# FEA_export
SolidWorks VBA macro for exporting FEA study data to an XML file.

>**IMPORTANT NOTE** This is far from a completed project.  This works for what I need it do to. There are many areas where values and code is partialy implomented, but commented out, so modification for further purposes should be easy

## Files
The SWP file is the SolidWorks macro which is a binary file.  The .BAS files are the exported modules which are text files.
The XML file is the template used for creating the data file.

## References
This macro requires a reference to the MicroSoft MSXML v6.0.
It also references the SolidWorsk Simulation 2020 type library, which needs to be updated for newer versions or switched to late binding.
To add this reference, in the VBA editor select the Tools menu, select References, find and check "Microsoft XML, v6.0"

## What It Does
This SolidWorks macro exports the study data for a FEA analysis into a XML file. It was created for importing into Word for creating a report, similar to the built in report generator in SolidWorks, although other uses are possible. In Word the XML file is easily imported as a Custom XML Part and fill in Content Controls.

The data exported is in the following 5 sections:
 - Study Properties
 - Mesh Properties
 - Materials
 - Loads
 - Restraints
 
 Specific data in each section can be seen by investigating the XML file.
