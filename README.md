# Digital Ordering

## Description
This is the complete Visual Studio project which is used for the digital ordering workflow at Fraunhofer IPM. Altough this repository does include all C#-Code that was written for the workflow, it doesn't include other components such as SharePoin Designer Workflows that are needed for the complete workflow as well. The whole ordering workflow is very specific and that's the reason why the program was not designed for different use cases. It is designed to meet the requirements of the process in Fraunhofer IPM exactly. That's also the reason why there is additional configuration to do in SharePoint. This published solution does for example also not create any sites, libraries and collumns. Those were created manually.

However, you can of course pull this repository and build and deploy this solution immediately to your SharePoint On-Premise-Server.

## Architecture
This is a technical overview of the different components that are needed.



## Requirements
- SharePoint 2010 On-Premise
- Full access on the SharePoint server(s): This is required for building and deploying the solution to SharePoint
- Visual Studio
- SharePoint Designer 2010
- Enabled Info Path Services in SharePoint
 
## Required SharePoint components (Sites, libraries, lists)
  Heres a list of the sites and libraries that are needed. There are also many collumns which need to exist. Feel free to contact me if you want 

- Site "http://intranet/bestellung"

The following librariees have to exist in the site.
- Form libraries
-- "Auftragsformular"
-- "Auftragsformular-Archiv"
- Document libraries:
-- "Auftragszettel"
-- "Begründung"
-- "Temp"
-- "Auftragszettel-Archiv"
