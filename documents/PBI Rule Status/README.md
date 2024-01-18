# PBI Rule Status

### Description

PBI Rule Status.xslm is a macro-enabled workbook that uses the Zenhub GraphQL API to extract ddf-core-poc PBI information from the CORE DX workspace and then combines it with DDF rule definitions downloaded from the CORE Rule Editor.  The workbook contains three sheets:

- **POC PBIs**: A list of all open and closed ddf-core-poc PBIs (Product Backlog Items) from the CORE DX zenhub workspace, displayed in descending order of issue number (i.e., most recently created appears first).  The sheet contains the following columns:
    - **number**: The issue number
    - **title**: The issue title
    - **body**: The description of the issue
    - **state**: The issue status (OPEN or CLOSED)
    - **pipeline.name**: The name of the pipeline that currently contains the PBI, which is always "Closed" for closed issues but may be "New Issues", "Icebox", "Product Backlog", "Sprint Backlog", "In Progress", "Review", "QA" (not currently used for the project), or "Done".
    - **estimate.value**: The amount of effort estimate for the PBI
    - **epic.title**: The name of the epic to which the PBI is assigned. If the PBI is assigned to more than one epic, only the first is shown
- **Define Rule PBIs**: A list of all the "Define rule" PBIs and corresponding (actual or expected) "Create rule" PBIs.  "Define rule" PBIs are identified as any PBI whose title starts with "Define" and contains both "rule" and a colon (:). "Create rule" PBIs are identified as any PBI whose title starts with "Create" and contains both "rule" and a colon (:). The correspondence between "Define rule" PBIs and "Create rule" PBIs is determined by the desciption (or body) of the "Create rule" PBI, which must end with # followed by the issue number for the corresponding "Define rule" PBI.  For example, a "Create rule" PBI with a description/body of "Create rule define for #26" would correspond with the "Define rule" whose issue number is 26.  More than one "Create rule" PBI may correspond with the same "Define rule" PBI.  The sheet contains 7 columns (A-G) with "Define rule" PBI attributes, which are as described above, and an additional 7 columns (H-N) for "Create rule" PBI attributes:
    - **Create Rule PBIs.\<attribute\>** (e.g., Create Rule PBIs.number): 5 columns which contain the attribute values for the corresponding "Create rule" PBI as described for the corresponding "POC PBIs" sheet column above.  If there is no corresponding "Create rule" PBI, these 5 columns will be blank.
    - **Create Rule Title**: The actual or expected title for the corresponding "Create rule" PBI. If a corresponding "Create rule" PBI exists, its title is shown, otherwise the expected title is displayed (which is a copy of the "Define rule" PBI's title with "Define" changed to "Create").
    - **Create Rule Desciption**: The actual or expected description/body for the corresponding "Create rule" PBI. If a corresponding "Create rule" PBI exists, its description/body is shown, otherwise the expected description/body is displayed (which is "Create rule defined for \#\<n\>" where \<n\> is the "Define rule" PBI's issue number).
- **Rules**: A list of rules exported from the CORE Rules Editor using the "Export Rules (as filtered)" button with the names of the Zenhub pipelines containing the corresponding "Define rule" and "Create rule" PBIs (in the "Define Rule Pipeline" and "Create Rule Pipeline" columns; columns J and K respectively).  The corresponding "Create rule" PBI is identified as the PBI whose title starts with "Create rule \<Rule ID\>:".  The corresponding "Define rule" PBI is identified either as the PBI whose title starts with "Define rule \<Rule ID\>:", or (preferably) as the "Define rule" that corresponds with the identified "Create rule" (as described above for the "Define Rule PBIs" sheet).  The general expectation is that the Rule ID will be added to the title of "Create rule" PBIs only, and this will be done only when the corresponding rule has been created in the Rule Editor.

### Configuration

The Zenhub GraphQL API requires an API key for authentication. To generate your Personal API Key, go to the [API section of your Zenhub Dashboard](https://app.zenhub.com/settings/tokens). From here, you can create a new Personal API Key for the GraphQL API by entering a key name (e.g., "DDFPOC"), clicking the "Generate new key" button and copying the generated key value (which is a long string starting with "gh_").  To make this key accessible by the PBI Rule Status workbook, it needs to be stored in an environment variable called "ZENHUB_API_KEY".  To create a user environment variable in Windows 11:
1. Type "environment" in the "Search" box on the taskbar and click on the "Edit environment variables for your account" option.
2. In the "Environment Variables" dialog, click on the "New..." button under the "User variables for \<username\>" list (at the top).
3. In the following "New User Variable" dialog fields:
   - **Variable name**: type ZENHUB_API_KEY
   - **Variable value**: paste the "gh_..." key value copied as described above.
4. Click OK in both dialogs.

*Refer to system documentation or search online for instructions to create an enviroment variable in other operating systems.*

### Functionality

The workbook contains macros to:
- Retrieve the Zenhub API key from the ZENHUB_API_KEY environment before refreshing the tables in the sheets containing PBI information.
- Prompt for the downloaded Rules CSV file before refreshing the table in the "Rules" sheet.
- Remove the API key and Rules CSV file values before saving the file.

The first two macros execute when the file is opened or the relevant tables are refreshed, and the third macro executes whenever the file is saved.

#### Before opening the PBI Rule Status file:

Download a CSV file from the CORE Rules Editor:
1. Log into the [CORE Rules Editor](https://rule-editor.cdisc.org/)
2. Clear the automatically applied filter for the Creator column
3. Type "DDF" (without quotes) in Search... box for the Standards column to show all DDF rules
4. Click on the "Export Rules (as filtered)" button (at the top of the page immediately to the left of the "Search YAML..." field)
5. Make note of or specify/change the name and location of the downloaded "Rules.csv" file (which may be assigned a numeric suffix if previously downloaded)

#### When opening the PBI Rule Status file:

- You may be prompted to enable macros and/or enable external content.  If prompted, click the "Enable..." button.
- All information will be automatically refreshed.  When prompted, navigate to and select the downlowed Rules CSV file in the "Select Downloaded Rules CSV File" dialog and click the "Load" button.

#### To refresh information when the PBI Rule Status file is already open:

1. Download a new copy of the Rules CSV file by following the "Before opening..." instructions above.
2. Refresh the tables in the three sheets:
   - *Either* individually: in each sheet:
       -  *Either* right-click within the table and choose "Refresh" or
       -  *Or* select any cell within the table and click using the "Refresh" button on the "Table Design" menu
    -  *Or* all together: select any cell within any of the tables and choose "Refresh All" from the dropdown on the the "Refresh" button on the "Table Design" menu.
  
*Note that using "Refresh" on the "Query" may not work correctly because this does not trigger the macros to retrieve the API key or prompt for a Rules CSV file.*

#### To save the file:

Use normal "Save" or "Save As..." functionality.  Before the file is saved, the API key and Rules CSV file name will be automatically removed from the underlying queries.


