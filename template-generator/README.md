# USDM Comformance Rules Test Data Template Generator

```
usage: python create_template.py [-h] -x XMI_FILE -c CT_FILE -a API_SPEC [-o OUTPUT_FILE]

USDM Comformance Rules Test Data Template Generator. Reads the Enterprise Architect XMI export file containing definition of the USDM UML model (specified in the -x option), the USDM Controlled Terminology Excel file (specified in the -c option), and the USDM API specification in YAML format (specified in the -a option) and creates an Excel template file (optionally specified in the -o option) for test data entry.

options:
  -h, --help            show this help message and exit
  -x XMI_FILE, --xmi_file XMI_FILE
                        USDM XMI export file (e.g., <DDF version>/Deliverables/UML/USDM_UML.xmi)
  -c CT_FILE, --ct_file CT_FILE
                        USDM CT Excel file (e.g., <DDF version>/Deliverables/CT/USDM_CT.xlsx)
  -a API_SPEC, --api_spec API_SPEC
                        USDM API specification YAML file (e.g., <DDF version>/Deliverables/API/USDM_API.yaml)
  -o OUTPUT_FILE, --output_file OUTPUT_FILE
                        [Optional] Specifies output file Excel file. Default is ./USDM_<USDM version>_Test_Data_Template.xlsx(e.g., USDM_2.6_Test_Data_Template.xlsx)
```

