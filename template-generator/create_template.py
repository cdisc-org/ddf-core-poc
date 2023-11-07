import argparse
import re
import openpyxl
import xlsxwriter
from bs4 import BeautifulSoup, Tag


def parse_arguments():
    parser = argparse.ArgumentParser(
        description="""
        USDM Comformance Rules Test Data Template Generator.
        Reads the Enterprise Architect XMI export file containing definition
        of the USDM UML model (specified in the -x option) and the USDM
        Controlled Terminology Excel file (specified in the -c option) and
        creates an Excel template file (optionally specified in the -o option)
        for test data entry."""
    )
    parser.add_argument(
        "-x",
        "--xmi_file",
        help="USDM XMI export file "
        + "(e.g.,<DDFversion>/Deliverables/UML/USDM_UML.xmi)",
        required=True,
    )
    parser.add_argument(
        "-c",
        "--ct_file",
        help="USDM CT Excel file "
        + "(e.g., <DDF version>/Deliverables/CT/USDM_CT.xlsx)",
        required=True,
    )
    parser.add_argument(
        "-o",
        "--output_file",
        help="[Optional] Specifies output file Excel file. "
        + "Default is ./USDM_Test_Data_Template.xlsx",
        default="USDM_Test_Data_Template.xlsx",
    )
    args = parser.parse_args()
    return args


def get_properties(entName: str, cls: Tag):
    clsSheet = entName + ".xpt" if len(entName) <= 27 else entName[:27] + ".xpt"
    dsws.write_url(clsn, 0, f"internal:'{clsSheet}'!A1", string=clsSheet)
    dsws.write_row(clsn, 1, [entName, entdict[entName]["Preferred Name"]])
    ws = workbook.add_worksheet(clsSheet)
    prps = []
    dscs = []
    typs = []
    crds = []
    if cls.generalization:
        gclsId = cls.generalization["general"]
        gcls = usdmxmi.find("packagedElement", attrs={"xmi:id": gclsId})
        gclsName = gcls["name"]
        for prp in gcls.find_all("ownedAttribute", attrs={"xmi:type": "uml:Property"}):
            prps += [prp["name"]]
            if prp["name"] in entdict[entName]["Properties"]:
                dscs += [entdict[entName]["Properties"][prp["name"]]["Preferred Name"]]
            elif prp["name"] in entdict[gclsName]["Properties"]:
                print(
                    f"Using general attribute description from '{gclsName}"
                    + f".{prp['name']}' for '{entName}.{prp['name']}'"
                )
                dscs += [entdict[gclsName]["Properties"][prp["name"]]["Preferred Name"]]
            else:
                print(
                    f"No entry found in {args.ct_file} for USDM "
                    + f"attribute '{entName}.{prp['name']}'"
                )
                dscs += [None]
            attr = usdmxmi.find("attribute", attrs={"xmi:idref": str(prp["xmi:id"])})
            typs += [attr.properties["type"]]
            crds += [get_card(attr.bounds)]
    for prp in cls.find_all("ownedAttribute", attrs={"xmi:type": "uml:Property"}):
        prps += [prp["name"]]
        if prp["name"] in entdict[entName]["Properties"]:
            dscs += [entdict[entName]["Properties"][prp["name"]]["Preferred Name"]]
        else:
            print(
                f"No entry found in {args.ct_file} for USDM "
                + f"attribute '{entName}.{prp['name']}'"
            )
            dscs += [None]
        attr = usdmxmi.find("attribute", attrs={"xmi:idref": str(prp["xmi:id"])})
        typs += [attr.properties["type"]]
        crds += [get_card(attr.bounds)]
    for prpn in (
        k
        for k, v in entdict[entName]["Properties"].items()
        if v["Role"] == "Attribute"
        and not cls.find(
            "ownedAttribute", attrs={"xmi:type": "uml:Property", "name": k}
        )
    ):
        print(
            f"Attribute '{entName}.{prpn}' defined in {args.ct_file} "
            + f"does not have a matching class attribute in {args.xmi_file}"
        )
    if entName != "Study":
        prps = ["parent_entity", "parent_id", "parent_rel"] + prps
        dscs = [
            "Parent Entity Name",
            "Parent Entity Id",
            "Name of Relationship from Parent Entity",
        ] + dscs
        typs = ["String", "String", "String"] + typs
        crds = ["[1]", "[1]", "[1]"] + crds
    ws.set_column(0, len(prps), 20)
    ws.write_row(0, 0, prps, header)
    ws.write_row(1, 0, dscs, sub_header)
    ws.write_row(2, 0, typs, sub_header)
    ws.write_row(3, 0, crds, sub_header)
    # Add a blank row with defined format to prevent auto-copying of format
    # from row above.
    ws.write_row(4, 0, [None] * len(prps), normal)


def get_card(bounds: Tag) -> str:
    if bounds["lower"] == bounds["upper"]:
        return "[{}]".format(bounds["lower"])
    else:
        return "[{}..{}]".format(bounds["lower"], bounds["upper"])


args = parse_arguments()

with open(args.xmi_file) as f:
    xmidata = f.read()

usdmxmi = BeautifulSoup(xmidata, "lxml-xml")

entdict = {}

ctwb = openpyxl.load_workbook(filename=args.ct_file, data_only=True)

ctws = ctwb["DDF Entities&Attributes"]

for ctrow in ctws.iter_rows(min_row=2):
    entname = ctrow[1].value
    elrole = ctrow[2].value
    elname = ctrow[3].value
    if elrole == "Entity":
        if elname != entname:
            print(
                f"Entity Name '{entname}' does not match Logical Data Model "
                + f"Name for Entity '{elname}'"
            )
        entdict[entname] = {
            "NCI C-code": ctrow[4].value,
            "Preferred Name": ctrow[5].value,
            "Definition": ctrow[7].value,
            "Properties": {},
        }
    else:
        cref: str = None
        cref = re.search(r"^Y \((.+?)\)$", str(ctrow[8].value).strip())

        entdict[entname]["Properties"][elname] = {
            "Role": elrole,
            "NCI C-code": ctrow[4].value,
            "Preferred Name": ctrow[5].value,
            "Definition": ctrow[7].value,
            "CodelistRef": cref.group(1) if cref else None,
        }

workbook = xlsxwriter.Workbook(args.output_file)
usdmver = (
    usdmxmi.find("properties", {"name": "USDM", "type": "Logical"})
    .find_parent("diagram")
    .project["version"]
)
workbook.set_custom_property("USDM Version", str(usdmver))

header = workbook.add_format()
header.set_bold()
header.set_align("top")

sub_header = workbook.add_format()
sub_header.set_italic()
sub_header.set_bg_color("#FFFFCC")
sub_header.set_text_wrap()
sub_header.set_align("top")

normal = workbook.add_format()
normal.set_align("top")
normal.set_num_format("@")

dsws = workbook.add_worksheet("Datasets")
prps = ["Filename", "Dataset Name", "Label"]
dsws.set_column(0, len(prps), 30)
dsws.write_row(0, 0, prps, header)

clsn = 0

for entName in entdict.keys():
    cls = usdmxmi.find(
        "packagedElement", attrs={"xmi:type": "uml:Class", "name": entName}
    )
    if cls:
        clsn += 1
        get_properties(entName, cls)
    else:
        print(
            f"Entity '{entName}' defined in {args.ct_file} does not have a "
            + f"matching class in {args.xmi_file}"
        )

for cls in (
    x
    for x in usdmxmi.find_all("packagedElement", attrs={"xmi:type": "uml:Class"})
    if x["name"] not in entdict
):
    print(
        f"USDM class '{cls['name']}' defined in {args.xmi_file} does not "
        + f"have a matching Entity in {args.ct_file}"
    )

workbook.close()
