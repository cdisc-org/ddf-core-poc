import argparse
import re
import yaml
import openpyxl
import xlsxwriter
from bs4 import BeautifulSoup, Tag
from copy import deepcopy
import inflect


def parse_arguments():
    parser = argparse.ArgumentParser(
        description="""
        USDM Comformance Rules Test Data Template Generator.
        Reads the Enterprise Architect XMI export file containing definition
        of the USDM UML model (specified in the -x option), the USDM
        Controlled Terminology Excel file (specified in the -c option), and the
        USDM API specification in YAML format (specified in the -a option) and
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
        "-a",
        "--api_spec",
        help="USDM API specification YAML file. "
        + "(e.g., <DDF version>/Deliverables/API/USDM_API.yaml)",
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


def replace_deep(data, a, b):
    if isinstance(data, str):
        return data.replace(a, b)
    elif isinstance(data, dict):
        return {k: replace_deep(v, a, b) for k, v in data.items()}
    elif isinstance(data, list):
        return [replace_deep(v, a, b) for v in data]
    else:
        return data


def get_properties(
    entName: str, cls: Tag, prps: list, prefix: str = None, lclsName: str = None
):
    if cls.generalization:
        gclsId = cls.generalization["general"]
        gcls = usdmxmi.find("packagedElement", attrs={"xmi:id": gclsId})
        gclsName = gcls["name"]
        get_properties(entName, gcls, prps, prefix, gclsName)
    for prp in (
        x
        for x in cls.find_all("ownedAttribute", attrs={"xmi:type": "uml:Property"})
        if x.has_attr("name")
    ):
        prps[0] += [".".join((prefix, prp["name"])) if prefix else prp["name"]]
        prps[1] += [get_description(entName, prp, prps, prefix, lclsName)]
        attr = usdmxmi.find("attribute", attrs={"xmi:idref": str(prp["xmi:id"])})
        prps[2] += [attr.properties["type"]]
        prps[3] += [get_card(attr.bounds)]
    for lnk in (
        x.find_parent("connector")
        for x in usdmxmi.find_all("source", attrs={"xmi:idref": cls["xmi:id"]})
        if x.find_parent("connector").has_attr("name")
    ):
        if lnk["name"] in entdict[lclsName if prefix else entName]["Properties"]:
            apiattr = (
                entdict[lclsName if prefix else entName]["Properties"][lnk["name"]][
                    "apiattr"
                ]
                if "apiattr"
                in entdict[lclsName if prefix else entName]["Properties"][lnk["name"]]
                else None
            )
            if apiattr is None or apiattr == lnk["name"]:
                if lnk.target.type["multiplicity"].endswith("1"):
                    lcls = usdmxmi.find(
                        "packagedElement", attrs={"xmi:id": lnk.target["xmi:idref"]}
                    )
                    if prefix and lnk["name"] in prefix.split("."):
                        print(
                            f"Circular relationship found: {lnk['name']} found in "
                            + prefix
                        )
                    else:
                        get_properties(
                            entName,
                            lcls,
                            prps,
                            ".".join((prefix, lnk["name"])) if prefix else lnk["name"],
                            lcls["name"],
                        )
            else:
                prps[0] += [".".join((prefix, apiattr)) if prefix else apiattr]
                prps[1] += [
                    entdict[lclsName if prefix else entName]["Properties"][lnk["name"]][
                        "Preferred Name"
                    ]
                ]
                prps[2] += ["{}.id".format(lnk.target.model["name"])]
                prps[3] += ["[{}]".format(lnk.target.type["multiplicity"])]
        elif not prefix:
            print(
                f"Relationship '{entName}.{lnk['name']}' defined in {args.xmi_file} "
                + f"does not have a matching entry in {args.ct_file}"
            )


def get_description(
    entName: str, prp: Tag, prps: list, prefix: str = None, lclsName: str = None
) -> str:
    if prefix:
        # If we're following a link, use the linked class attribute
        # if it's there, otherwise (usually for id) try to use the link
        # name as the attribute.
        if prp["name"] in entdict[lclsName]["Properties"]:
            return entdict[lclsName]["Properties"][prp["name"]]["Preferred Name"]
        elif "".join(prefix.split(".")[-1:]) in entdict[entName]["Properties"]:
            return entdict[entName]["Properties"]["".join(prefix.split(".")[-1:])][
                "Preferred Name"
            ]
        else:
            return None
    else:
        # If this is a normal class, just use it's attribute.
        if prp["name"] in entdict[entName]["Properties"]:
            return entdict[entName]["Properties"][prp["name"]]["Preferred Name"]
        else:
            if prp["name"] != "id":
                print(
                    f"No entry found in {args.ct_file} for USDM "
                    + f"attribute '{entName}.{prp['name']}'"
                )
            return None


def get_card(bounds: Tag) -> str:
    if bounds["lower"] == bounds["upper"]:
        return "[{}]".format(bounds["lower"])
    else:
        return "[{}..{}]".format(bounds["lower"], bounds["upper"])


def get_apiattr(entName: str, elname: str, elrole: str) -> str:
    if elname in apidict[entName]:
        return elname
    else:
        if elrole == "Attribute":
            print(f"No entry found in {args.api_spec} for '{entName}.{elname}'")
            return None
        else:
            elprts = re.findall(r"([A-Z]?[a-z]+)", elname.strip())
            if (
                inflect.singular_noun(elprts[-1]) is False
                or inflect.singular_noun(elprts[-1]) == elprts[-1]
            ):
                if elname + "Id" in apidict[entName]:
                    return elname + "Id"
                else:
                    print(
                        f"No entry found in {args.api_spec} for '{entName}."
                        + f"{elname}' or '{entName}.{elname}Id'"
                    )
            else:
                if len(elprts) == 1:
                    if inflect.singular_noun(elname) + "Ids" in apidict[entName]:
                        return inflect.singular_noun(elname) + "Ids"
                    elif elname + "Ids" in apidict[entName]:
                        print(
                            f"Using '{entName}.{elname}Ids' instead of "
                            + f"'{entName}.{inflect.singular_noun(elname)}Ids' "
                            + "for API attribute"
                        )
                        return elname + "Ids"
                    else:
                        print(
                            f"No entry found in {args.api_spec} for "
                            + f"'{entName}.{elname}', '{entName}."
                            + f"{inflect.singular_noun(elname)}Ids' or "
                            + f"'{entName}.{elname}Ids'"
                        )
                        return None
                else:
                    if (
                        "".join(elprts[:-1]) + inflect.singular_noun(elprts[-1]) + "Ids"
                        in apidict[entName]
                    ):
                        return (
                            "".join(elprts[:-1])
                            + inflect.singular_noun(elprts[-1])
                            + "Ids"
                        )
                    elif elname + "Ids" in apidict[entName]:
                        print(
                            f"Using '{entName}.{elname}Ids' instead of "
                            + f"'{entName}."
                            + "".join(elprts[:-1])
                            + inflect.singular_noun(elprts[-1])
                            + "Ids' for API attribute"
                        )
                        return elname + "Ids"
                    else:
                        print(
                            f"No entry found in {args.api_spec} for '{entName}."
                            + f"{elname}', '{entName}."
                            + "".join(elprts[:-1])
                            + inflect.singular_noun(elprts[-1])
                            + f"Ids' or '{entName}.{elname}Ids'"
                        )
                        return None


args = parse_arguments()

with open(args.xmi_file) as f:
    xmidata = f.read()

usdmxmi = BeautifulSoup(xmidata, "lxml-xml")

with open(args.api_spec, "r") as f:
    apispec = yaml.safe_load(f)

apidict = {}

for k, v in apispec["components"]["schemas"].items():
    if "-" in k:
        if k.endswith("-Input"):
            cname = "".join(k.split("-")[:1])
            if cname + "-Output" in apispec["components"]["schemas"]:
                vo = apispec["components"]["schemas"][cname + "-Output"]
                if v != replace_deep(
                    vo,
                    "Output",
                    "Input",
                ):
                    print(f"API Input/Output definitions do not match for {cname}")
                    print(f"{cname}-Input : {v}")
                    print(f"{cname}-Output: {vo}")
            else:
                print(f"No corresponding API Output definition for {k}")
    else:
        cname = k

    apidict[cname] = deepcopy(v["properties"])

inflect = inflect.engine()
inflect.defnoun("previous", "previous")

entdict = {}

ctwb = openpyxl.load_workbook(filename=args.ct_file, data_only=True)

ctws = ctwb["DDF Entities&Attributes"]

for ctrow in ctws.iter_rows(min_row=2):
    entName = ctrow[1].value
    elrole = ctrow[2].value
    elname = ctrow[3].value
    if elrole == "Entity":
        if elname != entName:
            print(
                f"Entity Name '{entName}' does not match Logical Data Model "
                + f"Name for Entity '{elname}'"
            )
        entdict[entName] = {
            "NCI C-code": ctrow[4].value,
            "Preferred Name": ctrow[5].value,
            "Definition": ctrow[7].value,
            "Properties": {},
        }
    else:
        cref: str = None
        cref = re.search(r"^Y \((.+?)\)$", str(ctrow[8].value).strip())

        entdict[entName]["Properties"][elname] = {
            "name": elname,
            "Role": elrole,
            "NCI C-code": ctrow[4].value,
            "Preferred Name": ctrow[5].value,
            "Definition": ctrow[7].value,
            "CodelistRef": cref.group(1) if cref else None,
        }

for entName, entDef in entdict.items():
    cls = usdmxmi.find(
        "packagedElement", attrs={"xmi:type": "uml:Class", "name": entName}
    )
    if cls.generalization:
        gclsName = usdmxmi.find(
            "packagedElement",
            attrs={
                "xmi:type": "uml:Class",
                "xmi:id": cls.generalization["general"],
            },
        )["name"]
        if gclsName in entdict:
            for gprp in entdict[gclsName]["Properties"].keys():
                if gprp not in entDef["Properties"]:
                    entDef["Properties"][gprp] = deepcopy(
                        entdict[gclsName]["Properties"][gprp]
                    )
                    print(
                        f"Using general '{gclsName}.{gprp}' "
                        + entDef["Properties"][gprp]["Role"].lower()
                        + f" in '{entName}' specialization"
                    )
    if entName in apidict:
        for prpName, prpDef in entDef["Properties"].items():
            prpDef["apiattr"] = get_apiattr(entName, prpName, prpDef["Role"])

for abscls in (
    x["name"]
    for x in usdmxmi.find_all("element", attrs={"xmi:type": "uml:Class"})
    if x.properties["isAbstract"] == "true" or x["name"] not in apidict
):
    entdict.pop(abscls)
    print(f"Excluding abstract class '{abscls}'")


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
dsprps = ["Filename", "Dataset Name", "Label"]
dsws.set_column(0, len(dsprps), 30)
dsws.write_row(0, 0, dsprps, header)

clsn = 0

for entName in entdict.keys():
    cls = usdmxmi.find(
        "packagedElement", attrs={"xmi:type": "uml:Class", "name": entName}
    )
    if cls:
        clsn += 1
        clsSheet = entName + ".xpt" if len(entName) <= 27 else entName[:27] + ".xpt"
        dsws.write_url(clsn, 0, f"internal:'{clsSheet}'!A1", string=clsSheet)
        dsws.write_row(clsn, 1, [entName, entdict[entName]["Preferred Name"]])
        ws = workbook.add_worksheet(clsSheet)
        prps = [[], [], [], []]
        get_properties(entName, cls, prps)
        for prpv in (
            v
            for k, v in entdict[entName]["Properties"].items()
            if not (
                cls.find(
                    "ownedAttribute", attrs={"xmi:type": "uml:Property", "name": k}
                )
                or any(
                    x
                    for x in usdmxmi.find_all(
                        "source", attrs={"xmi:idref": cls["xmi:id"]}
                    )
                    if x.find_parent("connector").has_attr("name")
                    and x.find_parent("connector")["name"] == k
                )
                or (
                    cls.generalization
                    and (
                        usdmxmi.find(
                            "packagedElement",
                            attrs={
                                "xmi:type": "uml:Class",
                                "xmi:id": cls.generalization["general"],
                            },
                        ).find(
                            "ownedAttribute",
                            attrs={"xmi:type": "uml:Property", "name": k},
                        )
                        or any(
                            x
                            for x in usdmxmi.find_all(
                                "source",
                                attrs={"xmi:idref": cls.generalization["general"]},
                            )
                            if x.find_parent("connector").has_attr("name")
                            and x.find_parent("connector")["name"] == k
                        )
                    )
                )
            )
        ):
            print(
                f"{prpv['Role']} '{entName}.{prpv['name']}' defined in {args.ct_file} "
                + "does not have a matching attribute or relationship in "
                + args.xmi_file
            )
        if entName != "Study":
            prps[0] = ["parent_entity", "parent_id", "parent_rel"] + prps[0]
            prps[1] = [
                "Parent Entity Name",
                "Parent Entity Id",
                "Name of Relationship from Parent Entity",
            ] + prps[1]
            prps[2] = ["String", "String", "String"] + prps[2]
            prps[3] = ["[1]", "[1]", "[1]"] + prps[3]
        ws.set_column(0, len(prps[0]), 20)
        ws.write_row(0, 0, prps[0], header)
        ws.write_row(1, 0, prps[1], sub_header)
        ws.write_row(2, 0, prps[2], sub_header)
        ws.write_row(3, 0, prps[3], sub_header)
        # Add a blank row with defined format to prevent auto-copying of format
        # from row above.
        ws.write_row(4, 0, [None] * len(prps[0]), normal)
    else:
        print(
            f"Entity '{entName}' defined in {args.ct_file} does not have a "
            + f"matching class in {args.xmi_file}"
        )

for cls in (
    x
    for x in usdmxmi.find_all("packagedElement", attrs={"xmi:type": "uml:Class"})
    if x["name"] not in entdict
    and usdmxmi.find("element", attrs={"xmi:idref": x["xmi:id"]}).properties[
        "isAbstract"
    ]
    == "false"
    and x["name"] in apidict
):
    print(
        f"USDM class '{cls['name']}' defined in {args.xmi_file} does not "
        + f"have a matching Entity in {args.ct_file}"
    )

workbook.close()
