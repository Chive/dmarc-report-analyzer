import sys
from datetime import datetime
from pathlib import Path
import xml.etree.ElementTree as ET
import xlsxwriter
import shutil
import os
import tempfile
import gzip

RESULTS_FILE = "results.xlsx"
TRUSTED_SOURCE_IP = None  # add your trusted source IP here


try:
    source_folder = Path(sys.argv[1])
except:
    print("Usage: python process.py <path/to/reports>")
    exit(1)

data = []


def get_archives(path):
    patterns = ("zip", "gz")
    for pattern in patterns:
        for filename in path.glob(f"*.{pattern}"):
            yield filename


print("Looking for zips to extract")
for filename in source_folder.glob("*.zip"):
    with tempfile.TemporaryDirectory() as tmpdirname:
        shutil.unpack_archive(filename, tmpdirname)
        print(f" extracted {filename} to {tmpdirname}")
        extracted_files = list(Path(tmpdirname).glob("*.xml"))
        for extracted_file in extracted_files:
            print(f" Moving {extracted_file} to {source_folder}")
            shutil.copy(str(extracted_file), source_folder)
            os.remove(str(extracted_file))

        if len(extracted_files) >= 1:
            print(f" Removing archive {filename}")
            os.remove(filename)


print("Looking for gzips to extract")
for filename in source_folder.glob("*.gz"):
    with gzip.open(filename, "rb") as archive_fh:
        outfile_name = filename.with_suffix("")
        with open(outfile_name, "wb") as out_fh:
            shutil.copyfileobj(archive_fh, out_fh)
            print(f" extracted {filename} to {source_folder}")
        print(f" Removing archive {filename}")
        os.remove(filename)

print("Processing reports")
for filename in source_folder.glob("*.xml"):
    print(f" Processing {filename}")
    tree = ET.parse(filename)
    root = tree.getroot()
    metadata = root.find("report_metadata")
    date_range = metadata.find("date_range")

    policy_published = root.find("policy_published")

    for record in root.findall("record"):
        row = record.find("row")
        policy_evaluated = row.find("policy_evaluated")
        disposition = policy_evaluated.find("disposition")
        dkim_alignment_result = policy_evaluated.find("dkim").text
        spf_alignment_result = policy_evaluated.find("spf").text

        auth_results = record.find("auth_results")
        try:
            dkim_auth_result = auth_results.find("dkim").find("result").text
        except AttributeError:
            dkim_auth_result = "fail"

        try:
            spf_auth_result = auth_results.find("spf").find("result").text
        except AttributeError:
            spf_auth_result = "fail"

        if spf_auth_result == "fail" or dkim_auth_result == "fail":
            dmarc_result = "fail"
        else:
            dmarc_result = "pass"

        data.append(
            {
                "provider": metadata.find("org_name").text,
                "dates": (
                    datetime.fromtimestamp(int(date_range.find("begin").text)),
                    datetime.fromtimestamp(int(date_range.find("end").text)),
                ),
                "source": row.find("source_ip").text,
                "volume": row.find("count").text,
                "dmarc": dmarc_result,
                "spf": {
                    "auth": spf_auth_result,
                    "align": spf_alignment_result,
                },
                "dkim": {
                    "auth": dkim_auth_result,
                    "align": dkim_alignment_result,
                },
            }
        )


data = sorted(data, key=lambda x: (x["dates"][0], x["dates"][1], x["provider"]))

results_path = source_folder / RESULTS_FILE
workbook = xlsxwriter.Workbook(results_path)
workbook.formats[0].set_font_size(9)
worksheet = workbook.add_worksheet()

header_format = workbook.add_format(
    {
        "bold": 1,
        "align": "center",
        "valign": "vcenter",
        "font_size": 9,
    }
)

worksheet.merge_range("A1:C1", "Report", header_format)
worksheet.write_row("A2", ["Provider", "Date Start", "Date End"], header_format)

worksheet.merge_range("D1:E1", "Source", header_format)
worksheet.write_row("D2", ["IP Address", "Email Volume"], header_format)

worksheet.merge_range("F1:F2", "DMARC", header_format)

worksheet.merge_range("G1:H1", "SPF", header_format)
worksheet.write_row("G2", ["Auth", "Align"], header_format)

worksheet.merge_range("I1:J1", "DKIM", header_format)
worksheet.write_row("I2", ["Auth", "Align"], header_format)

border_top_format = workbook.add_format({"top": 1})
worksheet.conditional_format(
    "A3:J3", {"type": "no_blanks", "format": border_top_format}
)

datetime_format = workbook.add_format(
    {"num_format": "dd.mm.yyyy HH:MM", "font_size": 9}
)

start = 2
for i, rowdata in enumerate(data):
    column = start + i
    worksheet.write(column, 0, rowdata["provider"])
    worksheet.write_row(column, 1, rowdata["dates"], datetime_format)
    worksheet.write_row(
        column,
        3,
        (
            rowdata["source"],
            rowdata["volume"],
            rowdata["dmarc"],
            rowdata["spf"]["auth"],
            rowdata["spf"]["align"],
            rowdata["dkim"]["auth"],
            rowdata["dkim"]["align"],
        ),
    )

    fail_format = workbook.add_format({"bg_color": "#ffcccc", "font_color": "#ca1c19"})
    worksheet.conditional_format(
        f"F{start+1}:J{column+1}",
        {"type": "cell", "criteria": "!=", "value": '"pass"', "format": fail_format},
    )
    warn_format = workbook.add_format({"bg_color": "#f9facc", "font_color": "#9d6409"})
    worksheet.conditional_format(
        f"D{start+1}:D{column+1}",
        {
            "type": "cell",
            "criteria": "!=",
            "value": f'"{TRUSTED_SOURCE_IP}"',
            "format": warn_format,
        },
    )

workbook.close()

print(f"Results saved to {results_path}")
