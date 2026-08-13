import argparse
import html
import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree


FAKE_VALUES = {
    "full_name": ["John Doe", "Peter Parker", "Ananya Sharma", "Maya Rao"],
    "email": ["john.doe@example.com", "peter.parker@example.com", "ananya.sharma@example.com"],
    "phone": ["+91 1234567890", "+91 9876501234", "+1 202 555 0100"],
    "company": ["Acme Private Limited", "Northstar Technologies Ltd", "Bluebird Capital LLP"],
    "address": [
        "221B Baker Street, London NW1 6XE",
        "742 Evergreen Terrace, Springfield 49007",
        "12 Park Avenue, Mumbai 400001",
    ],
    "ssn": ["321-54-9876", "111-22-3333", "222-33-4444"],
    "credit_card": ["4000 0000 0000 1000", "4000 0000 0000 1001"],
    "dob": ["01/01/1990", "15/08/1988", "1975-12-10"],
    "ip_address": ["203.0.113.10", "198.51.100.25", "192.0.2.77"],
}


PII_PATTERNS = {
    "email": r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b",
    "phone": r"(?<!\w)(?:\+?\d{1,3}[-. ]?)?(?:\(?\d{2,4}\)?[-. ]?)?\d{3,5}[-. ]?\d{4}(?!\w)",
    "ssn": r"\b\d{3}-\d{2}-\d{4}\b",
    "credit_card": r"\b(?:\d[ -]*?){13,19}\b",
    "ip_address": r"\b(?:(?:25[0-5]|2[0-4]\d|1?\d?\d)\.){3}(?:25[0-5]|2[0-4]\d|1?\d?\d)\b",
    "dob": r"\b(?:DOB|Date of Birth|Born)[: -]*((?:\d{1,2}[/-]\d{1,2}[/-]\d{2,4})|(?:\d{4}-\d{2}-\d{2}))\b",
    "company": r"\b([A-Z][A-Za-z&.'-]*(?: [A-Z][A-Za-z&.'-]*){0,5} (?:Limited|Ltd\.?|Private Limited|Pvt\.? Ltd\.?|LLP|Inc\.?|Corporation|Corp\.?|Company|Co\.))\b",
    "address": r"\b(?:Address|Mailing Address|Physical Address)[: -]+([^\r\n]+)",
    "full_name": r"\b(?:Name|Customer|Employee|Applicant|Director|Contact|Client|User)[: -]+([A-Z][a-z]+(?: [A-Z][a-z]+){1,3})\b",
}


def get_basic_eda(text):
    lines = text.splitlines()
    words = text.split()
    return {
        "total_characters": len(text),
        "total_lines": len(lines),
        "total_words": len(words),
    }


def read_text_from_docx(file_path):
    # A .docx file stores text inside word/document.xml.
    with zipfile.ZipFile(file_path) as docx_file:
        xml_data = docx_file.read("word/document.xml")

    root = ElementTree.fromstring(xml_data)
    namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    paragraphs = []

    for paragraph in root.findall(".//w:p", namespace):
        words = []
        for text_node in paragraph.findall(".//w:t", namespace):
            words.append(text_node.text or "")
        if words:
            paragraphs.append("".join(words))

    return "\n".join(paragraphs)


def read_input_file(file_path):
    file_path = Path(file_path)
    if file_path.suffix.lower() == ".docx":
        return read_text_from_docx(file_path)
    return file_path.read_text(encoding="utf-8")


def write_text_to_docx(text, output_path):
    paragraphs = text.splitlines() or [""]
    paragraph_xml = ""

    for paragraph in paragraphs:
        paragraph_xml += (
            '<w:p><w:r><w:t xml:space="preserve">'
            + html.escape(paragraph)
            + "</w:t></w:r></w:p>"
        )

    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        "<w:body>"
        + paragraph_xml
        + "<w:sectPr/></w:body></w:document>"
    )

    content_types_xml = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        "</Types>"
    )

    relationships_xml = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/></Relationships>'
    )

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as docx_file:
        docx_file.writestr("[Content_Types].xml", content_types_xml)
        docx_file.writestr("_rels/.rels", relationships_xml)
        docx_file.writestr("word/document.xml", document_xml)


def is_valid_credit_card(number):
    # Luhn check reduces false positives for long normal numbers.
    digits = re.sub(r"\D", "", number)
    if len(digits) < 13 or len(digits) > 19:
        return False

    total = 0
    for position, digit in enumerate(digits[::-1]):
        value = int(digit)
        if position % 2 == 1:
            value *= 2
            if value > 9:
                value -= 9
        total += value

    return total % 10 == 0


def find_pii(text):
    found_items = []

    for pii_type, pattern in PII_PATTERNS.items():
        for match in re.finditer(pattern, text, flags=re.IGNORECASE):
            if match.lastindex:
                value = match.group(1)
                start = match.start(1)
                end = match.end(1)
            else:
                value = match.group(0)
                start = match.start()
                end = match.end()

            if pii_type == "phone" and len(re.sub(r"\D", "", value)) < 10:
                continue
            if pii_type == "credit_card" and not is_valid_credit_card(value):
                continue

            found_items.append({
                "type": pii_type,
                "start": start,
                "end": end,
                "value": value.strip(),
            })

    return remove_duplicate_or_overlapping_items(found_items)


def remove_duplicate_or_overlapping_items(found_items):
    priority = {
        "email": 1,
        "ssn": 2,
        "credit_card": 3,
        "ip_address": 4,
        "dob": 5,
        "phone": 6,
        "address": 7,
        "company": 8,
        "full_name": 9,
    }

    sorted_items = sorted(
        found_items,
        key=lambda item: (priority[item["type"]], -(item["end"] - item["start"])),
    )

    final_items = []
    for item in sorted_items:
        overlaps = False
        for kept_item in final_items:
            if item["start"] < kept_item["end"] and item["end"] > kept_item["start"]:
                overlaps = True
                break
        if not overlaps:
            final_items.append(item)

    return sorted(final_items, key=lambda item: item["start"])


def get_fake_value(pii_type, original_value, used_fake_values):
    key = pii_type + "|" + original_value
    if key in used_fake_values:
        return used_fake_values[key]

    options = FAKE_VALUES[pii_type]
    used_count = len([key for key in used_fake_values if key.startswith(pii_type + "|")])
    fake_value = options[used_count % len(options)]
    used_fake_values[key] = fake_value
    return fake_value


def redact_text(text):
    found_items = find_pii(text)
    used_fake_values = {}
    mapping = []
    redacted_parts = []
    last_position = 0

    for item in found_items:
        fake_value = get_fake_value(item["type"], item["value"], used_fake_values)
        redacted_parts.append(text[last_position:item["start"]])
        redacted_parts.append(fake_value)
        mapping.append({
            "type": item["type"],
            "original": item["value"],
            "replacement": fake_value,
        })
        last_position = item["end"]

    redacted_parts.append(text[last_position:])
    return "".join(redacted_parts), mapping


def remove_labels_from_test_file(labelled_text):
    label_pattern = re.compile(r"\[\[(?P<type>[a-z_]+):(?P<value>.*?)\]\]")
    clean_text_parts = []
    correct_answers = []
    last_position = 0
    clean_text_position = 0

    for match in label_pattern.finditer(labelled_text):
        before_label = labelled_text[last_position:match.start()]
        clean_text_parts.append(before_label)
        clean_text_position += len(before_label)

        pii_type = match.group("type")
        value = match.group("value")
        clean_text_parts.append(value)
        correct_answers.append({
            "type": pii_type,
            "start": clean_text_position,
            "end": clean_text_position + len(value),
        })

        clean_text_position += len(value)
        last_position = match.end()

    clean_text_parts.append(labelled_text[last_position:])
    return "".join(clean_text_parts), correct_answers


def write_evaluation_report(test_file_path, report_path):
    labelled_text = Path(test_file_path).read_text(encoding="utf-8")
    clean_text, correct_answers = remove_labels_from_test_file(labelled_text)
    detected_items = find_pii(clean_text)

    expected = {(item["type"], item["start"], item["end"]) for item in correct_answers}
    actual = {(item["type"], item["start"], item["end"]) for item in detected_items}

    true_positive = len(expected & actual)
    false_positive = len(actual - expected)
    false_negative = len(expected - actual)

    precision = true_positive / (true_positive + false_positive) if actual else 0
    recall = true_positive / (true_positive + false_negative) if expected else 0
    total_checked = true_positive + false_positive + false_negative
    accuracy = true_positive / total_checked if total_checked else 0

    type_lines = []
    type_names = sorted({item["type"] for item in correct_answers + detected_items})
    for type_name in type_names:
        expected_for_type = {item for item in expected if item[0] == type_name}
        actual_for_type = {item for item in actual if item[0] == type_name}
        type_tp = len(expected_for_type & actual_for_type)
        type_fp = len(actual_for_type - expected_for_type)
        type_fn = len(expected_for_type - actual_for_type)
        type_precision = type_tp / (type_tp + type_fp) if type_tp + type_fp else 0
        type_recall = type_tp / (type_tp + type_fn) if type_tp + type_fn else 0
        type_lines.append(
            f"- {type_name}: precision {type_precision:.2f}, recall {type_recall:.2f}, correct {type_tp}"
        )

    report = (
        "# Evaluation Report\n\n"
        f"Test file: `{Path(test_file_path).name}`\n\n"
        f"Total labelled PII values: {len(correct_answers)}\n"
        f"Total detected PII values: {len(detected_items)}\n\n"
        f"- Accuracy: {accuracy:.2f}\n"
        f"- Precision: {precision:.2f}\n"
        f"- Recall: {recall:.2f}\n"
        f"- Correct redactions: {true_positive}\n"
        f"- Wrong redactions: {false_positive}\n"
        f"- Missed PII values: {false_negative}\n\n"
        "Score by PII type:\n\n"
        + "\n".join(type_lines)
        + "\n\n"
        "The script works better on fixed patterns like emails, phones, SSNs, cards, and IPs. "
        "It misses more names, companies, and addresses because they can be written in many styles.\n"
    )

    Path(report_path).write_text(report, encoding="utf-8")
    return accuracy, precision, recall


def make_large_test_file(output_path, total_records):
    first_names = ["Rashi", "Rohan", "Ananya", "Arjun", "Maya", "Priya", "Karan", "Neha"]
    last_names = ["Patil", "Dey", "Sharma", "Mehta", "Rao", "Nair", "Kapoor", "Iyer"]
    companies = [
        "Bright Future Technologies Private Limited",
        "Green Valley Capital LLP",
        "Sunrise Finance Ltd",
        "Blue Ocean Services Inc",
        "Apex Retail",
        "Nova Healthcare",
    ]
    streets = ["Park Street", "Market Road", "MG Avenue", "Lake View Lane", "Sector 18", "Flat 4B, Pearl Residency"]
    cities = ["Pune", "Mumbai", "Delhi", "Bengaluru"]
    cards = ["4111 1111 1111 1111", "5555 5555 5555 4444"]
    all_records = []

    for number in range(total_records):
        first_name = first_names[number % len(first_names)]
        last_name = last_names[number % len(last_names)]
        full_name = first_name + " " + last_name
        email = first_name.lower() + "." + last_name.lower() + str(number) + "@gmail.com"
        phone = "+91 9" + str(800000000 + number)[-9:]
        company = companies[number % len(companies)]
        address = (
            str(10 + number)
            + " "
            + streets[number % len(streets)]
            + ", "
            + cities[number % len(cities)]
            + " "
            + str(400000 + number)
        )
        dob = f"{(number % 28) + 1:02d}/01/{1980 + (number % 20)}"
        ssn = f"{100 + (number % 800)}-{10 + (number % 80):02d}-{1000 + (number % 8000):04d}"
        card = cards[number % len(cards)]
        ip = f"10.{number % 255}.{(number * 2) % 255}.{(number * 3) % 255}"
        style = number % 5

        if style == 0:
            record = (
                f"Ticket {10000 + number}\n"
                f"Customer: [[full_name:{full_name}]]\n"
                f"Email: [[email:{email}]]\n"
                f"Phone: [[phone:{phone}]]\n"
                f"Company: [[company:{company}]]\n"
                f"Address: [[address:{address}]]\n"
                f"DOB: [[dob:{dob}]]\n"
                f"SSN: [[ssn:{ssn}]]\n"
                f"Card: [[credit_card:{card}]]\n"
                f"IP: [[ip_address:{ip}]]\n"
                f"Order number kept as normal text: {70000 + number}\n"
            )
        elif style == 1:
            record = (
                f"Ticket {10000 + number}\n"
                f"Caller [[full_name:{full_name}]] reported issue from [[company:{company}]].\n"
                f"Reach them at [[email:{email}]] or [[phone:{phone}]].\n"
                f"They stay near [[address:{address}]].\n"
                f"Birth date is [[dob:{dob}]]. Network IP was [[ip_address:{ip}]].\n"
                f"SSN: [[ssn:{ssn}]], card used [[credit_card:{card}]].\n"
                f"Invoice number is {900000 + number} and should not be treated as private.\n"
            )
        elif style == 2:
            record = (
                f"Ticket {10000 + number}\n"
                f"Name: [[full_name:{full_name}]]\n"
                f"mail id - [[email:{email}]]\n"
                f"mobile no. [[phone:{phone}]]\n"
                f"works for [[company:{company}]]\n"
                f"office location [[address:{address}]]\n"
                f"Born [[dob:{dob}]]\n"
                f"social security [[ssn:{ssn}]]\n"
                f"payment card [[credit_card:{card}]]\n"
                f"login address [[ip_address:{ip}]]\n"
            )
        elif style == 3:
            record = (
                f"Ticket {10000 + number}\n"
                f"Applicant: [[full_name:{full_name}]]\n"
                f"Email: [[email:{email}]]\n"
                f"Contact number: [[phone:{phone}]]\n"
                f"Employer: [[company:{company}]]\n"
                f"Mailing Address: [[address:{address}]]\n"
                f"Date of Birth: [[dob:{dob}]]\n"
                f"SSN: [[ssn:{ssn}]]\n"
                f"Credit Card: [[credit_card:{card}]]\n"
                f"IP: [[ip_address:{ip}]]\n"
                f"Reference id 4111111111111111 is only a sample test number.\n"
            )
        else:
            record = (
                f"Ticket {10000 + number}\n"
                f"Spoke with [[full_name:{full_name}]] from [[company:{company}]] today.\n"
                f"Email address was [[email:{email}]]. Phone was [[phone:{phone}]].\n"
                f"Customer wrote home as [[address:{address}]].\n"
                f"DOB: [[dob:{dob}]]. SSN: [[ssn:{ssn}]].\n"
                f"Card number [[credit_card:{card}]]. Last login IP [[ip_address:{ip}]].\n"
                f"Ticket code {10000 + number} and order code {70000 + number} should remain.\n"
            )

        all_records.append(record)

    Path(output_path).write_text("\n".join(all_records), encoding="utf-8")


def main():
    parser = argparse.ArgumentParser(description="Redact private information from a file.")
    parser.add_argument("input_file", nargs="?", default="sample_ticket_log.txt")
    parser.add_argument("-o", "--output", default="redacted_output.docx")
    parser.add_argument("--mapping", default="redaction_mapping.json")
    parser.add_argument("--test-file", default="synthetic_eval_labeled_realistic_1200.txt")
    parser.add_argument("--make-test-data", type=int, default=1200)
    args = parser.parse_args()

    input_text = read_input_file(args.input_file)
    eda_summary = get_basic_eda(input_text)
    redacted_text, mapping = redact_text(input_text)

    write_text_to_docx(redacted_text, args.output)
    Path(args.mapping).write_text(json.dumps(mapping, indent=2), encoding="utf-8")
    Path("eda_summary.json").write_text(json.dumps(eda_summary, indent=2), encoding="utf-8")

    test_file = Path(args.test_file)
    if args.make_test_data > 0:
        make_large_test_file(test_file, args.make_test_data)

    if test_file.exists():
        accuracy, precision, recall = write_evaluation_report(test_file, "evaluation_report.md")
        print("Evaluation finished:")
        print(f"Accuracy: {accuracy:.2f}")
        print(f"Precision: {precision:.2f}")
        print(f"Recall: {recall:.2f}")

    print(f"Redacted file created: {args.output}")
    print(f"Replacement list created: {args.mapping}")
    print("EDA summary created: eda_summary.json")


if __name__ == "__main__":
    main()
