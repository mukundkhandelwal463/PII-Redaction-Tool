# PII Redaction Tool

This project redacts private information from a document. The input file may contain names, emails, phone numbers, company names, addresses, SSNs, credit card numbers, dates of birth, and IP addresses. The script replaces these real values with fake values and creates a safer Word file.

Example:

```text
Rashi Patil -> John Doe
rashhi.patil@gmail.com -> john.doe@example.com
+91 9876543210 -> +91 1234567890
```

## What I Did

I wrote `redactor.py` in Python. It reads a `.txt` or simple `.docx` file, finds PII using regex patterns, replaces the detected values with fake values, and saves the result as `redacted_output.docx`.

I also added simple EDA and evaluation. EDA gives basic counts like words, lines, and characters. Evaluation checks how well the script detects PII.

## Why This Approach

I used regex because many PII types have fixed patterns. Emails, SSNs, IP addresses, phone numbers, and credit cards can be found well using rules.

Names, companies, and addresses are harder because they can be written in many ways. In a real production system, I would improve this by adding an NLP/NER tool like spaCy or Presidio.

## How to Run

```bash
python redactor.py sample_ticket_log.txt -o redacted_output.docx
```

Main output files:

- `redacted_output.docx`: redacted Word file
- `redaction_mapping.json`: original values and fake replacements
- `eda_summary.json`: basic EDA summary
- `evaluation_report.md`: accuracy, precision, and recall report

## Function Summary

- `get_basic_eda(text)`: counts characters, lines, and words.
- `read_text_from_docx(file_path)`: extracts text from a `.docx` file.
- `read_input_file(file_path)`: reads either `.txt` or `.docx` input.
- `write_text_to_docx(text, output_path)`: writes redacted text into a Word file.
- `is_valid_credit_card(number)`: checks card numbers using the Luhn check.
- `find_pii(text)`: detects PII using regex patterns.
- `remove_duplicate_or_overlapping_items(found_items)`: removes duplicate or overlapping detections.
- `get_fake_value(pii_type, original_value, used_fake_values)`: gives a fake value and keeps replacements consistent.
- `redact_text(text)`: replaces detected PII with fake values.
- `remove_labels_from_test_file(labelled_text)`: removes labels from test data for evaluation.
- `write_evaluation_report(test_file_path, report_path)`: calculates accuracy, precision, and recall.
- `make_large_test_file(output_path, total_records)`: creates 1200 realistic test records.
- `main()`: runs the complete flow.

## Evaluation

The script creates a realistic test file with 1200 records and 10,800 labelled PII values.

Current result:

```text
Accuracy: 0.72
Precision: 0.90
Recall: 0.79
```

Precision is high, so most redacted values are correct. Recall is lower because some names, company names, and addresses are missed when they appear in natural sentences.

## Limitation

Regex is simple and explainable, but it is not perfect for open-ended text. For better results, I would add an NLP named entity recognition model.
