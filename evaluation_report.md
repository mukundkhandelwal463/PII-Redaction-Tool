# Evaluation Report

Test file: `synthetic_eval_labeled_realistic_1200.txt`

Total labelled PII values: 10800
Total detected PII values: 9440

- Accuracy: 0.72
- Precision: 0.90
- Recall: 0.79
- Correct redactions: 8480
- Wrong redactions: 960
- Missed PII values: 2320

Score by PII type:

- address: precision 1.00, recall 0.40, correct 480
- company: precision 0.40, recall 0.27, correct 320
- credit_card: precision 0.83, recall 1.00, correct 1200
- dob: precision 1.00, recall 0.80, correct 960
- email: precision 1.00, recall 1.00, correct 1200
- full_name: precision 0.75, recall 0.60, correct 720
- ip_address: precision 1.00, recall 1.00, correct 1200
- phone: precision 1.00, recall 1.00, correct 1200
- ssn: precision 1.00, recall 1.00, correct 1200

The script works better on fixed patterns like emails, phones, SSNs, cards, and IPs. It misses more names, companies, and addresses because they can be written in many styles.
