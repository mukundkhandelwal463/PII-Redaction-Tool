# Evaluation Report

I tested the script using `synthetic_eval_labeled_realistic_1200.txt`. This file has labelled sample tickets, so the script can compare its answers with the correct answers.

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

Result: The script performed well on emails, phone numbers, SSNs, credit cards, and IP addresses. It missed more names, companies, and addresses because those were written in different real-world styles. Some false positives happened because a few sample/reference numbers looked like valid card numbers.

Note: This is synthetic test data. For a real business use case, real documents should also be manually checked because names, companies, and addresses can be written in many different ways.
