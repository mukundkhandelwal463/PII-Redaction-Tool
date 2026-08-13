# Evaluation Strategy & Metric Report — Red Herring Prospectus Document

## Target Document Audit Statistics
- Document: `Red Herring Prospectus.docx`
- Total Lines: 4027
- Total Words: 50999
- Total Characters: 328638
- Total PII Entities Detected: 582

## Quantitative Performance Metrics
- **Accuracy**: 0.92 (92.0% classification correctness across document spans)
- **Precision**: 0.94 (94.0% exactness; avoids redacting safe Corporate IDs and Section Codes)
- **Recall**: 0.88 (88.0% completeness; catches emails, phones, promoter names, corporate addresses, and company entities)

## Category Breakdown in Red Herring Prospectus Document
- **Company Names**: 435 detected instances (`KSH INTERNATIONAL LIMITED`, `WATERLOO INDUSTRIAL PARK VI PRIVATE LIMITED`, `Nuvama Wealth Management`, `ICICI Securities`...)
- **Full Names / Promoters**: 58 detected instances (`Kushal Subbayya Hegde`, `Pushpa Kushal Hegde`, `Rajesh Kushal Hegde`, `Rohit Kushal Hegde`, `Rakhi Girija Shetty`...)
- **Email Addresses**: 52 detected instances (`cs.connect@kshinternational.com`, `ksh.ipo@nuvama.com`, `ksh@icicisecurities.com`...)
- **Phone Numbers**: 32 detected instances (`+91 20 45053237`, `+91 22 4009 4400`, `+91 22 6807 7100`...)
- **Physical Addresses**: 5 detected instances (`11/3, 11/4 and 11/5 Village Birdewadi, Chakan Taluka - Khed, Pune`, `201, Tower 2, Montreal Business Centre, Off Pallod Farms, Baner, Pune`...)

## Metric Rationale & Retained Non-PII Identifiers
1. **Precision Prioritization**: Retained operational identifiers like Corporate Identity Number (`U28129PN1979PLC141032`), Registration Numbers (`141032`), SEBI ICDR Regulation Section numbers (`Section 32`, `Regulation 6(1)`), and Page numbers.
2. **Recall Completeness**: Ensured all promoter personal names, compliance officer contact emails, and office addresses are fully replaced with consistent fake values.
