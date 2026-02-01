# Test Set Results Summary

This document summarizes the final out-of-sample results of the fraud detection system.
All metrics reported here correspond to the held-out test set.
No thresholds or parameters were adjusted after training.

## Global Performance

- Fraud recall (transactions): **85.6%**
- Fraud recall (monetary value): **91.5%**
- Precision: **8.6%**
- Transactions flagged: **3.83%**
- Monetary volume flagged: **4.93%**

## Performance by Transaction Amount

### Transactions Below Mean Amount

- Fraud recall (transactions): **51.8%**
- Fraud recall (monetary value): **54.0%**
- Precision: **2.6%**
- Transactions flagged: **3.92%**
- Monetary volume flagged: **4.73%**

### Transactions Above Mean Amount

- Fraud recall (transactions): **94.6%**
- Fraud recall (monetary value): **95.6%**
- Precision: **12.9%**
- Transactions flagged: **3.77%**
- Monetary volume flagged: **5.00%**

## Interpretation

The system prioritizes the detection of high-impact fraudulent transactions.
With less than 4% of transactions flagged, over 90% of fraudulent monetary value is captured.

Lower-value fraud remains harder to detect and is intentionally deprioritized
to avoid excessive customer friction.

## Notes

- Model: LightGBM
- Threshold policy: amount-aware, frozen on training data
- Evaluation: single held-out test set
