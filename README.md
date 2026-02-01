# Cost-Sensitive Fraud Detection with Dynamic Thresholding

This project presents a production-oriented fraud detection system focused on maximizing
fraud recall by monetary value while keeping customer friction low.

Instead of optimizing a single global threshold, a simple and interpretable decision policy
based on transaction amount is applied on top of a LightGBM model.

## Key Results (Test Set)

- Fraud recall (monetary value): **91.5%**
- Fraud recall (transactions): **85.6%**
- Transactions flagged: **3.8%**
- Monetary volume flagged: **4.9%**

## Approach

- Exploratory clustering (K-means) to understand transaction patterns
- Logistic regression used as a diagnostic tool for feature selection
- LightGBM model for scoring
- Amount-aware decision thresholds frozen on training data

## Repository Structure

- `notebooks/`: narrative case study (Kaggle-style)
- `src/`: training and scoring scripts
- `results/`: summarized evaluation results

## Kaggle Notebook

The full case study is available on Kaggle:
(https://www.kaggle.com/code/ohrizonte/cost-sensitive-fraud-detection-with-dynamic-thresh)

## Disclaimer

This repository does not include raw datasets or trained models.
The project focuses on methodology, decision logic, and evaluation.
