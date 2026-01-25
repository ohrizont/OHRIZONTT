# Customer Churn Risk Ranking – Credit Card Customers

This project focuses on predicting and ranking customer churn risk using behavioral data from credit card customers.

Rather than treating churn as a pure binary classification problem, the objective is to generate a **risk ranking** that allows prioritization of customers for retention campaigns.

## Dataset
The project uses the **Credit Card Customers (Bank Churners)** dataset from Kaggle:

https://www.kaggle.com/datasets/sakshigoyal7/credit-card-customers

The dataset is not included in this repository due to licensing restrictions.

## Methodology
1. Data preprocessing and feature engineering
2. Unsupervised segmentation (K-Means, HDBSCAN)
3. Feature selection based on behavioral signals
4. Supervised modeling:
   - Logistic Regression (baseline)
   - Decision Tree
   - Random Forest
   - Gradient Boosting (final model)
5. Model evaluation using AUC and lift
6. Customer scoring and ranking

## Results
- Best model: **Gradient Boosting**
- AUC: ~0.91
- Lift @ Top 10%: ~4.6×

## Repository structure
- `notebooks/`: exploratory analysis and modeling
- `src/`: reusable Python modules
- `figures/`: plots used in analysis

## Key takeaway
Behavioral variables and recent changes in activity dominate churn prediction, and ranking-based evaluation provides more actionable insights than accuracy-based metrics.

## License
MIT
