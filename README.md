# Customer Segmentation: RFM & Seasonality Analysis 📊

This project focuses on identifying customer behavior patterns for a retail dataset using **Unsupervised Machine Learning**.

## 🎯 Business Objective
The goal is to segment customers to allow for targeted marketing strategies, identifying who are the "Champions", who are at risk of "Churn", and how seasonality (like Valentine's Day or Mother's Day) affects purchasing behavior.

## 🛠️ Data Pipeline & Technical Highlights
- **Data Cleaning:** Implemented a robust cleaning script using `RapidFuzz` for fuzzy matching of city names (correcting typos like "Londre" to "London").
- **Feature Engineering:** Created RFM metrics (Recency, Frequency, Monetary) and a monthly seasonality matrix.
- **Dimensionality Reduction:** Used **PCA (Principal Component Analysis)** to visualize high-dimensional clusters in a 2D space.
- **Clustering:** Applied the **K-Means** algorithm, using the Elbow Method and Silhouette Score to find the optimal number of clusters (k=5).

## 📈 Key Insights
- **Seasonal Peaks:** Significant sales growth identified in February (Valentine's Day) and March (Mother's Day UK).
- **Customer Profiles:** - *Cluster VIP:* High frequency and high spend.
  - *Cluster Occasional:* Low engagement, mostly active during specific holidays.

## 🧰 Tech Stack
- **Python:** Pandas, NumPy, Scikit-learn, Matplotlib, Seaborn.
- **Algorithms:** K-Means, PCA, Linear Regression (for quantity vs. total analysis).
