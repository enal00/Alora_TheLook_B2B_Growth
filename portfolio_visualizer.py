
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load data
df = pd.read_csv('rfm_segmentation_CLEAN.csv')

# Membersihkan data: Anggap nilai negatif sebagai data error/future date (set ke 0)
df['recency'] = df['recency'].apply(lambda x: 0 if x < 0 else x)

# Visualisasi: RFM Segmentation Treemap/Bar
plt.figure(figsize=(10, 6))
sns.barplot(data=df.groupby('segment')['monetary'].sum().reset_index(), x='segment', y='monetary')
plt.title('Total Revenue Contribution per Segment')
plt.ylabel('Total Monetary ($)')
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig('rfm_revenue_plot.png')
print("Visualisasi tersimpan sebagai rfm_revenue_plot.png")
