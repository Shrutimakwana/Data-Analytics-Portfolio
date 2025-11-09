import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

#load the excel file
df=pd.read_excel('food_orders.xlsx')

#show first five rows
print("Preview of Data : \n")
print(df.head())

#print total revenue
total_revenue=df['Total'].sum()
print(f"\n Total Revenue is: {total_revenue}")

#Top 5 selling items
top_items=df.groupby('Item')['Quantity'].sum().sort_values(ascending=False).head()
print("\n Top 5 Selling Items:")
print(top_items)

#Revenue by city
city_revenue=df.groupby('City')['Total'].sum().sort_values(ascending=False)
print("Revenue By City:")
print(city_revenue)

#Convert data for time based analysis
df['OrderDate']=pd.to_datetime(df['OrderDate'])

#Monthly Sales Trend
monthly_sales=df.groupby(df['OrderDate'].dt.to_period('M'))['Total'].sum()
#print(monthly_sales)

#Plot the monthly Sales
plt.figure(figsize=(8,5))
monthly_sales.plot(kind='line',marker='o',color='teal',linewidth=2)
plt.title("Monthly Sales Trend",fontsize=14,fontweight='bold')
plt.xlabel('Month')
plt.ylabel('Revenue')
plt.grid(True,linestyle='--',alpha=0.6)
plt.tight_layout()
plt.show()
#City wise revenue bar chart
plt.figure(figsize=(8,5))
city_revenue.plot(kind='bar',color='coral',edgecolor='black')
plt.title("City-wise Revenue Bar Chart")
plt.xlabel("City")
plt.ylabel("Revenue")
plt.xticks(rotation=50,ha="right")
plt.tight_layout()
plt.show()

#Item-wise sales bar chart
plt.figure(figsize=(8,5))
top_items.plot(kind='bar',color='skyblue',edgecolor='black')
plt.title("Top 5 Selling Items",fontsize=14,fontweight='bold')
plt.xlabel('Item')
plt.ylabel('Quantity SOld')
plt.xticks(rotation=30,ha='right')
plt.tight_layout()
plt.show()

#Revenue By Category
category_revenue=df.groupby('Category')['Total'].sum().sort_values(ascending=False)
plt.figure(figsize=(8,5))
sns.barplot(x=category_revenue.values,y=category_revenue.index,palette='mako')
plt.title("Revenue By Category",fontsize=14,fontweight='bold')
plt.xlabel("total Revenue")
plt.ylabel("Category")
plt.tight_layout()
plt.show()

#Export Summary to Excel
Summary={
    'Total Revenue(Rs.)':[total_revenue],
    'Top Item':[top_items.index[0]],
    'Top Item Quantity':[top_items.iloc[0]],
    'Top City':[city_revenue.index[0]],
    'Top City Revenue':[city_revenue.iloc[0]],
    'Top Ctegory Revenue':[category_revenue.index[0]]
}
summary_df=pd.DataFrame(Summary)
summary_df.to_excel('food_summary.xlsx',index=False)
print("\n Summary Exported to 'food_summary.xlsx' ")