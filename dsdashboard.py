import pandas as pd
import streamlit as st 
import openpyxl
import os
import plotly.express as px
import matplotlib.pyplot as plt
import seaborn as sns

st.markdown("<h1 style='text-align: center; font-family: Arial; font-size: 36px;'>Minger Sales Analysis</h1>", unsafe_allow_html=True)#this was added to put the title to center 
df= pd.read_excel('cleaned_data.xlsx')

df = pd.read_excel('cleaned_data.xlsx', engine='openpyxl')

#Adding the Minger logo
col1, col2, col3= st.columns(3)
with col1:
    st.image("Minger logo.png", width=300)

#Visualization 01 & 02- Sum of sales and profit
selected_year = st.selectbox("Select Year", ['All'] + list(df["Order Date"].dt.year.unique()))

#setting up filter so you could see sales and profit for each year
if selected_year != 'All':
    filtered_df = df[df["Order Date"].dt.year == selected_year]
else:
    filtered_df = df

total_sales = filtered_df['Sales'].sum()
total_profit = filtered_df['Profit'].sum()
col1, col2 = st.columns(2)

with col1:
    st.subheader('Total Sales')
    st.metric(label="Sales", value=f"${total_sales:,.0f}", delta=None)

with col2:
    st.subheader('Total Profit')
    st.metric(label="Profit", value=f"${total_profit:,.0f}", delta=None)



#Visualization 03 & 04 - Top 5 products with sales and profitability
Topproducts = df.groupby('Product Name')['Sales'].sum().nlargest(5).round()
col1.subheader('Top 5 Best Selling Products')
col1.bar_chart(Topproducts.sort_values() , color= "#CCCCFF", use_container_width=True)

Profitable_products = df.groupby("Product Name")["Profit"].sum().nlargest(5).round()
col2.subheader('Top 5 Profitable Products')
col2.bar_chart(Profitable_products.sort_values(), color= "#6495ED", use_container_width=True)

# visualization 04 pie chart to show Category wise Proft
with col1.subheader("Category wise Profit"):
    fig1 = px.pie(df, values="Profit", names="Category", template="plotly")
    fig1.update_traces(textposition="inside", textinfo="percent+label")
    col1.plotly_chart(fig1, use_container_width=True)

#visualization 05 pie chart to show Category Wise Sales
with col2.subheader("Segment wise sales"):
    fig2=px.pie(df, values="Sales",names="Segment",template="seaborn")
    fig2.update_traces(textposition="inside",textinfo="percent+label")
    col2.plotly_chart(fig2, use_container_width=True)

#visualization 06 Time Series Analysis to show yearly and montly sales 
st.subheader('Time Series Analysis')

# created a Filter to check by  year
selected_years = st.multiselect("Select Year", df["Order Date"].dt.year.unique())

# Filter the dataframe based on selected years
filtered_df = df[df["Order Date"].dt.year.isin(selected_years)]

linechart = pd.DataFrame(filtered_df.groupby(filtered_df["Order Date"].dt.strftime("%Y : %b"))["Sales"].sum()).reset_index()
fig2 = px.line(linechart, x="Order Date", y="Sales", labels={"Sales": "Sales"}, height=500, width=700,
               template="plotly")
st.plotly_chart(fig2, use_container_width=True)

# visualization 07 chelorepath map for sales
countrysales=df.groupby('Country')[['Sales','Profit']].sum().reset_index()
fig_sales=px.choropleth(countrysales,locations='Country',locationmode='country names',color='Sales',hover_name='Country',color_continuous_scale='rainbow',title='Sales per Country',template='ggplot2',width=1000)
fig_sales.update_layout(title_font_size=28)

st.plotly_chart(fig_sales)

#visualization 08 map for profit 
countryprofts=df.groupby('Country')[['Profit','Sales']].sum().reset_index()
figprofit=px.choropleth(countryprofts,locations='Country',locationmode='country names',color='Profit',hover_name='Country',color_continuous_scale='rainbow',title='Profit per Country',template='ggplot2',width=1000)
fig_sales.update_layout(title_font_size=28)
st.plotly_chart(figprofit)

#Visualization 09 number of Orders per ship mode
st.subheader("Number of Orders by Ship Mode")
orders_by_ship_mode = df['Ship Mode'].value_counts()

#Total sales by Ship Mode
sales_by_ship_mode = df.groupby('Ship Mode')['Sales'].sum()

# Bar chart
plt.style.use('dark_background')
plt.figure(figsize=(10, 5))
colors = ['#21618C', '#58D68D', '#FFC300', '#FF5733']  
plt.bar(orders_by_ship_mode.index, orders_by_ship_mode.values, color=colors,)
plt.xlabel('Ship Mode')
plt.ylabel('Number of Orders')
plt.xticks(rotation=45) 

st.pyplot(plt) 

plt.style.use('dark_background')

#visualization 10 scatter plot to show Sales and discount for each categry 
st.subheader("Sales and Discount rate according to category")

#Slicer setup
productcategories = df['Category'].unique()
selectedcategory = st.selectbox('Select Category', productcategories)

# Filtering data based on selected category
filtered_data = df[df['Category'] == selectedcategory]

#Scatter Plot
fig = px.scatter(filtered_data, x='Discount', y='Sales',title=f'Sales vs. Discount for {selectedcategory}',labels={'Discount': 'Discount (%)', 'Sales': 'Sales ($)'},hover_data=['Product Name'],color='Discount', template='plotly') 

st.plotly_chart(fig)

# visualization 11 Month wise Sub-Category Sales Summary
st.subheader("Monthly wise Sub-Category Sales Summary")

# Display a table of month wise sub-category sales
df["month"] = df["Order Date"].dt.month_name()
sub_category_Year = pd.pivot_table(data=df, values="Sales", index=["Sub-Category"], columns="month")
st.write(sub_category_Year.style.background_gradient(cmap="Greens"))

# visualization 12 Histogram of sales for Sub-Category
st.subheader("Histogram of Sales for Sub-Category")
fig = px.histogram(df, x="Sub-Category", title="Histogram of Sales for Sub-Category")
st.plotly_chart(fig, use_container_width=True)

































































