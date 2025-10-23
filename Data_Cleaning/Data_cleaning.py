import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

sheets = pd.read_excel('/content/Regional Sales Dataset.xlsx', sheet_name= None)

sheets

df_sales = sheets['Sales Orders']
df_customers = sheets['Customers']
df_products = sheets['Products']
df_regions =  sheets['Regions']
df_budgets = sheets['2017 Budgets']

df_sales.head(5)

df_sales.isnull().sum()

df = df_sales.merge(
    df_products,
    how = 'left',
    left_on = 'Product Description Index',
    right_on= 'Index'
)

df_sales.head(5)

df = df_sales.merge(
df_customers,
how='left',
left_on='Customer Name Index',
right_on='Customer Index'
)

df = df.merge(
    df_products,
    how='left',
    left_on='Product Description Index',
    right_on='Index'
)

df = df.merge(
    df_regions,
    how='left',
    left_on='Delivery Region Index',
    right_on='id'
)

df_state_reg = sheets['State Regions']

df = df.merge(
    df_state_reg["State Code", "Region"],
    how='left',
    left_on='state_code',
    right_on='State Code'
)

df.head(1)

df = df.merge(
    df_state_reg[["State Code", "Region"]],
    how='left',
    left_on='state_code',
    right_on='State Code'
)

df_state_reg.head(2)

df_state_reg.columns = df_state_reg.iloc[0]
df_state_reg = df_state_reg.drop(0)

df = df.merge(
    df_state_reg[["State Code", "Region"]],
    how='left',
    left_on='state_code',
    right_on='State Code'
)

print(df_state_reg.columns.tolist())

df = df.merge(
    df_state_reg[["State Code", "Region"]],
    how='left',
    left_on='state_code',
    right_on='State Code'
)

[print(df.columns.tolist)]

df = df.rename(columns={'state_code_x': 'state_code'})

df = df.merge(
    df_state_reg[["State Code", "Region"]],
    how='left',
    left_on='state_code',
    right_on='State Code'
)

df.to_csv('file.csv')

# Step 1: Drop duplicate _y columns
df = df.drop(columns=[col for col in df.columns if col.endswith('_y')])

# Step 2: Rename _x columns back to their original names
df = df.rename(columns=lambda x: x[:-2] if x.endswith('_x') else x)

# Step 3: Optional - verify cleanup
print(df.columns.tolist())

cols_to_drop = ["Customer Index",'Index', 'id', 'State Code']
df = df.drop(columns=cols_to_drop, errors='ignore')

df.columns = df.columns.str.lower()

df.columns.values

sheets = pd.read_excel('/content/Regional Sales Dataset.xlsx', sheet_name= None)

df = sheets['Sales Orders']
df_customers = sheets['Customers']
df_products = sheets['Products']
df_regions = sheets['Regions']
df_budgets = sheets['2017 Budgets']
df = df.merge(
    df_customers,
    how='left',
    left_on='Customer Name Index',
    right_on='Customer Index'
)
df = df.merge(
    df_products,
    how='left',
    left_on='Product Description Index',
    right_on='Index'
)
df = df.merge(
    df_regions,
    how='left',
    left_on='Delivery Region Index',
    right_on='id'
)
df_state_reg = sheets['State Regions']
df_state_reg.columns = df_state_reg.iloc[0]
df_state_reg = df_state_reg.drop(0)
df = df.rename(columns={'state_code_x': 'state_code'})
df = df.merge(
    df_state_reg[["State Code", "Region"]],
    how='left',
    left_on='state_code',
    right_on='State Code'
)
df = df.drop(columns=[col for col in df.columns if col.endswith('_y')])
df = df.rename(columns=lambda x: x[:-2] if x.endswith('_x') else x)
cols_to_drop = ["Customer Index",'Index', 'id', 'State Code']
df = df.drop(columns=cols_to_drop, errors='ignore')
df.columns = df.columns.str.lower()

df.columns.values

import pandas as pd

sheets = pd.read_excel('/content/Regional Sales Dataset.xlsx', sheet_name= None)

df = sheets['Sales Orders']
df_customers = sheets['Customers']
df_products = sheets['Products']
df_regions = sheets['Regions']
df_budgets = sheets['2017 Budgets']
df = df.merge(
    df_customers,
    how='left',
    left_on='Customer Name Index',
    right_on='Customer Index'
)
df = df.merge(
    df_products,
    how='left',
    left_on='Product Description Index',
    right_on='Index'
)
df = df.merge(
    df_regions,
    how='left',
    left_on='Delivery Region Index',
    right_on='id'
)
df_state_reg = sheets['State Regions']
df_state_reg.columns = df_state_reg.iloc[0]
df_state_reg = df_state_reg.drop(0)
df = df.rename(columns={'state_code_x': 'state_code'})
df = df.merge(
    df_state_reg[["State Code", "Region"]],
    how='left',
    left_on='state_code',
    right_on='State Code'
)
df = df.drop(columns=[col for col in df.columns if col.endswith('_y')])
df = df.rename(columns=lambda x: x[:-2] if x.endswith('_x') else x)
cols_to_drop = ["Customer Index",'Index', 'id', 'State Code']
df = df.drop(columns=cols_to_drop, errors='ignore')
df.columns = df.columns.str.lower()

df.columns.values

cols_to_keep = [
'ordernumber',
'orderdate',
'customer names',
'channel',
'product name',
'order quantity',
'unit price',
'line total',
'total unit cost',
'state_code',
'county',
'state',
'region',
'latitude',
'longitude',
'2017 budgets'
]

df.columns.values



df= df.merge(
    df_budgets,
    how='left',
    on='product name'
)

df_budgets.columns = df_budgets.columns.str.lower()

df = df[cols_to_keep]

df.to_csv('data.csv')