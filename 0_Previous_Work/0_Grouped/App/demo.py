import numpy as np
import pandas as pd

dataset = pd.read_excel('dataset.xlsx', sheet_name='AUO')

row, column = dataset.shape
print("Dataset size : ", row , "x" , column)
print("Null Row : ", row + 1)
print("Null column : ", column + 1)

arr = dataset['Vgate'].to_numpy()
# input 
print("search value : ")
value = float(input())  
above = arr[np.searchsorted(arr,value,'left')-1]
below = arr[np.searchsorted(arr,value,'right')]
X = value
x1 = min(above, below)
x2 =  max(above, below)
dd = dataset.query("Vgate == @above or Vgate == @below")
dd = dd.sort_index(ascending=True)

ABSID1_arr = dd.query("Vgate == @x1")['ABSID'].to_numpy()
y1 = ABSID1_arr[0]
ABSID2_arr = dd.query("Vgate == @x2")['ABSID'].to_numpy()
y2 = ABSID2_arr[0]

print("(x1, x2) : ",(x1, x2))
print("(y1, y2) : ",(y1, y2))
print("X : ",X)
Y = y1 + (X-x1)*(y2-y1)/(x2-x1)
print("Y", Y)