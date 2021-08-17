# Data Preprocessing Template

# Importing the libraries
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
from datetime import date, datetime
from getData import GetData
SIZE = (36,18)
FILE = 'Data.xlsx'
# Importing the dataset
#dataset = pd.read_csv('Data.csv')
Time = GetData(FILE)
Time = Time[:3]+str(int(Time[-2:])+3)
raw =  pd.read_excel(FILE)
raw = pd.read_excel(FILE, converters= {'A1': pd.to_datetime})
req = raw.iloc[-1, 1:-1].values
temp = []
for i in req:
    temp.append(int(i))
req = [temp]
dataset = raw.iloc[:-1, :]
X = dataset.iloc[:, 1:-1].values
y = dataset.iloc[:, -1].values
req2 = int(y[-1])
time = dataset.iloc[:, 0].values

# Splitting the dataset into the Training set and Test set
from sklearn.model_selection import train_test_split
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size = 0.2, random_state = 0)
from sklearn.linear_model import LinearRegression
regressor = LinearRegression()
regressor.fit(X_train, y_train)

# Predicting the Test set results
y_pred = regressor.predict(X_test)
np.set_printoptions(precision=2)
x_axis = np.arange(len(y_pred))
accuracy = []
pred = []
actual = []
plt.plot(x_axis,y_pred,color = 'blue',marker = '')
plt.plot(x_axis,y_test,color = 'red',marker = '' )

for x, y in zip(x_axis, y_pred):
    plt.text(x, y, str(x), color="blue", fontsize=12)
plt.margins(0.1)
for x, y in zip(x_axis, y_test):
    plt.text(x, y, str(x), color="red", fontsize=12)
plt.margins(0.1)





for i in range(1,len(y_test)):
    if y_test[i]-y_test[i-1] > 0:
        actual.append('UP')
    else:
        actual.append('Down')
    
    if y_pred[i]-y_pred[i-1] > 0:
        pred.append('UP')
    else:
        pred.append('Down')
        

for i in range(0,len(pred)):
    #print(i+1,"__",actual[i],'___',pred[i])
    if actual[i] == pred[i]:
        accuracy.append(0)
    else:
        accuracy.append(1)

print(sum(accuracy)*100/len(accuracy))
temp = int(regressor.predict(req)[0])
print('Predicted Value '+str(temp)+' from '+str(req2)+' at Time:'+Time  +'\n')
f = open("result.txt", "a")
f.write(str(temp))

if temp-req2 > 0:
    print('UP')
    f.write(':UP from'+str(req2)+'at Time:'+Time  +'\n')
else:
    print('Down')
    f.write(':Down from'+str(req2)+'at Time:'+Time+  '\n')
f.close()
print('Press any key to quit')
input()

plt.show()


