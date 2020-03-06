import pandas as pd
import numpy as np
import matplotlib.pyplot as plt


df = pd.read_csv('california_housing_train.csv')
#df.plot.area()
#print (df.plot())

df["households"].bar()
plt.show()