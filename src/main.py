import pandas as pd
import numpy as np

# reading main dataframe
df = pd.read_csv("../resources/TOPSIS_01.csv", sep=';', index_col=0, skiprows=3, header=None)
# resetting rows and column indexes
# df.reset_index(inplace=True, drop=True)
df.columns = range(df.shape[1])

# reading weights as pandas Series
weights = pd.read_csv("../resources/TOPSIS_01.csv", sep=';', nrows=1,
                      header=None, index_col=0, decimal=',').squeeze()
weights.reset_index(inplace=True, drop=True)

# reading signs as pandas Series
signs = pd.read_csv("../resources/TOPSIS_01.csv", sep=';', skiprows=1,
                    nrows=1, header=None, index_col=0, dtype='str').squeeze()
signs.reset_index(inplace=True, drop=True)

# calculating squared root of sum of squares for each column
root_sum_squared = np.sqrt((df**2).sum(axis=0))

# normalisation and multiplying each column by weight
weighted = (df / root_sum_squared) * weights

# calculating A* (max if positive column else min) and A' (min if positive column else max) for each column
a_star = pd.Series([col.max() if signs[i] == '+' else col.min() for (i, col) in weighted.items()])
a_prime = pd.Series([col.min() if signs[i] == '+' else col.max() for (i, col) in weighted.items()])

# calculating S*
# subtracting A* from each column and squared
s_star_squared = (weighted - a_star) ** 2
# calculating sum for each row and square root
s_star = np.sqrt(s_star_squared.sum(axis=1))

# calculating S'
# subtracting A' from each column and squared
s_prime_squared = (weighted - a_prime) ** 2
# calculating sum for each row and square root
s_prime = np.sqrt(s_prime_squared.sum(axis=1))

# calculating CI
ci = s_prime/(s_star + s_prime)
ci = pd.DataFrame(ci)
ci.columns = ["CI"]

# creating candidate ranking - sorting descending by CI score
ranked = ci.sort_values(["CI"], ascending=False)

# saving output to Excel file
ranked.to_excel("../resources/ranking.xlsx")
