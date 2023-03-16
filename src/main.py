import sys
import os
import posixpath
import pandas as pd
import numpy as np
from datetime import datetime


def prepare_filenames():
    # arguments check
    n_args = len(sys.argv)
    if n_args < 2 or n_args > 3:
        raise Exception("ERROR: Invalid argument number \nUsage: ./topsis_solver.sh input_filename [output_directory]")
    in_filename = sys.argv[1]

    if n_args == 3:
        out_dir = sys.argv[2]
    else:
        out_dir = "."

    if not os.path.exists(in_filename):
        raise Exception("ERROR: Input file does not exist")
    if not os.path.exists(out_dir):
        raise Exception("ERROR: Output directory does not exist")

    # creating unique output name
    now = datetime.now()
    out_filename = "ranking_" + str(now).replace(" ", "_").replace("-", "_").replace(":", "_").split(".")[0] + ".xlsx"
    out = posixpath.join(out_dir, out_filename)

    return in_filename, out


if __name__ == "__main__":
    try:
        input_filename, output_filename = prepare_filenames()
    except Exception as e:
        print(e)
        sys.exit(1)

    # reading main dataframe
    df = pd.read_csv(input_filename, sep=';', index_col=0, skiprows=3, header=None)

    # resetting rows and column indexes
    # df.reset_index(inplace=True, drop=True)
    df.columns = range(df.shape[1])

    # reading weights as pandas Series
    weights = pd.read_csv(input_filename, sep=';', nrows=1,
                          header=None, index_col=0, decimal=',').squeeze()
    weights.reset_index(inplace=True, drop=True)
    if weights.sum() != 1:
        print("ERROR: weights do not sum to 1")
        sys.exit(1)

    # reading signs as pandas Series
    signs = pd.read_csv(input_filename, sep=';', skiprows=1,
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
    ranked.to_excel(output_filename)
    print("Report file generated at " + output_filename)
