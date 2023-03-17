import sys
import os
# import posixpath
import pandas as pd
import numpy as np
from datetime import datetime
from io import StringIO


def prepare_filenames():
    # arguments check
    n_args = len(sys.argv)
    if n_args < 2 or n_args > 3:
        raise Exception("ERROR: Invalid argument number \nUsage: ./topsis_solver.sh input_filename [output_directory]")
    in_filename = sys.argv[1]

    if n_args == 3:
        out_dir = sys.argv[2]
    else:
        out_dir = os.path.dirname(in_filename)

    if not os.path.exists(in_filename):
        raise Exception("ERROR: Input file does not exist")
    if not os.path.exists(out_dir):
        raise Exception("ERROR: Output directory does not exist")

    # creating unique output name
    now = datetime.now()
    out_filename = "report_" + str(now).replace(" ", "_").replace("-", "_").replace(":", "_").split(".")[0] + ".xlsx"
    out = os.path.join(out_dir, out_filename)

    return in_filename, out


def prepare_input_buffer(filename):
    """
    Function deals with blank lines and encoding errors in input file
    :param filename: path to file to read from
    :return: StringIO buffer with cleaned up file content
    """
    io = StringIO()

    with open(filename, 'r', errors='replace') as r:
        for line in r:
            if line.strip():
                io.write(line)
    io.seek(0)
    return io


if __name__ == "__main__":
    try:
        input_filename, output_filename = prepare_filenames()
    except Exception as e:
        print(e)
        sys.exit(1)

    # preparing file buffer
    buffer = prepare_input_buffer(input_filename)

    # reading main dataframe
    df = pd.read_csv(buffer, sep=';', index_col=0, skiprows=3, header=None, encoding_errors='replace')

    # resetting rows and column indexes
    df.columns = range(df.shape[1])
    # print(df)

    # reading weights as pandas Series
    buffer.seek(0)
    weights = pd.read_csv(buffer, sep=';', nrows=1,
                          header=None, index_col=0, decimal=',', encoding_errors='replace').squeeze()
    weights.reset_index(inplace=True, drop=True)
    if weights.sum() != 1:
        print("ERROR: weights do not sum to 1")
        sys.exit(1)
    # print(weights)

    # reading signs as pandas Series
    buffer.seek(0)
    signs = pd.read_csv(buffer, sep=';', skiprows=1,
                        nrows=1, header=None, index_col=0, dtype='str', encoding_errors='replace').squeeze()
    signs.reset_index(inplace=True, drop=True)
    # print(signs)

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
    print("Report file generated at " + os.path.abspath(output_filename))
