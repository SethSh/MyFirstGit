from aocd import get_data
from aocd.models import Puzzle
import pandas as pd

puzzle = Puzzle(year=2021, day=1)
input_data = puzzle.input_data
string_list = input_data.split("\n")

integer_map = map(int, string_list)
integer_list = list(integer_map)

df = pd.DataFrame(integer_list, columns = ['depth'])
print(df)
diff =df.depth.diff().dropna() 
print(diff)
print((diff > 0).value_counts())