import pandas as pd
from rapidfuzz import fuzz
from tqdm import tqdm
from concurrent.futures import ThreadPoolExecutor

# Threshold for match (can be adjusted)
SIMILARITY_THRESHOLD = 90

def compare_row(row1, df2):
    best_score = -1
    best_idx = -1
    for idx2, row2 in df2.iterrows():
        score = fuzz.token_sort_ratio(str(row1), str(row2))
        if score > best_score:
            best_score = score
            best_idx = idx2
    return best_score, best_idx

def process_row(args):
    idx1, row1, df2 = args
    score, best_idx = compare_row(row1, df2)
    if score >= SIMILARITY_THRESHOLD:
        return {
            "matched": True,
            "line_file1": idx1 + 2,
            "line_file2": best_idx + 2,
            "row_file1": row1.tolist(),
            "row_file2": df2.iloc[best_idx].tolist()
        }
    else:
        return {
            "matched": False,
            "line_file1": idx1 + 2,
            "row_file1": row1.tolist()
        }

def main():
    df1 = pd.read_csv("firm54.csv")
    df2 = pd.read_csv("firm97.csv")

    difference_rows = []
    unmatched_rows = []

    with ThreadPoolExecutor() as executor:
        results = list(tqdm(executor.map(process_row, [(i, row, df2) for i, row in df1.iterrows()]), total=len(df1)))

    for result in results:
        if result["matched"]:
            difference_rows.append([
                result["line_file1"], result["line_file2"]
            ] + result["row_file1"] + result["row_file2"])
        else:
            unmatched_rows.append([
                result["line_file1"]
            ] + result["row_file1"])

    # Generate headers
    diff_headers = ["Line_File1", "Line_File2"] + \
                   [f"File1_{col}" for col in df1.columns] + \
                   [f"File2_{col}" for col in df2.columns]
    unmatched_headers = ["Line_File1"] + [f"File1_{col}" for col in df1.columns]

    # Save results
    pd.DataFrame(difference_rows, columns=diff_headers).to_csv("difference.csv", index=False)
    pd.DataFrame(unmatched_rows, columns=unmatched_headers).to_csv("unmatched.csv", index=False)

    print("Matching complete. 'difference.csv' and 'unmatched.csv' generated.")

if __name__ == "__main__":
    main()
  
