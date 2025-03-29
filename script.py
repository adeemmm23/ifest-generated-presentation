import pandas as pd
import os

FOLDER: str = 'files'

os.makedirs(FOLDER, exist_ok=True)

for filename in os.listdir('files'):
    os.remove(f'{FOLDER}/{filename}')

awards_index: dict[str, int] = {
    "honorable": 7,
    "bronze": 9,
    "silver": 11,
    "gold": 13,
    "top10": 15
}


def csv_to_dict(filename: str) -> dict[str, list[list[str]]]:
    # Read CSV, preserving blank lines
    df = pd.read_csv(filename, skip_blank_lines=False)
    awards = {}
    current_award = None
    current_list = []

    for _, row in df.iterrows():
        if pd.notna(row.iloc[0]):  # If first column is not NaN, it's a new award
            if current_award:  # Save previous group
                awards.setdefault(current_award, []).append(current_list)
            current_award = row.iloc[0].strip()
            current_list = [row.iloc[1].strip()] if pd.notna(
                row.iloc[1]) else []
        else:  # Continuation of previous award
            if pd.notna(row.iloc[1]):
                current_list.append(row.iloc[1].strip())

    if current_award:  # Append last group
        awards.setdefault(current_award, []).append(current_list)

    return awards


def main():
    df = pd.read_csv('data.csv')

    projects = csv_to_dict('projects.csv')

    for award, namesList in projects.items():
        with open(f'{FOLDER}/{award.lower()}.txt', 'w',  encoding="utf-8") as f:
            for names in namesList:
                f.write(f'generateSlide {awards_index[award.lower()]}, "')
                for index, name in enumerate(names):
                    f.write(f'{name}')
                    if index != len(names) - 1:
                        f.write('" + vbNewLine + "')
                f.write('"')
                f.write('\n')


if __name__ == '__main__':
    main()
