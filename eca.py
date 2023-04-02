import os
import pandas as pd
import re

folder_path = "D:\\ECA"

def extract_season_and_episode(short_desc):
    season = re.search(r"Seizoen (\d+)", short_desc)
    if season:
        season_number = int(season.group(1))
    else:
        season_number = None

    episode = re.search(r"Aflevering (\d+)|afl\. (\d+)|afl (\d+)", short_desc)
    if episode:
        episode_number = int([group for group in episode.groups() if group is not None][0])
    else:
        episode_number = None

    return season_number, episode_number

all_data = []

for file in os.listdir(folder_path):
    if file.endswith(".xlsx"):
        file_path = os.path.join(folder_path, file)
        df = pd.read_excel(file_path, skiprows=2)
        df["Season number"], df["Episode number"] = zip(*df["Short description"].apply(extract_season_and_episode))
        all_data.append(df)

        # Save the output to the same Excel file
        df.to_excel(file_path, index=False)

combined_df = pd.concat(all_data, axis=0, ignore_index=True)
print(combined_df)
