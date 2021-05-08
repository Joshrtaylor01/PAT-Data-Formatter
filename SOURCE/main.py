import csv
import os
import pandas as pd
from numpy import percentile


def get_xlsx_files():
    # Loops through the current working directory
    # and returns all the appropriate files
    file_list = []
    for file in os.listdir():
        if file.startswith("pat") and file.endswith(".xlsx"):
            # Removes the .xlsx extension because I had already coded
            # to work without the extension
            file_list.append(file[:-5])
    return file_list


def create_year_groups(file):
    # Im certain there is a more effective way to do this but
    # I refuse to learn more about .xlsx reading and supporting
    # this strange proprietary format
    read_file = pd.read_excel(f"{file}.xlsx", skiprows=6)
    read_file.to_csv(f"{file}.csv", index=None)

    year_groups = dict()
    with open(f"{file}.csv", mode="r") as csv_file:
        csv_data = csv.reader(csv_file, delimiter=",")
        # Skips the first line and makes it the header
        header = next(csv_data)
        for row in csv_data:
            if row[9] not in year_groups:
                # If the current students year group is not in the dict add it
                year_groups.update({row[9]: []})
            year_groups[row[9]] += [row]

    # Adds the header to the dictionary so only one thing is being returned
    year_groups.update({"header": header})
    os.remove(f"{file}.csv")
    return year_groups


def make_seperate_csvs(year_groups, file):
    # Removes the header so that it doesn't get passed through the for loop
    header = year_groups.pop("header")
    for year_group in year_groups.keys():
        # Strips the random numbers from the end
        # of the file name to make it easier to read
        file = file[:22] + year_group
        with open(f"{file}.csv", mode="w", newline="") as csv_file:
            csv_writer = csv.writer(csv_file, delimiter=",")
            csv_writer.writerow(header)
            csv_writer.writerows(year_groups[year_group])


def five_number_summary(year_groups, file):
    for year_group in year_groups.keys():
        scores = []
        for student in year_groups[year_group]:
            # Score is in the 4th row, adds it to score
            scores.append(float(student[3]))
        quartiles = percentile(scores, [25, 50, 75])
        scores.sort()
        min, max = scores[0], scores[-1]
        five_num_sum = [min, quartiles[0], quartiles[1], quartiles[2], max]
        # Saves creating a new dict, update the old one with new data
        # Means I can reuse some code
        five_num_sum.insert(0, year_group)
        year_groups.update({year_group: five_num_sum})
    # This is argueably the most convoluted way to do this but the final csv
    # needs to be in order, so yeah, it works and thats all.
    year_groups = dict(sorted(year_groups.items(), key=lambda item: int(item[0][4:])))
    file = file[:22] + "five-num-sum"
    header = ["Year Group", "Min", "Q1", "Median", "Q2", "Max"]
    # print(year_groups)
    with open(f"{file}.csv", mode="w", newline="") as csv_file:
        csv_writer = csv.writer(csv_file, delimiter=",")
        csv_writer.writerow(header)
        for year_group in year_groups.keys():
            csv_writer.writerow(year_groups[year_group])


def main():
    file_list = get_xlsx_files()
    for file in file_list:
        year_groups = create_year_groups(file)
        make_seperate_csvs(year_groups, file)
        five_number_summary(year_groups, file)


if __name__ == "__main__":
    main()
