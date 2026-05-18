import pandas as pd
from pathlib import Path

#we put the data of each laliga table from 2020-21 to 2024-25 season in a single excel sheet

inputPath = Path("/Users/saeedfakhouri/Desktop/LaLiga_CIB_Project/data/LaLiga_Stats_CIB.xlsx")

#here we take into account the data, not the actual header in the excel sheet, we need to cleanup the sheet
dataExcel = pd.read_excel(inputPath, sheet_name="LaLigaTables", header=None)

cleanRows = []
currentSeason = None

#what this loop does is analyze the entirety of the excel sheet containing the data of each season,
#it divides each season and stores it inside the cleanRows list, and we then create a data frame from it
for index, row in dataExcel.iterrows():
    firstCell = row[0]
    #here we do row by row, and when there is a season row we store the season name and skip that header row
    if isinstance(firstCell, str) and firstCell.startswith("Season"):
        currentSeason = firstCell.replace("Season ", "")
        continue

    if pd.notna(row[1]) and pd.notna(row[2]):
        cleanRows.append({
            "Season": currentSeason,
            "Position" : row[1],
            "Team" : row[2],
            "Played" : row[3],
            "Wins" : row[4],
            "Draws" : row[5],
            "Losses" : row[6],
            "GoalsFor" : row[7],
            "GoalsAgainst" : row[8],
            "GoalDifference" : row[9],
            "Points": row[10]
        })

fullData = pd.DataFrame(cleanRows)

fullData["Team"] = (fullData["Team"].astype(str).str.strip())

fullData["GoalDifference"] = (fullData["GoalDifference"].astype(str).str.replace("−","-",regex=False))


numberColumns = [
    "Position", "Played", "Wins", "Draws", "Losses", "GoalsFor", "GoalsAgainst", "GoalDifference", "Points"
]

#Here we clean up the cells inside of excel, so that pandas can have access to numbers (so strings turn to NaN, numbers turn to actual numerical values)
for col in numberColumns:
    fullData[col] = pd.to_numeric(fullData[col], errors="coerce")

seasonOrder = {
    "2020-21": 1,
    "2021-22": 2,
    "2022-23": 3,
    "2023-24": 4,
    "2024-25": 5
}

fullData["SeasonOrder"] = fullData["Season"].map(seasonOrder)

#group all rows belonging to a team together, then count seasons played, average positions, etc. which will all be
#a part of the KPI summary (Key Performance Indicator)
kpiSummary = fullData.groupby("Team").agg(
    SeasonsPlayed=("Season","count"),
    AveragePosition=("Position","mean"),
    AveragePoints=("Points","mean"),
    TotalWins=("Wins","sum"),
    TotalLosses=("Losses","sum"),
    TotalGoalsFor=("GoalsFor","sum"),
    TotalGoalsAgainst=("GoalsAgainst","sum"),
    AverageGoalDifference=("GoalDifference","mean")
).reset_index()

kpiSummary = kpiSummary.sort_values(by="AveragePoints",ascending=False)


fullData = fullData.sort_values(by=["Team","SeasonOrder"])

#creating types of data in the league's list, which will be our KRI (Key Risk Indicator)
#after some testing, we realise that some teams hadn't been in the league for consecutive seasons, which
#affected our monitoring, so we resolved the issue with the season gap mechanism
fullData["PreviousSeasonOrder"] = fullData.groupby("Team")["SeasonOrder"].shift(1)
fullData["PreviousPosition"] = fullData.groupby("Team")["Position"].shift(1)
fullData["PreviousPoints"] = fullData.groupby("Team")["Points"].shift(1)

fullData["SeasonGap"] = fullData["SeasonOrder"] - fullData["PreviousSeasonOrder"]

fullData["PositionDrop"] = 0
fullData["PointDropPercentage"] = 0

continuousSeason = fullData["SeasonGap"] == 1

fullData.loc[continuousSeason, "PositionDrop"] = (
    fullData["Position"] - fullData["PreviousPosition"]
)

fullData.loc[continuousSeason, "PointDropPercentage"] = (
    (fullData["PreviousPoints"] - fullData["Points"]) / (fullData["PreviousPoints"] )
) * 100

#making a new section for the alerts, this is all part of KRI (Key Risk Indicator) monitoring
alertsList = fullData[
    (fullData["PositionDrop"] >= 4)
    | (fullData["PointDropPercentage"] >= 20)
    | ((fullData["Position"] > 15) & (fullData["Position"] <= 17) )
    | (fullData["Position"] >= 18)
    | (fullData["GoalDifference"] < 0)
].copy()

alertsList["RiskReason"] = ""

alertsList.loc[alertsList["PositionDrop"] >= 4, "RiskReason"] += "Major position drop; "
alertsList.loc[alertsList["PointDropPercentage"] >= 20, "RiskReason"] += "Points dropped over 20%;"
alertsList.loc[(alertsList["Position"] > 15) & (alertsList["Position"] <= 17), "RiskReason"] += "Nearing relegation zone; "
alertsList.loc[alertsList["Position"] >= 18, "RiskReason"] += "Relegation zone; "
alertsList.loc[alertsList["GoalDifference"] < 0, "RiskReason"] += "Negative goal difference; "



#here we select the top 10 teams to show they are the best performing with a table
topTeams = kpiSummary.head(10)

#here we have a table for riskiest teams
riskByTeam = alertsList.groupby("Team").size().reset_index(name="NumberOfAlerts")
riskByTeam = riskByTeam.sort_values(by="NumberOfAlerts",ascending=False)

#here we get a table of the seasons that had the highest risk (ie: problems)
riskBySeason = alertsList.groupby("Season").size().reset_index(name="TotalAlerts")

outputPath = Path("reports/LaLiga_KPI_KRI_Reports.xlsx")
#here we use openpyxl to write all of our new data onto a new excel report
with pd.ExcelWriter(outputPath, engine="openpyxl") as writer:
    fullData.to_excel(writer, sheet_name="CleanData", index=False)
    kpiSummary.to_excel(writer,sheet_name="KPISummary", index=False)
    alertsList.to_excel(writer, sheet_name="KRIAlerts", index=False)
    topTeams.to_excel(writer, sheet_name="TopTeams", index=False)
    riskByTeam.to_excel(writer, sheet_name="RiskByTeam", index=False)
    riskBySeason.to_excel(writer, sheet_name="RiskBySeason", index=False)
