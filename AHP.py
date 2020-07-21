#! usr/env/python3
'''
* Program to automatically rank the smartphone models based on the user given feature preference rating.

* This program has many variables that need to be changed when being used with any other spreadsheet other than "Smartphone Comparison Sheet.xlsx"

* The header information and the relative positionig of initial rows and columns must not be changed!! (the starting row and col index must be adjusted accordingly if done so!)

* You can,however, add more comparison features as well as alternatives (add more columns and rows) to the xlsx file and the program will handle that automatically, as long as the original structure of the sheet is intact (just adding similar rows and columns DO NOT AFFECT the working of this program!!)

* If any extra feature is added that is to be minimized, then the name of that feature must be added to 'minFeatures' list in the 'convert_to_satty' function.

* There are some heuristics that have been employed when scaling values of different bounds in 'convert_to_satty' function. You can change them if you want to modify the reltive differences of the resulting values.

* If the total features are more than 15, then the program will work without any problem, however, for the calcualtion of correct consistency index, the 'randomConsistencyIndex' needs to be updated. (The perfermonce matrix formed will be consistent in most of the cases.)

Author info:
Name    : Nimesh Khandelwal
E-Mail  : nimesh6798@gmail.com
github  : github.com/nimesh00

'''


import xlrd
import numpy as np
from collections import Counter

workbookName = "Smartphone Comparison Sheet.xlsx"
sheetName = "Sheet1"



def filter_features(table, features, models):
    filtered_table = []
    filtered_features = []
    partial_filtered_models = []
    for i in range(len(table)):
        num_zeros = 0
        for j in range(len(table[0])):
            if table[i][j] == 0:
                num_zeros += 1
        if num_zeros > 5:
            continue
        else:
            partial_filtered_models.append(models[i])
            filtered_table.append(table[i][table[i] != 0])
            filtered_features.append(features[table[i] != 0])

    # DO NOT MODIFY BELOW THIS LINE (DEVELOPER COMMENT)
    x = [len(i) for i in filtered_table]
    c = Counter(x)
    c = dict(c)
    most_common = max(c, key = c.get)
    uniform_table = []
    filtered_models = []
    for i in range(len(filtered_table)):
        if len(filtered_table[i]) == most_common:
            uniform_table.append(filtered_table[i])
            filtered_models.append(partial_filtered_models[i])
    # DO NOT TOUCH ABOVE THIS LINE (DEVELOPER COMMENT)
    for feature in filtered_features:
        if len(feature) == most_common:
            filtered_features = list(feature)
            break
    return np.stack(np.array(uniform_table)), filtered_features, filtered_models

def convert_to_satty(table, features):
    minFeatures = ['Cost', 'Weight']
    scaleDownResolution = 10
    scaleUpResolution  = 3

    # DO NOT CHANGE BELOW THIS (DEVELOPER COMMENT)
    transpose_table = table.T
    transpose_satty_table = np.zeros(transpose_table.shape)
    for i in range(len(transpose_table)):
        max_val = np.max(transpose_table[i])
        min_val = np.min(transpose_table[i])
        if features[i] not in minFeatures:
            if min_val > int(max_val / scaleDownResolution):
                min_val = int(max_val / scaleDownResolution)
        else:
            if max_val < int(min_val * scaleUpResolution):
                max_val = int(min_val * scaleUpResolution)
        for j in range(len(transpose_table[0])):
            if features[i] not in minFeatures:
                transpose_satty_table[i][j] = int(9 * (transpose_table[i][j] - min_val) / (max_val - min_val)) + 1
            else:
                transpose_satty_table[i][j] = int(9 * (transpose_table[i][j] - max_val) / (min_val - max_val)) + 1
    return transpose_satty_table.T

def evaluateFeature(feature):
    performanceMatrix = np.zeros((len(feature), len(feature)))
    for i in range(len(feature)):
        for j in range(len(feature)):
            if performanceMatrix[i][j] != 0:
                continue
            if feature[i] == feature[j]:
                performanceMatrix[i][j] = 1
            elif feature[i] - feature[j] > 0:
                performanceMatrix[i][j] = feature[i] - feature[j] + 1
                performanceMatrix[j][i] = 1 / (feature[i] - feature[j] + 1)
    performanceMatrix = performanceMatrix / performanceMatrix.sum(axis = 0)
    performanceVector = performanceMatrix.sum(axis = 1)
    return performanceVector

def evaluateFeaturePreference(feature):
    performanceMatrix = np.zeros((len(feature), len(feature)))
    for i in range(len(feature)):
        for j in range(len(feature)):
            if performanceMatrix[i][j] != 0:
                continue
            if feature[i] == feature[j]:
                performanceMatrix[i][j] = 1
            elif feature[i] - feature[j] > 0:
                performanceMatrix[i][j] = feature[i] - feature[j] + 1
                performanceMatrix[j][i] = 1 / (feature[i] - feature[j] + 1)
    eigenMatrix = performanceMatrix / performanceMatrix.sum(axis = 0)
    eigenVector = eigenMatrix.mean(axis = 1)
    maxEigenValue = np.dot(eigenVector, performanceMatrix.sum(axis = 0))
    return np.array([eigenVector]).T, maxEigenValue

def evaluateCriteria(table):
    # transforming the matrix to get feature as rows
    feature_table = table.T
    avgCriteriaTableTranspose = np.zeros(feature_table.shape)
    for i, feature in enumerate(feature_table):
        avgCriteriaTableTranspose[i] = evaluateFeature(feature)
    return avgCriteriaTableTranspose.T


def randomConsistencyIndex(N):
    # Just need to extend this dict for more than 15 features
    ri_dict = {3: 0.52, 4: 0.89, 5: 1.11, 6: 1.25, 7: 1.35, 8: 1.40, 9: 1.45, 10: 1.49, 11: 1.52, 12: 1.54, 13: 1.56, 14: 1.58, 15: 1.59}
    return ri_dict[N]


def checkForConsistency(maxEigenValue, N):
    consistencyIndex = (maxEigenValue - N) / (N - 1)
    RI = randomConsistencyIndex(N)
    consistencyRatio = consistencyIndex / RI
    return consistencyRatio

def main():
    wb = xlrd.open_workbook(workbookName)
    sheet = wb.sheet_by_index(0)

    features = []
    feature_table = []
    models = []
    for j in range(3, sheet.ncols):
        feature_table.append([sheet.cell_value(2, j)])
        # print(feature_table)
        for i in range(3, sheet.nrows):
            cellValue = sheet.cell_value(i, j)
            if cellValue == xlrd.empty_cell.value:
                cellValue = 0
            feature_table[j - 3].append(cellValue)
    for i in range(3, sheet.nrows):
        cellValue = sheet.cell_value(i, 2)
        if cellValue == xlrd.empty_cell.value:
            cellValue = 0
        models.append(cellValue)
    feature_table = np.array(feature_table)
    # print(features)
    # feature_table = feature_table[:][1:]
    features = feature_table[:, 0]

    # feature information table with each column representing each feature and each row representing each alternative
    feature_table = feature_table[:, 1:].T.astype(np.float)

    # remove the row, column pair containing zero (no information)
    # print("feature table: \n", feature_table)
    filtered_table, filtered_features, filtered_models = filter_features(feature_table, features, models)
    # print("filtered features:, \n", filtered_features)
    # print("filtered table: \n", filtered_table)
    # print("filtered models: \n", filtered_models)

    # rows: Alternatives; columns: features
    # higher value is preferred
    satty_table = convert_to_satty(filtered_table, filtered_features)
    # print("satty table: \n", satty_table)
    avgCriteriaTable = evaluateCriteria(satty_table)
    # print("Average Criteria Matrix: \n", avgCriteriaTable)

    # Depends strictly on personal preference!!
    
    print("Enter the relative preference rating (1-10) for each feature given in the list below.")
    print("(You do not need to rank the features, just how much you prefer the given feature!!)")
    print("Enter your response as comma-separated values in a single line and press enter: ")
    print("Eg: 9, 6, 5, 4, ......")
    print("Feature List: ", filtered_features)
    print("Total {} features: ".format(len(filtered_features)))
    preferenceRating = []
    while True:
        try:
            preferenceRating = input("Type Here: ")
            preferenceRating = preferenceRating.split(",")
            preferenceRating = [int(num) for num in preferenceRating]
            if len(preferenceRating) != len(filtered_features):
                print("You entered for {} features!!".format(len(preferenceRating)))
                continue
            break
        except ValueError:
            print("Invalid value entered!!")
        except:
            print("Unexcepted error!!")
            import sys
            sys.exit()

    preferenceVector, maxEigenValue = evaluateFeaturePreference(preferenceRating)
    # print("preference vector, maxeigenvalue: ", preferenceVector, maxEigenValue)
    consistencyRatio = checkForConsistency(maxEigenValue, len(filtered_features))
    print("Consistency Ratio: ", consistencyRatio)
    if consistencyRatio > 0.1:
        print("Some preferences are not consistent with others, try again with different values!!")
    # print(preferenceVector)
    weightedPreferenceVector = (avgCriteriaTable @ preferenceVector).T[0]
    print(filtered_models)
    print(weightedPreferenceVector)
    rankedModels = [model for _,model in sorted(zip(weightedPreferenceVector, filtered_models), reverse=True)]
    print("Ranking based on the Calculations: \n", rankedModels)


if __name__ == "__main__":
    main()