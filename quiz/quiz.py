# python quiz.py year_filename.xlsx count
# python quiz.py 2020summer.xlsx 5

from openpyxl import load_workbook
import random, sys, glob
import pandas as pd, numpy as np

def checkFilename(filename):
	if ".xlsx" not in filename:
		filename = filename + "*.xlsx"
	return glob.glob(filename)[0]

def getCandidatesWith(data, score):
	candidates = []
	for row in data:
		if row[2] == score:
			candidates.append(row[3])
	return candidates

def getAndPrintScoreDistribution(data):
	print('== 점수 분포 ==')
	distribution = data.value_counts(sort=False)[::-1]
	print(distribution)
	return distribution

def printNames(selected):
	for name in selected:
		print('-', name)

if len(sys.argv) < 3:
	print("사용법: python quiz.py filename.xlsx count")
else:
	filename = checkFilename(sys.argv[1])
	print(filename, "is reading...")

	file = pd.read_excel(filename)
	data = file.values

	scores = file['점수'].unique()
	scores = np.sort(scores)[::-1]
	distribution = getAndPrintScoreDistribution(file['점수'])

	count = int(sys.argv[2])
	candidates = getCandidatesWith(data, scores[0])
	random.shuffle(candidates)

	print()
	print('== 당첨자 ==')
	if distribution[scores[0]] >= count:
		firstprize = candidates[:count]
		printNames(firstprize)
		# end
	else:
		firstprize = candidates
		printNames(firstprize)

		sec_candidates = getCandidatesWith(data, scores[1])
		sec_candidates = list(set(sec_candidates) - set(candidates))
		random.shuffle(sec_candidates)
		sec_prize = sec_candidates[:count-len(firstprize)]
		printNames(sec_prize)