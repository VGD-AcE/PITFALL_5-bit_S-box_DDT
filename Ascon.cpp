

#include <stdio.h>
#include <windows.h>
#include <iostream>
#include <string>
#include <vector>
#include <set>
#include <algorithm>
#include <iterator>
#include "libxl.h"

using namespace libxl;

char S[32] = { 4,11,31,20,26,21,9,2,27,5,8,18,29,3,6,28,
			   30,19,7,14,0,13,17,24,16,12,1,25,22,10,15,23 };     //Ascon 5-bit S-box

// Function: Find the intersection of two sets
std::vector<int> calculateIntersection(const std::vector<int>& set1, const std::vector<int>& set2) {

	std::vector<int> intersection;

	std::set_intersection(
		set1.begin(), set1.end(),
		set2.begin(), set2.end(),
		std::back_inserter(intersection)
	);

	return intersection;
}

int main()
{
	// 1. Define a three-dimensional variable-length array: 32*4*?
	std::vector<std::vector<std::vector<int>>> differ_LSB_2(32, std::vector<std::vector<int>>(4));

	// The serial number used to store the column containing the element Num (required in 3)
	int columnsWith_Num[32][32] = { 0 };

	//	int flag;
	int out;
	// Solving difference equations.
	for (int i = 1; i < 32; i++) {
		for (int j = 1; j < 32; j++) {
			for (int in = 0; in < 32; in++) {
				out = S[in] ^ S[i ^ in];
				if (j == out) {
					differ_LSB_2[i][j % 4].push_back(in);
				}
			}
			std::sort(differ_LSB_2[i][j % 4].begin(), differ_LSB_2[i][j % 4].end()); // Reorder the set
		}
	}

	// 2. Create an Excel document object
	libxl::Book* book = xlCreateBook();
	book->setKey(L"libxl", L"windows-28232b0208c4ee0369ba6e68abv6v5i3");
	if (book) {

		libxl::Sheet* sheet = book->addSheet(L"Sheet1");// Add a worksheet

		sheet->setCol(1, 0, 30); // Set table column width
		sheet->setCol(1, 1, 35);
		sheet->setCol(1, 2, 35);
		sheet->setCol(1, 3, 35);
		sheet->setCol(1, 4, 35);

		sheet->writeStr(0, 1, L"The lowest two bits of the S-box output difference are 00"); // Set table title
		sheet->writeStr(0, 2, L"The lowest two bits of the S-box output difference are 01");
		sheet->writeStr(0, 3, L"The lowest two bits of the S-box output difference are 10");
		sheet->writeStr(0, 4, L"The lowest two bits of the S-box output difference are 11");

		// Print the entire  DDT table (output on the terminal while printing and storing in the excel table).
		for (int i = 1; i < 32; ++i) {
			for (int j = 0; j < 4; ++j) {
				std::cout << "differ_LSB_2[" << i << "][" << j << "]: ";
				std::wstring cellData;

				// If the current cell contains the element Num, record the column number.
				for (int Num = 0; Num < 32; ++Num) {
					if (std::find(differ_LSB_2[i][j].begin(), differ_LSB_2[i][j].end(), Num) != differ_LSB_2[i][j].end()) {
						columnsWith_Num[Num][i] = j;
					}
				}

				// Print the elements in the current cell.
				for (int k = 0; k < differ_LSB_2[i][j].size(); ++k) {
					std::cout << differ_LSB_2[i][j][k] << " ";

					// Concatenate the elements in the current cell into a string separated by commas.
					cellData += std::to_wstring(differ_LSB_2[i][j][k]);
					if (k < differ_LSB_2[i][j].size() - 1) {
						cellData += L",";
					}
				}

				// Write data in Excel cells.
				sheet->writeStr(i, j + 1, cellData.c_str());

				// Print out the data in the terminal.
				std::cout << std::endl;
			}

			// Print leftmost column in Excel.
			std::wstring newNumberStr = std::to_wstring(i);
			std::wstring currentValue = L"The input difference (fault value) is:" + newNumberStr;
			sheet->writeStr(i, 0, currentValue.c_str());
		}

		// Save Excel file.
		book->save(L"output.xlsx");

		// Release resources.
		book->release();
		std::cout << "Excel file generated successfully!" << std::endl;
	}
	else {
		std::cerr << "Unable to create Excel document object!" << std::endl;
	}

	// 3. Used to store the set of elements in cells in which each row contains element Num.
	for (int Num = 0; Num < 32; ++Num) {

		// Output the ordinal number of the column containing the element Num.
		std::cout << "Columns with element " << Num << ":";
		columnsWith_Num[Num][0] = 4;
		for (int column : columnsWith_Num[Num]) {
			std::cout << column << " ";
		}
		std::cout << std::endl;

		// Calculate intersection
		std::vector<int> result = calculateIntersection(differ_LSB_2[1][columnsWith_Num[Num][1]], differ_LSB_2[2][columnsWith_Num[Num][2]]);
		for (int i = 3; i < 32; ++i) {
			result = calculateIntersection(result, differ_LSB_2[i][columnsWith_Num[Num][i]]);
		}
		std::cout << "Intersection of sets containing " << Num << " in each row:";
		for (int element : result) {
			std::cout << element << " ";
		}
		std::cout << std::endl;
	}


	system("pause");
	return 0;
}