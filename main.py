import xlrd
import matplotlib.pyplot as plt
import numpy as np
import math
import statistics

class GeometricBrownianMotion(object):
    def __init__(self, initial_price, drift, volatility, dt, T):
        self.current_price = initial_price
        self.initial_price = initial_price
        self.drift = drift
        self.volatility = volatility
        self.dt = dt
        self.T = T
        self.prices = []
        self.simulate_paths()

    def simulate_paths(self):
        while self.T - self.dt > 0:
            dWt = np.random.normal(0, math.sqrt(self.dt))
            dYt = self.drift*self.dt + self.volatility*dWt
            self.current_price += dYt
            self.prices.append(self.current_price)
            self.T -= self.dt


def get_data(file, begin, end):
    wb = xlrd.open_workbook(file)
    sheet = wb.sheet_by_index(0)

    data = []
    for i in range(begin, end):
        data.append(sheet.cell_value(i, 1))
    return data


def get_average(data, path_amount):
    paths = path_amount
    initial_price = data[0]
    drift = .08
    volatility = .1
    dt = 1 / 365
    T = 1
    price_paths = []

    for i in range(0, paths):
        price_paths.append(GeometricBrownianMotion(initial_price, drift, volatility, dt, T).prices)

    values = price_paths[0]
    return statistics.mean(values), price_paths


class AllAverages(object):
    def __init__(self, xlxs_files, path_amount, column=17, row=530):
        self.xlxs_files = xlxs_files
        self.path_amount = path_amount
        self.column = column
        self.row = row
        self.averages = {}

    def get_result(self):
        for name, xlxs_file in self.xlxs_files.items():
            total_average_value, all_paths = get_average(get_data(xlxs_file, self.column, self.row), self.path_amount)
            self.averages[name] = round(total_average_value, 2)
            self.get_plot(all_paths)
        return self.averages

    @staticmethod
    def get_plot(plot_data):
        for price_path in plot_data:
            plt.plot(price_path)
        plt.show()


if _name__== '__main_':
    all_excel_files = {
        # Dict contains:
            # Key: name of stock
            # Value: Excel file
        # ... Add new Excel File
        # If data does not come from Refenitiv tool adjust column and row args un AllAverages to match the required cells
         }

    solution = AllAverages(all_excel_files, path_amount=100)
    solution = solution.get_result()
    print(solution)
