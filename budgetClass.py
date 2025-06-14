from readExcelClass import ReadExcelClass
import pandas as pd


class BudgetClass:
    def __init__(self, df_class=ReadExcelClass()):
        self.Jan = MonthlyBudgetClass()
        self.Feb = MonthlyBudgetClass()
        self.Mar = MonthlyBudgetClass()
        self.Apr = MonthlyBudgetClass()
        self.May = MonthlyBudgetClass()
        self.Jun = MonthlyBudgetClass()
        self.Jul = MonthlyBudgetClass()
        self.Aug = MonthlyBudgetClass()
        self.Sep = MonthlyBudgetClass()
        self.Oct = MonthlyBudgetClass()
        self.Nov = MonthlyBudgetClass()
        self.Dec = MonthlyBudgetClass()
        self.used = MonthlyBudgetClass()
        self.left = MonthlyBudgetClass()
        self.df = df_class

    def get_monthly_budget(self):
        for i in range(self.df.get_line_num()):
            if 0 <= i <= 5 and 'project based' not in str(self.df.dataframe.loc[i, 'Frequency']):
                if pd.notna(self.df.dataframe.loc[i, 'total 2023 (0+12)']):
                    self.Jan.frame_contract_operations += self.df.dataframe.loc[i, 'total 2023 (0+12)']
                if pd.notna(self.df.dataframe.loc[i, 'total 2023(3+9)']):
                    self.Mar.frame_contract_operations += self.df.dataframe.loc[i, 'total 2023(3+9)']
                if pd.notna(self.df.dataframe.loc[i, 'total 2023(4+8)']):
                    self.Apr.frame_contract_operations += self.df.dataframe.loc[i, 'total 2023(4+8)']
                if pd.notna(self.df.dataframe.loc[i, 'total 2023(5+7)']):
                    self.May.frame_contract_operations += self.df.dataframe.loc[i, 'total 2023(5+7)']
                if pd.notna(self.df.dataframe.loc[i, 'total 2023(6+6)']):
                    self.Jun.frame_contract_operations += self.df.dataframe.loc[i, 'total 2023(6+6)']
                if pd.notna(self.df.dataframe.loc[i, 'total 2023(7+5)']):
                    self.Jul.frame_contract_operations += self.df.dataframe.loc[i, 'total 2023(7+5)']
                if pd.notna(self.df.dataframe.loc[i, 'total 2023(8+4)']):
                    self.Aug.frame_contract_operations += self.df.dataframe.loc[i, 'total 2023(8+4)']
                if pd.notna(self.df.dataframe.loc[i, 'total 2023(9+3)']):
                    self.Sep.frame_contract_operations += self.df.dataframe.loc[i, 'total 2023(9+3)']
                if pd.notna(self.df.dataframe.loc[i, 'USED']):
                    self.used.frame_contract_operations += self.df.dataframe.loc[i, 'USED']
                if pd.notna(self.df.dataframe.loc[i, 'LEFT 1']):
                    self.left.frame_contract_operations += self.df.dataframe.loc[i, 'LEFT 1']
            elif 0 <= i <= 5 and 'project based' in str(self.df.dataframe.loc[i, 'Frequency']):
                if pd.notna(self.df.dataframe.loc[i, 'total 2023 (0+12)']):
                    self.Jan.frame_contract_1time_cost += self.df.dataframe.loc[i, 'total 2023 (0+12)']
                if pd.notna(self.df.dataframe.loc[i, 'total 2023(3+9)']):
                    self.Mar.frame_contract_1time_cost += self.df.dataframe.loc[i, 'total 2023(3+9)']
                if pd.notna(self.df.dataframe.loc[i, 'total 2023(4+8)']):
                    self.Apr.frame_contract_1time_cost += self.df.dataframe.loc[i, 'total 2023(4+8)']
                if pd.notna(self.df.dataframe.loc[i, 'total 2023(5+7)']):
                    self.May.frame_contract_1time_cost += self.df.dataframe.loc[i, 'total 2023(5+7)']
                if pd.notna(self.df.dataframe.loc[i, 'total 2023(6+6)']):
                    self.Jun.frame_contract_1time_cost += self.df.dataframe.loc[i, 'total 2023(6+6)']
                if pd.notna(self.df.dataframe.loc[i, 'total 2023(7+5)']):
                    self.Jul.frame_contract_1time_cost += self.df.dataframe.loc[i, 'total 2023(7+5)']
                if pd.notna(self.df.dataframe.loc[i, 'total 2023(8+4)']):
                    self.Aug.frame_contract_1time_cost += self.df.dataframe.loc[i, 'total 2023(8+4)']
                if pd.notna(self.df.dataframe.loc[i, 'total 2023(9+3)']):
                    self.Sep.frame_contract_1time_cost += self.df.dataframe.loc[i, 'total 2023(9+3)']
                if pd.notna(self.df.dataframe.loc[i, 'USED']):
                    self.used.frame_contract_1time_cost += self.df.dataframe.loc[i, 'USED']
                if pd.notna(self.df.dataframe.loc[i, 'LEFT 1']):
                    self.left.frame_contract_1time_cost += self.df.dataframe.loc[i, 'LEFT 1']
            elif 'project based' not in str(self.df.dataframe.loc[i, 'Frequency']):
                if pd.notna(self.df.dataframe.loc[i, 'total 2023 (0+12)']):
                    self.Jan.other_project_operations += self.df.dataframe.loc[i, 'total 2023 (0+12)']
                if pd.notna(self.df.dataframe.loc[i, 'total 2023(3+9)']):
                    self.Mar.other_project_operations += self.df.dataframe.loc[i, 'total 2023(3+9)']
                if pd.notna(self.df.dataframe.loc[i, 'total 2023(4+8)']):
                    self.Apr.other_project_operations += self.df.dataframe.loc[i, 'total 2023(4+8)']
                if pd.notna(self.df.dataframe.loc[i, 'total 2023(5+7)']):
                    self.May.other_project_operations += self.df.dataframe.loc[i, 'total 2023(5+7)']
                if pd.notna(self.df.dataframe.loc[i, 'total 2023(6+6)']):
                    self.Jun.other_project_operations += self.df.dataframe.loc[i, 'total 2023(6+6)']
                if pd.notna(self.df.dataframe.loc[i, 'total 2023(7+5)']):
                    self.Jul.other_project_operations += self.df.dataframe.loc[i, 'total 2023(7+5)']
                if pd.notna(self.df.dataframe.loc[i, 'total 2023(8+4)']):
                    self.Aug.other_project_operations += self.df.dataframe.loc[i, 'total 2023(8+4)']
                if pd.notna(self.df.dataframe.loc[i, 'total 2023(9+3)']):
                    self.Sep.other_project_operations += self.df.dataframe.loc[i, 'total 2023(9+3)']
                if pd.notna(self.df.dataframe.loc[i, 'USED']):
                    self.used.other_project_operations += self.df.dataframe.loc[i, 'USED']
                if pd.notna(self.df.dataframe.loc[i, 'LEFT 1']):
                    self.left.other_project_operations += self.df.dataframe.loc[i, 'LEFT 1']
            elif 'project based' in str(self.df.dataframe.loc[i, 'Frequency']):
                if pd.notna(self.df.dataframe.loc[i, 'total 2023 (0+12)']):
                    self.Jan.other_project_1time_cost += self.df.dataframe.loc[i, 'total 2023 (0+12)']
                if pd.notna(self.df.dataframe.loc[i, 'total 2023(3+9)']):
                    self.Mar.other_project_1time_cost += self.df.dataframe.loc[i, 'total 2023(3+9)']
                if pd.notna(self.df.dataframe.loc[i, 'total 2023(4+8)']):
                    self.Apr.other_project_1time_cost += self.df.dataframe.loc[i, 'total 2023(4+8)']
                if pd.notna(self.df.dataframe.loc[i, 'total 2023(5+7)']):
                    self.May.other_project_1time_cost += self.df.dataframe.loc[i, 'total 2023(5+7)']
                if pd.notna(self.df.dataframe.loc[i, 'total 2023(6+6)']):
                    self.Jun.other_project_1time_cost += self.df.dataframe.loc[i, 'total 2023(6+6)']
                if pd.notna(self.df.dataframe.loc[i, 'total 2023(7+5)']):
                    self.Jul.other_project_1time_cost += self.df.dataframe.loc[i, 'total 2023(7+5)']
                if pd.notna(self.df.dataframe.loc[i, 'total 2023(8+4)']):
                    self.Aug.other_project_1time_cost += self.df.dataframe.loc[i, 'total 2023(8+4)']
                if pd.notna(self.df.dataframe.loc[i, 'total 2023(9+3)']):
                    self.Sep.other_project_1time_cost += self.df.dataframe.loc[i, 'total 2023(9+3)']
                if pd.notna(self.df.dataframe.loc[i, 'USED']):
                    self.used.other_project_1time_cost += self.df.dataframe.loc[i, 'USED']
                if pd.notna(self.df.dataframe.loc[i, 'LEFT 1']):
                    self.left.other_project_1time_cost += self.df.dataframe.loc[i, 'LEFT 1']


class MonthlyBudgetClass:
    def __init__(self):
        self.frame_contract_1time_cost = 0
        self.frame_contract_operations = 0
        self.other_project_1time_cost = 0
        self.other_project_operations = 0

    def set_frame_contract_1time_cost(self, num):
        self.frame_contract_1time_cost = num

    def set_frame_contract_operations(self, num):
        self.frame_contract_operations = num

    def set_other_project_1time_cost(self, num):
        self.other_project_1time_cost = num

    def set_other_project_operations(self, num):
        self.other_project_operations = num

