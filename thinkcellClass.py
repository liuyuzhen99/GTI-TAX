import pandas as pd
from thinkcell import Thinkcell
from budgetClass import BudgetClass
import os


class ThinkcellClass:
    def __init__(self, template_name='', budget=BudgetClass()):
        self.template_name = template_name
        self.tc = Thinkcell()
        self.chart = []
        self.budget = budget
        self.filename = os.getcwd() + '\\output.ppttc'

    def add_template(self):
        self.tc.add_template(self.template_name)

    def add_chart(self):
        chart1_1 = {
            "chart_name": "chart1.1",
            "dataframe": pd.DataFrame(
                columns=["category", "1+11", "2+10", "3+9", "4+8", "5+7", "6+6", "7+5", "8+4", "9+3", "10+2",
                         "11+1", "12+0"],
                data=[
                    ["Other operation costs(yearly)", round(self.budget.Jan.other_project_operations / 1000),
                     round(self.budget.Feb.other_project_operations / 1000),
                     round(self.budget.Mar.other_project_operations / 1000),
                     round(self.budget.Apr.other_project_operations / 1000),
                     round(self.budget.May.other_project_operations / 1000),
                     round(self.budget.Jun.other_project_operations / 1000),
                     round(self.budget.Jul.other_project_operations / 1000),
                     round(self.budget.Aug.other_project_operations / 1000),
                     round(self.budget.Sep.other_project_operations / 1000),
                     round(self.budget.Oct.other_project_operations / 1000),
                     round(self.budget.Nov.other_project_operations / 1000),
                     round(self.budget.Dec.other_project_operations / 1000)],
                    ["Frame Contracts Operations(yearly)", round(self.budget.Jan.frame_contract_operations / 1000),
                     round(self.budget.Feb.frame_contract_operations / 1000),
                     round(self.budget.Mar.frame_contract_operations / 1000),
                     round(self.budget.Apr.frame_contract_operations / 1000),
                     round(self.budget.May.frame_contract_operations / 1000),
                     round(self.budget.Jun.frame_contract_operations / 1000),
                     round(self.budget.Jul.frame_contract_operations / 1000),
                     round(self.budget.Aug.frame_contract_operations / 1000),
                     round(self.budget.Sep.frame_contract_operations / 1000),
                     round(self.budget.Oct.frame_contract_operations / 1000),
                     round(self.budget.Nov.frame_contract_operations / 1000),
                     round(self.budget.Dec.frame_contract_operations / 1000)],
                    ["Other/Project based(1 time)", round(self.budget.Jan.other_project_1time_cost / 1000),
                     round(self.budget.Feb.other_project_1time_cost / 1000),
                     round(self.budget.Mar.other_project_1time_cost / 1000),
                     round(self.budget.Apr.other_project_1time_cost / 1000),
                     round(self.budget.May.other_project_1time_cost / 1000),
                     round(self.budget.Jun.other_project_1time_cost / 1000),
                     round(self.budget.Jul.other_project_1time_cost / 1000),
                     round(self.budget.Aug.other_project_1time_cost / 1000),
                     round(self.budget.Sep.other_project_1time_cost / 1000),
                     round(self.budget.Oct.other_project_1time_cost / 1000),
                     round(self.budget.Nov.other_project_1time_cost / 1000),
                     round(self.budget.Dec.other_project_1time_cost / 1000)],
                    ["Frame Contract(1 time)", round(self.budget.Jan.frame_contract_1time_cost / 1000),
                     round(self.budget.Feb.frame_contract_1time_cost / 1000),
                     round(self.budget.Mar.frame_contract_1time_cost / 1000),
                     round(self.budget.Apr.frame_contract_1time_cost / 1000),
                     round(self.budget.May.frame_contract_1time_cost / 1000),
                     round(self.budget.Jun.frame_contract_1time_cost / 1000),
                     round(self.budget.Jul.frame_contract_1time_cost / 1000),
                     round(self.budget.Aug.frame_contract_1time_cost / 1000),
                     round(self.budget.Sep.frame_contract_1time_cost / 1000),
                     round(self.budget.Oct.frame_contract_1time_cost / 1000),
                     round(self.budget.Nov.frame_contract_1time_cost / 1000),
                     round(self.budget.Dec.frame_contract_1time_cost / 1000)],
                ],
            ),
        }
        self.chart.append(chart1_1)
        chart1_2 = {
            "chart_name": "chart1.2",
            "dataframe": pd.DataFrame(
                columns=["category", "used", "left"],
                data=[
                    ["other operation costs(yearly)", round(self.budget.used.other_project_operations / 1000),
                     round(self.budget.left.other_project_operations / 1000)],
                    ["Frame Contracts Operations(yearly)", round(self.budget.used.frame_contract_operations / 1000),
                     round(self.budget.left.frame_contract_operations / 1000)],
                    ["Other/Project based(1 time)", round(self.budget.used.other_project_1time_cost / 1000),
                     round(self.budget.left.other_project_1time_cost / 1000)],
                    ["Frame Contracts(1 time)", round(self.budget.used.frame_contract_1time_cost / 1000),
                     round(self.budget.left.frame_contract_1time_cost / 1000)],
                ],
            ),
        }
        self.chart.append(chart1_2)
        chart2_1 = {
            "chart_name": "chart2.1",
            "dataframe": pd.DataFrame(
                columns=["category", "1+11", "2+10", "3+9", "4+8", "5+7", "6+6", "7+5", "8+4", "9+3", "10+2",
                         "11+1", "12+0"],
                data=[
                    ["Entry cost", round(self.budget.Jan.frame_contract_1time_cost / 1000),
                     round(self.budget.Feb.frame_contract_1time_cost / 1000),
                     round(self.budget.Mar.frame_contract_1time_cost / 1000),
                     round(self.budget.Apr.frame_contract_1time_cost / 1000),
                     round(self.budget.May.frame_contract_1time_cost / 1000),
                     round(self.budget.Jun.frame_contract_1time_cost / 1000),
                     round(self.budget.Jul.frame_contract_1time_cost / 1000),
                     round(self.budget.Aug.frame_contract_1time_cost / 1000),
                     round(self.budget.Sep.frame_contract_1time_cost / 1000),
                     round(self.budget.Oct.frame_contract_1time_cost / 1000),
                     round(self.budget.Nov.frame_contract_1time_cost / 1000),
                     round(self.budget.Dec.frame_contract_1time_cost / 1000)],
                    ["Operations", round(self.budget.Jan.frame_contract_operations / 1000),
                     round(self.budget.Feb.frame_contract_operations / 1000),
                     round(self.budget.Mar.frame_contract_operations / 1000),
                     round(self.budget.Apr.frame_contract_operations / 1000),
                     round(self.budget.May.frame_contract_operations / 1000),
                     round(self.budget.Jun.frame_contract_operations / 1000),
                     round(self.budget.Jul.frame_contract_operations / 1000),
                     round(self.budget.Aug.frame_contract_operations / 1000),
                     round(self.budget.Sep.frame_contract_operations / 1000),
                     round(self.budget.Oct.frame_contract_operations / 1000),
                     round(self.budget.Nov.frame_contract_operations / 1000),
                     round(self.budget.Dec.frame_contract_operations / 1000)],
                ],
            ),
        }
        self.chart.append(chart2_1)
        chart2_2 = {
            "chart_name": "chart2.2",
            "dataframe": pd.DataFrame(
                columns=["category", "used", "left"],
                data=[
                    ["Frame Contracts Operations(yearly)", round(self.budget.used.frame_contract_operations / 1000),
                     round(self.budget.left.frame_contract_operations / 1000)],
                    ["Frame Contracts(1 time)", round(self.budget.used.frame_contract_1time_cost / 1000),
                     round(self.budget.left.frame_contract_1time_cost / 1000)],
                ],
            ),
        }
        self.chart.append(chart2_2)
        chart3 = {
            "chart_name": "chart3",
            "dataframe": pd.DataFrame(
                columns=["category", "0+12", "potential risk", "Now"],
                data=[
                    ["", "e",
                     round((self.budget.Jan.frame_contract_1time_cost -
                            (self.budget.used.frame_contract_1time_cost + self.budget.left.frame_contract_1time_cost))
                           / 1000),
                     ""],
                    ["left", "", "", round(self.budget.left.frame_contract_1time_cost / 1000)],
                    ["used", "", "", round(self.budget.used.frame_contract_1time_cost / 1000)],
                ],
            ),
        }
        self.chart.append(chart3)
        chart4 = {
            "chart_name": "chart4",
            "dataframe": pd.DataFrame(
                columns=["category", "0+12", "potential risk", "Now"],
                data=[
                    ["", "e",
                     round((self.budget.Jan.frame_contract_operations -
                            (self.budget.used.frame_contract_operations + self.budget.left.frame_contract_operations))
                           / 1000),
                     ""],
                    ["left", "", "", round(self.budget.left.frame_contract_operations / 1000)],
                    ["used", "", "", round(self.budget.used.frame_contract_operations / 1000)],
                ],
            ),
        }
        self.chart.append(chart4)
        chart5_1 = {
            "chart_name": "chart5.1",
            "dataframe": pd.DataFrame(
                columns=["category", "1+11", "2+10", "3+9", "4+8", "5+7", "6+6", "7+5", "8+4", "9+3", "10+2",
                         "11+1", "12+0"],
                data=[
                    ["Entry cost", round(self.budget.Jan.other_project_1time_cost / 1000),
                     round(self.budget.Feb.other_project_1time_cost / 1000),
                     round(self.budget.Mar.other_project_1time_cost / 1000),
                     round(self.budget.Apr.other_project_1time_cost / 1000),
                     round(self.budget.May.other_project_1time_cost / 1000),
                     round(self.budget.Jun.other_project_1time_cost / 1000),
                     round(self.budget.Jul.other_project_1time_cost / 1000),
                     round(self.budget.Aug.other_project_1time_cost / 1000),
                     round(self.budget.Sep.other_project_1time_cost / 1000),
                     round(self.budget.Oct.other_project_1time_cost / 1000),
                     round(self.budget.Nov.other_project_1time_cost / 1000),
                     round(self.budget.Dec.other_project_1time_cost / 1000)],
                    ["Operations", round(self.budget.Jan.other_project_operations / 1000),
                     round(self.budget.Feb.other_project_operations / 1000),
                     round(self.budget.Mar.other_project_operations / 1000),
                     round(self.budget.Apr.other_project_operations / 1000),
                     round(self.budget.May.other_project_operations / 1000),
                     round(self.budget.Jun.other_project_operations / 1000),
                     round(self.budget.Jul.other_project_operations / 1000),
                     round(self.budget.Aug.other_project_operations / 1000),
                     round(self.budget.Sep.other_project_operations / 1000),
                     round(self.budget.Oct.other_project_operations / 1000),
                     round(self.budget.Nov.other_project_operations / 1000),
                     round(self.budget.Dec.other_project_operations / 1000)],
                ],
            ),
        }
        self.chart.append(chart5_1)
        chart5_2 = {
            "chart_name": "chart5.2",
            "dataframe": pd.DataFrame(
                columns=["category", "used", "left"],
                data=[
                    ["Frame Contracts Operations(yearly)", round(self.budget.used.other_project_operations / 1000),
                     round(self.budget.left.other_project_operations / 1000)],
                    ["Frame Contracts(1 time)", round(self.budget.used.other_project_1time_cost / 1000),
                     round(self.budget.left.other_project_1time_cost / 1000)],
                ],
            ),
        }
        self.chart.append(chart5_2)
        chart6 = {
            "chart_name": "chart6",
            "dataframe": pd.DataFrame(
                columns=["category", "0+12", "potential risk", "Now"],
                data=[
                    ["", "e",
                     round((self.budget.Jan.other_project_1time_cost -
                            (self.budget.used.other_project_1time_cost + self.budget.left.other_project_1time_cost))
                           / 1000),
                     ""],
                    ["left", "", "", round(self.budget.left.other_project_1time_cost / 1000)],
                    ["used", "", "", round(self.budget.used.other_project_1time_cost / 1000)],
                ],
            ),
        }
        self.chart.append(chart6)
        chart7 = {
            "chart_name": "chart7",
            "dataframe": pd.DataFrame(
                columns=["category", "0+12", "potential risk", "Now"],
                data=[
                    ["", "e",
                     round((self.budget.Jan.other_project_operations -
                            (self.budget.used.other_project_operations + self.budget.left.other_project_operations))
                           / 1000),
                     ""],
                    ["left", "", "", round(self.budget.left.other_project_operations / 1000)],
                    ["used", "", "", round(self.budget.used.other_project_operations / 1000)],
                ],
            ),
        }
        self.chart.append(chart7)
        for chart in self.chart:
            self.tc.add_chart_from_dataframe(
                template_name=self.template_name,
                chart_name=chart["chart_name"],
                dataframe=chart["dataframe"],
            )

    def save_file(self):
        self.tc.save_ppttc(filename=self.filename)
