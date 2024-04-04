from pathlib import Path  # Standard Python Module
import xlwings as xw  # pip install xlwings
import plotly.express as px  # pip install plotly-express
import plotly  # pip install plotly
import pandas as pd  # pip install pands


def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    df = sheet.range("A1").expand().options(pd.DataFrame, index=False).value

    tasks = df["Task"]
    start = df["Start"]
    finish = df["Finish"]
    complete = df["Complete in %"]

    fig = px.timeline(
        df,
        x_start=start,
        x_end=finish,
        y=tasks,
        color=complete,
        title="Task Overview"
    )
    file_path = str(Path(__file__).parent / "Task_Overview.html")
    plotly.offline.plot(fig, filename=file_path, auto_open=False)
