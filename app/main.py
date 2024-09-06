from statistics import stdev
from typing import AnyStr

import pandas as pd
import matplotlib.pyplot as plt
from numpy import isnan

from pandas import DataFrame
from os import scandir


def oldest_file(objects) -> AnyStr:
    edit_time = 0
    filename: AnyStr | None = None
    for obj in objects:
        if obj.is_file() and ".xlsx" in obj.name and '~' not in obj.name and edit_time < obj.stat().st_mtime:
            filename = obj.path
            edit_time = obj.stat().st_mtime
    return filename


def get_data_frame() -> DataFrame | None:
    """
    Возвращает DataFrame созданый из xlsx в папке Input
    :return: DataFrame созданый из xlsx в папке Input
    """
    if any(scandir('..//Input//')):
        with scandir('..//Input//') as entries:
            return pd.read_excel(oldest_file(entries))
    return None


def read_df(df: DataFrame) -> dict[str, list[float]]:
    data: dict[str, list[float]] = {}
    last_key: str = ""
    for i, row in df.iterrows():
        if isinstance(row["Метод экстракции"], str):
            last_key = row["Метод экстракции"]
            data[last_key]: list[float] = []
        if isnan(row["Выход ДНК: ДНК(нг)*Vэллюции(мкл)"]):
            continue
        data[last_key].append(row["Выход ДНК: ДНК(нг)*Vэллюции(мкл)"])
    return data


def work_with_data(dt: dict[str, list[float]]) -> list[list[str | float]]:
    res: list[list[str | float]] = [[], [], []]
    for key in dt.keys():
        length: int = len(dt[key])
        if length == 0:
            continue
        average: float = sum(dt[key]) / length
        deviation: float = stdev(dt[key])
        res[0].append(key)
        res[1].append(float(average))
        res[2].append(float(deviation))
    return res


def make_plot(data: list[list[str | float]]) -> None:
    plt.bar(data[0], data[1], label="Среднее значение")
    plt.errorbar(data[0], data[1], yerr=data[2], fmt="o", color="k")
    plt.xlabel("Метод экстракции")
    plt.ylabel("ДНК(нг) * V эллюции(мкл)")
    plt.title("Сравнение методов экстракции")
    plt.legend()
    plt.savefig(f'..//Output//Сравнение методов экстракции.png')
    plt.show()


def main() -> None:
    make_plot(work_with_data(read_df(get_data_frame())))


if __name__ == '__main__':
    main()

