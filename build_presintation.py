import matplotlib.pyplot as plt
from typing import List, Tuple, Dict, Any, Union
from collections.abc import Collection
import json
from pptx_renderer import PPTXRenderer


#https://issekdash.hse.ru/viewer/public?dashboardGuid=0e54f8988a3241129ffae56ca76001cd&showNav=true
with open("data.json", "r") as f:
    data = json.load(f)

#https://issekdash.hse.ru/viewer/public?dashboardGuid=0e54f8988a3241129ffae56ca76001cd&showNav=true
with open("data2.json", "r") as f:
    data2 = json.load(f)


def select(selected_year: int, sector: str, data: Dict[str, Any]):
    # can raise ValueError
    indx = list(map(lambda x: x["idParts"][-1], data["colsInfo"])).index(sector)
    for value, year, quarter in zip(data["values"][indx], *zip(*data["rows"])):
        if int(year) == selected_year:
            yield value, quarter


def my_plot(plt, bar_names: List[str], raw_values: List[Tuple[float, float]], title: str = "title", figsize=(20, 10),
            width=0.8, magic_values=(0.1, 0.1, 0), anotate_diff=True, anotate_diff_round=1):
    colors = []
    values = []
    arrow_data = []
    x = 0
    for before_val, after_val in raw_values:
        if before_val > after_val:
            values += [before_val, after_val]
            colors += ["red", "grey"]
        else:
            values += [after_val, before_val]
            colors += ["green", "grey"]
        arrow_data.append({"after_val": after_val, "before_val": before_val, "x": x + magic_values[0], "y": after_val,
                           "dy": after_val - before_val,
                           "defference": f'{round(abs(before_val - after_val) / before_val * 100, 1)}%',
                           "defference_pos": (x - magic_values[1], magic_values[2] + (before_val + after_val) / 2)})
        x += 1
    plt.figure(figsize=figsize)
    plt.title(title)
    double_bar_names = sum(map(list, zip(bar_names, bar_names)), [])
    # print(double_bar_names, values, colors)
    plt.bar(double_bar_names, values, color=colors, width=width)
    for arrow in arrow_data:
        plt.annotate('',
                     xy=(arrow["x"], arrow["y"]),
                     xytext=(arrow["x"], arrow["y"] - arrow["dy"]),
                     arrowprops={'arrowstyle': "->"})
        # plt.arrow(x= h + 0.3, y= data["пассажирооборот"][i][2] , dx= 0 , dy= data["пассажир
        plt.annotate(arrow["defference"], xy=arrow["defference_pos"])
        if anotate_diff:
            plt.annotate(
                f"{round(arrow['before_val'], anotate_diff_round)}->{round(arrow['after_val'], anotate_diff_round)}",
                xy=(arrow["defference_pos"][0], max(arrow["after_val"], arrow["before_val"]) + 0.1))


# indexes = data.keys()
# sectors = set(map(lambda x: x["idParts"][-1], sum(map(lambda x: x["colsInfo"], data.values()), [])))
# my_plot(plt, ['all', 'дальнего следования', 'пригородного сообщения'], [(133.3, 78.1), (99.1, 120.5), (34.4, 24.7)])
def create_my_plot(sector, index, year1, year2, y_lab=None, **kwargs):
    global data
    try:
        valeues1, names1 = zip(*select(year1, sector=sector, data=data[index]))
        valeues2, names2 = zip(*select(year2, sector=sector, data=data[index]))
    except ValueError:
        return f'pictures/error.jpg'
    length = min(len(names1), len(names2))
    valeues1 = valeues1[:length]
    valeues2 = valeues2[:length]
    names = names1[:length]
    plt.clf()
    if "title" not in kwargs:
        kwargs["title"] = f"Динамика изменения {year1}-{year2} показателя «{index}»"
    my_plot(plt, names,
            list(zip(valeues1, valeues2)), **kwargs)
    if y_lab is None:
        plt.ylabel(index.split(",")[-1])
    link = f'pictures/{sector.replace(" ", "_").replace(".", "").replace(",", "")}_{index.replace(" ", "_")}_{year1}_{year2}.png'
    plt.savefig(link)
    plt.close()
    return link


def avg(smth: Collection[Union[float, int]]):
    smth_sum = 0
    for i in smth:
        smth_sum += i
    return smth_sum / len(smth)


def get_avg_index_by_sector(year, sector, index):
    data_to_calc, _ = zip(*select(year, sector=sector, data=data[index]))
    print(data_to_calc)
    return avg(data_to_calc)


def process_data(data: Dict[str,Any]):
    values = sum(data['values'], [])
    i = 0
    for year in list(map(lambda x: x["idParts"][-1], data["colsInfo"])):
        for month in data['rows']:
            month, *_ = month
            yield {"value": values[i], "year": year, "month": month}
            i += 1
            if i == len(values):
                break


def create_my_plot2(sector: str, year: int = None, group: int = 1, month: str = None, title: str = None,
                    figsize=(20, 10)):
    p_data = list(process_data(data2[sector]))
    if year is not None:
        p_data = list(filter(lambda x: int(x["year"]) == year, p_data))
    if month is not None:
        p_data = list(filter(lambda x: x["month"] == month, p_data))
    if group != 1:
        x = []
        y = []
        for i in range(0, len(p_data), group):
            group_data = p_data[i:i + group]
            if len(group_data) < group:
                continue
            if group_data[0]['year'] != group_data[-1]['year']:
                year_dif = group_data[0]['year'] + "-" + group_data[-1]['year']
            else:
                year_dif = group_data[0]['year']
            x += [f"{group_data[0]['month']}-{group_data[-1]['month']}\n{year_dif}"]
            y += [sum(map(lambda x: x["value"], group_data))]
    else:
        x = list(map(lambda x: x["month"] + "\n" + x["year"], p_data))
        y = list(map(lambda x: x["value"], p_data))
    plt.figure(figsize=figsize)
    plt.ylim(0, 300)
    plt.plot(x, y)
    plt.ylabel("млрд руб")
    if title is None:
        title = f"Динамика изменения оборота cетора «{sector}»"
    plt.title(title)
    filename = sector.replace(" ", "_").replace(".", "").replace(",", "") + str(group) + (
        str(year) if year is not None else "") + (str(month) if month is not None else "")
    link = f'pictures/{filename}.png'
    plt.savefig(link)
    plt.close()
    return link


def avg_round_by_sector(sector, round_num=2):
    return round(avg(list(map(lambda x: x["value"], list(process_data(data2[sector]))))), round_num)



p = PPTXRenderer("Tamplate.pptx")
p.render(
    "output.pptx",
    {
        "create_my_plot": create_my_plot,
        "data": data,
        "avg":avg,
        "avg_round_by_sector":avg_round_by_sector,
        "create_my_plot2":create_my_plot2
    }
)