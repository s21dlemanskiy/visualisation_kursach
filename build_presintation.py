import matplotlib.pyplot as plt
from typing import List, Tuple, Dict, Any, Union
from collections.abc import Collection
import json
from pptx_renderer import PPTXRenderer
from numpy import polyfit, poly1d


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
            width=0.8, magic_values=(0.34, 0.42, 0), anotate_diff_percent=True, anotate_diff_values=True, anotate_diff_round=1,
            general_font_size=10, axes_font_size=10, axis_title_font_size=10, title_fontsize=10, y_lab=None, y_bounds=None):
    """define some global size vars"""
    plt.rc('font', size=general_font_size)  # controls default text sizes
    plt.rc('axes', titlesize=axis_title_font_size)  # fontsize of the axes title
    plt.rc('axes', labelsize=axes_font_size)  # fontsize of the x and y labels
    # plt.rc('xtick', labelsize=size)  # fontsize of the tick labels
    # plt.rc('ytick', labelsize=size)  # fontsize of the tick labels
    # plt.rc('legend', fontsize=size)  # legend fontsize
    plt.rc('figure', titlesize=title_fontsize)  # fontsize of the figure title
    """calculate and define some important values"""
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
        arrow_data.append({"after_val": after_val, "before_val": before_val,
                           "x": x + width * magic_values[0], "y": after_val,
                           "dy": after_val - before_val,
                           "defference": f'{"-" if before_val > after_val else "+"}{round(abs(before_val - after_val) / before_val * 100, 1)}%',
                           "defference_pos": (x - width * magic_values[1], magic_values[2] * max(after_val, before_val) + (before_val + after_val) / 2),
                           "color": colors[-2]})
        x += 1
    """style"""
    plt.ylabel(y_lab)
    plt.figure(figsize=figsize)
    plt.title(title, fontsize=title_fontsize)
    double_bar_names = sum(map(list, zip(bar_names, bar_names)), [])
    if y_bounds is not None:
        plt.ylim(*y_bounds)
    """ploting"""
    plt.bar(double_bar_names, values, color=colors, width=width)
    """add arrow and some text"""
    for arrow in arrow_data:
        """arrow"""
        plt.annotate('',
                     xy=(arrow["x"], arrow["y"]),
                     xytext=(arrow["x"], arrow["y"] - arrow["dy"]),
                     arrowprops={'arrowstyle': "->"})
        """percent"""
        if anotate_diff_percent:
            percent_anotaton = plt.annotate(arrow["defference"], xy=arrow["defference_pos"])
            percent_anotaton.set_bbox(dict(facecolor='white', alpha=0.7, edgecolor=arrow["color"], boxstyle='round', pad=0.3))
        """plt val was to val become"""
        if anotate_diff_values:
            plt.annotate(
                f"{round(arrow['before_val'], anotate_diff_round)}->{round(arrow['after_val'], anotate_diff_round)}",
                xy=(arrow["defference_pos"][0], max(arrow["after_val"], arrow["before_val"]) + 0.1))
    if y_bounds is not None:
        plt.ylim(*y_bounds)


# indexes = data.keys()
# sectors = set(map(lambda x: x["idParts"][-1], sum(map(lambda x: x["colsInfo"], data.values()), [])))
# my_plot(plt, ['all', 'дальнего следования', 'пригородного сообщения'], [(133.3, 78.1), (99.1, 120.5), (34.4, 24.7)])
def create_my_plot(sector, index, year1, year2, y_lab=None, **kwargs):
    # varible [data] from other scope
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
    if y_lab is None:
        y_lab = index.split(",")[-1]
    """ploting"""
    my_plot(plt, names,
            list(zip(valeues1, valeues2)), y_lab=y_lab, **kwargs)
    """save plot to png local file"""
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


def add_trend_to_plot(plt, x, y, *args, **kwargs):
    to_float = lambda arr: [i for i in range(0, len(arr))]
    z = polyfit(to_float(x), y, 1)
    p = poly1d(z)
    plt.plot(x, p(to_float(x)), **kwargs)

def create_my_plot2(sector: str, year: int = None, group: int = 1, month: str = None, title: str = None,
                    figsize=(20, 10), general_font_size=10, axes_font_size=10, axis_title_font_size=10, title_fontsize=10,
                    y_bounds=(0, 500), trend=True, y_lab=None):
    # varible [data2] from other scope
    """define some global size vars"""
    plt.rc('font', size=general_font_size)  # controls default text sizes
    plt.rc('axes', titlesize=axis_title_font_size)  # fontsize of the axes title
    plt.rc('axes', labelsize=axes_font_size)  # fontsize of the x and y labels
    # plt.rc('xtick', labelsize=size)  # fontsize of the tick labels
    # plt.rc('ytick', labelsize=size)  # fontsize of the tick labels
    # plt.rc('legend', fontsize=size)  # legend fontsize
    plt.rc('figure', titlesize=title_fontsize)  # fontsize of the figure title
    """some work with data"""
    p_data = list(process_data(data2[sector]))
    if year is not None:
        p_data = list(filter(lambda x: int(x["year"]) == year, p_data))
    if month is not None:
        p_data = list(filter(lambda x: x["month"] == month, p_data))

    """if we want to group by n month - do it"""
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
    """plot figure"""
    plt.figure(figsize=figsize)
    if trend:
        add_trend_to_plot(plt, x, y, color="grey", linestyle="--")
    plt.plot(x, y)
    """ add some styling """
    plt.ylim(*y_bounds)
    if y_lab is None:
        y_lab = "млрд руб"
    plt.ylabel(y_lab)
    if title is None:
        title = f"Динамика изменения оборота cектора «{sector}»"
    plt.title(title, fontsize=title_fontsize)
    """save plot to png local file"""
    filename = sector.replace(" ", "_").replace(".", "").replace(",", "") + str(group) + (
        str(year) if year is not None else "") + (str(month) if month is not None else "")
    link = f'pictures/{filename}.png'
    plt.savefig(link)
    plt.close()
    return link


def create_my_plot3(sector_list: str, title: str = None,
                    figsize=(20, 10), general_font_size=10, axes_font_size=10, axis_title_font_size=10, title_fontsize=10,
                    y_bounds=(0, 500), y_lab=None):
    # varible [data2] from other scope
    """define some global size vars"""
    plt.rc('font', size=general_font_size)  # controls default text sizes
    plt.rc('axes', titlesize=axis_title_font_size)  # fontsize of the axes title
    plt.rc('axes', labelsize=axes_font_size)  # fontsize of the x and y labels
    plt.rc('figure', titlesize=title_fontsize)  # fontsize of the figure title
    plt.figure(figsize=figsize)
    for sector in sector_list:
        """some work with data"""
        p_data = list(process_data(data2[sector]))
        x = list(map(lambda x: x["month"] + "\n" + x["year"], p_data))
        y = list(map(lambda x: x["value"], p_data))
        """plot figure"""
        plt.plot(x, y, label=sector)
    plt.legend(loc="upper left")
    """ add some styling """
    plt.ylim(*y_bounds)
    if y_lab is None:
        y_lab = "млрд руб"
    plt.ylabel(y_lab)
    if title is None:
        title = f"Динамика изменения оборота cекторов ИКТ"
    plt.title(title, fontsize=title_fontsize)
    """save plot to png local file"""
    sector_format = lambda sector: sector.replace(" ", "_").replace(".", "").replace(",", "")
    filename = "my_plot3" + "_".join(list(map(sector_format, sector_list)))
    link = f'pictures/{filename}.png'
    plt.savefig(link)
    plt.close()
    return link


def avg_round_by_sector(sector, round_num=2):
    return round(avg(list(map(lambda x: x["value"], list(process_data(data2[sector]))))), round_num)


SMALL_FONT = 8
MEDIUM_FONT = 10
BIGGER_FONT = 12

p = PPTXRenderer("Tamplate.pptx")
p.render(
    "output.pptx",
    {
        "create_my_plot": create_my_plot,
        "create_my_plot3":create_my_plot3,
        "data": data,
        "avg":avg,
        "avg_round_by_sector":avg_round_by_sector,
        "create_my_plot2":create_my_plot2,
        "SMALL_FONT": SMALL_FONT,
        "MEDIUM_FONT":MEDIUM_FONT,
        "BIGGER_FONT":BIGGER_FONT,
    }
)


