import os
import shutil
from datetime import datetime

import pandas as pd
import matplotlib.pyplot as plt


def set_password(
    filepath: str = "wydyf-password.xlsx", prefix: str = "./public/"
) -> None:
    data = pd.read_excel(filepath)
    for index, row in data.iterrows():
        name = row["1、您的姓名"]
        dir_path = os.path.join(prefix, name)
        os.makedirs(dir_path, exist_ok=True)
        password_txt = open(os.path.join(dir_path, ".password"), mode="w")
        password_txt.write(row["2、密码"])


def plot_data(data: pd.DataFrame, dir_path: str = "./") -> pd.DataFrame:
    subjects = data[["答疑科目", "服务时长/分钟"]].groupby(by="答疑科目").sum()
    subjects.sort_values(by="服务时长/分钟", ascending=False, inplace=True)
    n: int = 9
    plot_data: pd.Series = subjects["服务时长/分钟"].head(n=n)
    plot_data["其他"] = subjects["服务时长/分钟"][n:].sum()
    plt.figure(dpi=600)
    plt.pie(plot_data, labels=plot_data.index, autopct="%.1f%%")
    plt.tight_layout()
    plt.savefig(os.path.join(dir_path, "pie.png"))
    plt.close()

    plt.figure(dpi=600)
    plt.barh(y=plot_data.index, width=plot_data)
    plt.xlabel("服务时长/分钟")
    plt.tight_layout()
    plt.savefig(os.path.join(dir_path, "bar.png"))
    plt.close()

    return subjects


def append_statistics(
    data: pd.DataFrame, dir_path: str = "./", num_comments: int = 5
) -> list[str]:
    subjects = plot_data(data=data, dir_path=dir_path)

    lines = []
    lines.append(f"Last Updated on: {datetime.now()}")
    lines.append(f"")
    lines.append(f"总答疑 {len(data)} 次 {round( subjects['服务时长/分钟'].sum() / 60, 2)} 小时")
    lines.append(f"")
    rating = data[["评分—业务能力", "启发程度", "服务态度", "满意程度"]].copy(deep=True)
    rating.replace(to_replace="(空)", value=5, inplace=True)
    rating = rating.astype(dtype=int)
    rating = rating.mean()
    lines.append(f"- 业务能力: {round(rating['评分—业务能力'], 2)} / 5.0")
    lines.append(f"- 启发程度: {round(rating['启发程度'], 2)} / 5.0")
    lines.append(f"- 服务态度: {round(rating['服务态度'], 2)} / 5.0")
    lines.append(f"- 满意程度: {round(rating['满意程度'], 2)} / 5.0")
    lines.append(f"")
    lines.append(f"![bar](bar.png)")
    lines.append(f"")
    lines.append(f"![pie](pie.png)")
    lines.append(f"")
    comments = data["想说的话"]
    comments = comments[comments != "(空)"]
    comments.sort_index(ascending=False, inplace=True)
    lines.append(f"## Latest {num_comments} Comments")
    lines.append(f"")
    for comment in comments[:num_comments]:
        lines.append(f"- {comment}")
        lines.append(f"")
    return lines


def feedback_volunteer(prefix: str, name: str, data: pd.DataFrame) -> None:
    dir_path = os.path.join(prefix, name)
    data = pd.DataFrame(data)
    os.makedirs(name=dir_path, exist_ok=True)
    data.reset_index(drop=True, inplace=True)
    data.index.set_names("序号", inplace=True)
    data.to_csv(os.path.join(dir_path, f"{name}.csv"))

    readme = open(file=os.path.join(dir_path, "README.md"), mode="w")
    lines = []
    lines.append(f"# 👋 Hi, {name}")
    lines.append(f"")
    lines += append_statistics(data=data, dir_path=dir_path)
    readme.writelines([line + "\n" for line in lines])


def feedback(filepath: str = "wydyf-feedback.xlsx", prefix: str = "./public/") -> None:
    plt.rcParams["font.sans-serif"] = ["Microsoft YaHei"]
    plt.rcParams["axes.unicode_minus"] = False

    all_data = pd.read_excel(filepath, index_col=0)
    all_data.drop(columns=["来源", "来源详情", "来自IP", "您的姓名", "总分"], inplace=True)

    lines = []
    lines.append(f"## 欢迎加入 **未央书院答疑坊**!")
    lines.append(f"")
    lines.append(f"## Statistics")
    lines.append(f"")
    lines += append_statistics(data=all_data, dir_path=prefix)
    lines.append(f"## Contribute")
    lines.append(f"")
    lines.append(f"想要加入更多功能? 想要修复 Bug?")
    lines.append(f"")
    lines.append(
        f"欢迎在 [liblaf/thu-wy-dyf](https://github.com/liblaf/thu-wy-dyf) 发起 Issue / Pull Request! 我们参考 [Conventional Commits](https://www.conventionalcommits.org/en/v1.0.0/) 提交 commits."
    )
    lines.append(f"")
    os.makedirs(name=prefix, exist_ok=True)
    readme_md = open(os.path.join(prefix, "README.md"), mode="w")
    readme_md.writelines([line + "\n" for line in lines])

    for name, data in all_data.groupby(by="志愿者"):
        feedback_volunteer(prefix=prefix, name=name, data=data)


def calc_volunteer_hours(
    filepath: str = "wydyf-feedback.xlsx",
    start_time: datetime = datetime(year=2022, month=1, day=1),
    end_time: datetime = datetime(year=2077, month=1, day=1),
    prefix: str = ".",
):
    all_data = pd.read_excel(filepath, index_col=0)
    all_data.drop(columns=["来源", "来源详情", "来自IP", "您的姓名", "总分"], inplace=True)
    all_data["提交答卷时间"] = pd.to_datetime(all_data["提交答卷时间"])
    all_data = all_data[all_data["提交答卷时间"] >= start_time]
    all_data = all_data[all_data["提交答卷时间"] < end_time]
    # print(all_data["提交答卷时间"])
    records: list[dict] = []
    for name, data in all_data.groupby(by="志愿者"):
        records.append({"name": name, "time": data["服务时长/分钟"].sum()})
    records: pd.DataFrame = pd.DataFrame.from_dict(records)
    records.to_excel(os.path.join(prefix, "volunteer-hours.xlsx"))
    # print(pd.DataFrame.from_dict(records))


if __name__ == "__main__":
    PREFIX = "./public/"
    calc_volunteer_hours(
        start_time=datetime(year=2022, month=4, day=25),
        end_time=datetime(year=2022, month=7, day=17),
    )
    shutil.rmtree(PREFIX, ignore_errors=True)
    set_password(prefix=PREFIX)
    feedback(prefix=PREFIX)
