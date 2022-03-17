import os
from datetime import datetime
import pandas as pd
import matplotlib.pyplot as plt


def set_password(filepath: str = "wydyf-password.xlsx", prefix: str = "./public/"):
    data = pd.read_excel(filepath)
    for index, row in data.iterrows():
        name = row["1、您的姓名"]
        dir_path = os.path.join(prefix, name)
        os.makedirs(dir_path, exist_ok=True)
        password_txt = open(os.path.join(dir_path, "password.txt"), mode="w")
        password_txt.write(row["2、密码"])


def feedback_volunteer(prefix: str, name: str, data: pd.DataFrame):
    dir_path = os.path.join(prefix, name)
    data = pd.DataFrame(data)
    os.makedirs(name=dir_path, exist_ok=True)
    data.reset_index(drop=True, inplace=True)
    data.index.set_names("序号", inplace=True)
    data.drop(columns=["来源", "来源详情", "来自IP", "您的姓名", "总分"], inplace=True)
    data.to_csv(os.path.join(dir_path, f"{name}.csv"))

    subjects = data[["答疑科目", "服务时长/分钟"]].groupby(by="答疑科目").sum()
    subjects.sort_values(by="服务时长/分钟", ascending=False, inplace=True)
    plt.figure()
    plt.pie(subjects["服务时长/分钟"], labels=subjects.index, autopct="%.1f%%")
    plt.savefig(os.path.join(dir_path, "pie.png"))
    plt.close()

    plt.figure()
    plt.bar(x=subjects.index, height=subjects["服务时长/分钟"])
    plt.ylabel("服务时长/分钟")
    plt.savefig(os.path.join(dir_path, "bar.png"))
    plt.close()

    readme = open(file=os.path.join(dir_path, "README.md"), mode="w")
    lines = []
    lines.append(f"# 👋 Hi, {name}")
    lines.append(f"")
    lines.append(f"Last Updated on: {datetime.now()}")
    lines.append(f"")
    lines.append(f"总答疑 {len(data)} 次 {round( subjects['服务时长/分钟'].sum() / 60, 2)} 小时")
    lines.append(f"")
    rating = data[["评分—业务能力", "启发程度", "服务态度", "满意程度"]]
    rating = rating.mean()
    lines.append(f"- 业务能力: {rating['评分—业务能力']} / 5.0")
    lines.append(f"- 启发程度: {rating['启发程度']} / 5.0")
    lines.append(f"- 服务态度: {rating['服务态度']} / 5.0")
    lines.append(f"- 满意程度: {rating['满意程度']} / 5.0")
    lines.append(f"")
    lines.append(f"![bar](bar.png)")
    lines.append(f"")
    lines.append(f"![pie](pie.png)")
    lines.append(f"")
    readme.writelines([line + "\n" for line in lines])


def feedback(filepath: str = "wydyf-feedback.xlsx", prefix: str = "./public/"):
    plt.rcParams["font.sans-serif"] = ["Microsoft YaHei"]
    plt.rcParams["axes.unicode_minus"] = False

    xlsx = pd.read_excel(filepath, index_col=0)
    for name, data in xlsx.groupby(by="志愿者"):
        feedback_volunteer(prefix=prefix, name=name, data=data)


if __name__ == "__main__":
    set_password()
    feedback()
