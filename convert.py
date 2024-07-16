import yaml
import pandas as pd

with open("index.yaml", "r") as file:
    yaml_data = yaml.safe_load(file)

paths = {}


def flatten_dict(d, parent_key="", sep="."):
    for key, value in d.items():
        if not isinstance(value, dict):
            if not paths.get(parent_key):
                paths[parent_key] = {}
                paths[parent_key][key] = value

            paths[parent_key][key] = value
            continue
        new_key = f"{parent_key}{sep}{key}" if parent_key else key
        flatten_dict(value, new_key, ".")


with pd.ExcelWriter("output.xlsx", engine="openpyxl") as writer:
    for sheet, data in yaml_data.items():
        dataset = flatten_dict(data)
        data_list = [{"Path": key, **value} for key, value in paths.items()]

        df = pd.DataFrame(data_list)

        columns_order = ["Path", "required", "type", "owner", "usage"]
        df = df[columns_order]
        df.to_excel(writer, sheet_name=sheet, index=False)
