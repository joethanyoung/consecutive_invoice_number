# import required packages
import os
import pandas as pd
from tqdm import tqdm

# set working directory to current directory
current_directory = os.getcwd()
input_files = [file for file in os.listdir(current_directory) if file.endswith(".xlsx")]

# set up excluded values according to specific acccounting entries
exclude_values = [
  "业务及管理费-差旅费", "应付手续费及佣金-手续费支出", "应付手续费及佣金-手续费支出-保全业务费",
  "佣金支出-间接佣金-其他", "业务及管理费-绿化费", "业务及管理费-人力成本-中转", "待摊费用-中转",
  "应交税费-代扣代缴税金-员工个人所得税", "业务及管理费-车船使用费", "其他应付款-LMS中转", "业务及管理费-水电费",
  "应付职工薪酬-短期薪酬-工会经费", "业务及管理费-邮电费", "业务及管理费-职工工资-短期薪酬-工会经费",
  "业务及管理费-防预费", "应交税费-应交增值税未交-未交增值税", "应交税费-税金及附加-教育费附加",
  "应交税费-代扣代缴税金-代理人个人所得税", "应交税费-应交增值税代扣代缴-代理人", "应交税费-税金及附加-地方教育费附加",
  "税金及附加-印花税", "应付手续费及佣金-手续费支出-出单费", "其他应收款-PS-社会统筹保险",
  "应交税费-税金及附加-城建税", "应交税费-税金及附加-其他附加", "业务及管理费-职工工资-短期薪酬-职工福利费-员工福利费",
  "业务及管理费-学会会费", "应付手续费及佣金-佣金支出-银保代理佣金", "业务及管理费-银行结算费",
  "其他应收款-PS-工会经费", "业务及管理费-企业财产保险费", "业务及管理费-职工工资-短期薪酬-外勤提奖",
  "应付手续费及佣金-佣金支出-收展代理佣金", "业务及管理费-安全防卫费", "税金及附加-车船使用税",
  "长期待摊费用-中转", "业务及管理费-修理费", "业务及管理费-诉讼费", "业务及管理费-劳动保险费",
  "固定资产-中转", "业务及管理费-职工工资-短期薪酬-职工福利费-员工福利计划", "应付手续费及佣金-佣金支出-个险代理佣金",
  "应交税费-代扣代缴税金-代理人税金及附加", "业务及管理费-咨询费-咨询费", "应交税费-代扣代缴税金-其他代扣代缴",
  "业务及管理费-职工工资-短期薪酬-临时工工资", "营业外支出-其他", "营业外支出-罚没款项",
  "税金及附加-土地使用税", "业务及管理费-同业工会会费", "税金及附加-房产税"
]

# set a loop for processing files including read files and detect str of '发票号码', convert its to numeric and set the difference
for input_file_path in tqdm(input_files, desc="Processing files", unit="file"):
  df = pd.read_excel(input_file_path, skiprows=1)
if '发票号码' not in df.columns:
  print(f"{input_file_path} does not contain '发票号码', skipping...")
continue
df = df[~df['会计科目'].isin(exclude_values)]
df = df[
  (df['发票号码'].astype(str).str.len() == 8) & (df['发票号码'].astype(str).str.isnumeric())
].sort_values(by='发票号码')
df['diff'] = df['发票号码'].astype(int).diff().fillna(0).astype(int)
df['group'] = (df['diff'] > 1).cumsum()
consecutive_groups = df[df.groupby('group')['发票号码'].transform('count') > 1]

# Filter out rows with consecutive invoice numbers and the same 业务单号
con_invoice = consecutive_groups[~consecutive_groups['收款人'].astype(str).str.len().isin([2, 3])]

# Group by 'group', and keep groups where at least one 业务单号 is different
filtered_con_invoice = con_invoice.groupby('group').filter(lambda x: x['业务单号'].nunique() > 1)

# Print the filtered results
print(filtered_con_invoice)

# Save the filtered results to an Excel file
input_filename, _ = os.path.splitext(os.path.basename(input_file_path))
output_filename = f"{input_filename}_连号发票.xlsx"
filtered_con_invoice.to_excel(output_filename, index=False)