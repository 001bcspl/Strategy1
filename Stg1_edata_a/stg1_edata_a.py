import pandas as pd
from datetime import datetime
import os

INPUT_FILE = "edata_a.xlsx"
OUTPUT_FILE = "logs/stg1_edata_a.xlsx"
EACH_BUNDLE = 35
AMOUNT_INVESTED = 5000

os.makedirs("logs", exist_ok=True)

df = pd.read_excel(INPUT_FILE, dtype={"DATE": str, "TIME": str})
df["DATETIME"] = pd.to_datetime(df["DATE"] + " " + df["TIME"], format="%Y%m%d %H:%M:%S")

start_time = datetime.strptime("09:30", "%H:%M").time()
end_time = datetime.strptime("15:30", "%H:%M").time()

df = df[(df["DATETIME"].dt.time >= start_time) & (df["DATETIME"].dt.time <= end_time)].reset_index(drop=True)

df["LTPCE"] = df["LTPCE"].astype(float)
df["LTPPE"] = df["LTPPE"].astype(float)

crossover_price = 0
crossover_detect = 0
total_bundles = 0
sl_price = 0
target = 0
loss_when_sl_hit = 0
sl_value = 0
profit_at_target = 0

log_rows_for_excel = []

log_header_format = "{:<8} {:<11} {:<8} {:<8} {:<6} {:<10}"
log_row_format_num = "{:<8} {:<11} {:<8.1f} {:<8.1f} {:<6} {:<10.1f}"
log_row_format_str = "{:<8} {:<11} {:<8} {:<8} {:<6} {:<10}"

header_line = log_header_format.format("TIME", "EVENT", "BP", "SL", "QTY", "PL")
print(header_line)
log_rows_for_excel.append(["TIME", "EVENT", "BP", "SL", "QTY", "PL"])


for i in range(len(df)):
    time = df.loc[i, "TIME"]
    ce_ltp = df.loc[i, "LTPCE"]
    pe_ltp = df.loc[i, "LTPPE"]

    current_event = "NONE"
    bp_output = "N"
    sl_output = "N"
    qty_output = "N"
    pl_output = "N"

    if i > 0:
        ce_prev = df.loc[i - 1, "LTPCE"]
        pe_prev = df.loc[i - 1, "LTPPE"]

        if ce_prev < pe_prev and ce_ltp > pe_ltp and crossover_detect == 0:
            crossover_detect = 1
            crossover_price = ce_ltp
            sl_price = round(df.loc[max(0, i - 5):i, "LTPCE"].min(), 2)
            sl_value = round(crossover_price - sl_price, 2)

            if sl_value <= 0:
                current_event = "SL_ERR"
                excel_row_data = [time, current_event, "N", "N", "N", "N"]
                console_row = log_row_format_str.format(time, current_event, "N", "N", "N", "N")
                print(console_row)
                log_rows_for_excel.append(excel_row_data)
                continue

            target = round(crossover_price + 2 * sl_value, 2)
            total_bundles = round(AMOUNT_INVESTED / (EACH_BUNDLE * sl_value))
            if total_bundles == 0:
                total_bundles = 1

            loss_when_sl_hit = round((sl_price - crossover_price) * EACH_BUNDLE * total_bundles, 2)
            profit_at_target = round((target - crossover_price) * EACH_BUNDLE * total_bundles, 2)

            current_event = "CROSSOVER"
            bp_output = crossover_price
            sl_output = sl_price
            qty_output = total_bundles
            pl_output = 0.0

        elif crossover_detect == 1:
            runt_pl = round((ce_ltp - crossover_price) * EACH_BUNDLE * total_bundles, 2)
            current_event = "CROSSOVER"
            bp_output = ce_ltp
            sl_output = sl_price
            qty_output = total_bundles
            pl_output = runt_pl

    if isinstance(bp_output, (int, float)):
        formatted_bp = f"{bp_output:.1f}"
    else:
        formatted_bp = str(bp_output)

    if isinstance(sl_output, (int, float)):
        formatted_sl = f"{sl_output:.1f}"
    else:
        formatted_sl = str(sl_output)

    if isinstance(pl_output, (int, float)):
        formatted_pl = f"{pl_output:.1f}"
    else:
        formatted_pl = str(pl_output)
   
    console_row = log_row_format_str.format(time, current_event, formatted_bp, formatted_sl, str(qty_output), formatted_pl)
    print(console_row)

    log_rows_for_excel.append([time, current_event, bp_output, sl_output, qty_output, pl_output])

summary_text = "\n--- Crossover Summary ---\n"
summary_data_for_excel = []

if crossover_detect == 1:
    summary_text += f"""Cross Over: {crossover_price:.1f}
SL: {sl_price:.1f}
Target: {target:.1f}
Amount Invested: {AMOUNT_INVESTED}
Each Bundle: {EACH_BUNDLE}
Total Bundles: {total_bundles}
Loss When SL Hit: {loss_when_sl_hit:.2f}
SL Value: {sl_value:.2f}
PL When Target Hit: {profit_at_target:.2f}
"""
    # Prepare summary data for Excel (raw values)
    summary_data_for_excel = [
        ["", "", "", "", "", ""], # Blank row for separation
        ["--- Crossover Summary ---", "", "", "", "", ""],
        ["Cross Over", crossover_price, "", "", "", ""],
        ["SL", sl_price, "", "", "", ""],
        ["Target", target, "", "", "", ""],
        ["Amount Invested", AMOUNT_INVESTED, "", "", "", ""],
        ["Each Bundle", EACH_BUNDLE, "", "", "", ""],
        ["Total Bundles", total_bundles, "", "", "", ""],
        ["Loss When SL Hit", loss_when_sl_hit, "", "", "", ""],
        ["SL Value", sl_value, "", "", "", ""],
        ["PL When Target Hit", profit_at_target, "", "", "", ""],
    ]
else:
    summary_text += "No crossover events found.\n"
    summary_data_for_excel = [
        ["", "", "", "", "", ""],
        ["--- Crossover Summary ---", "", "", "", "", ""],
        ["No crossover events found", "", "", "", "", ""]
    ]

print(summary_text)

output_df = pd.DataFrame(log_rows_for_excel[1:], columns=log_rows_for_excel[0])

summary_df = pd.DataFrame(summary_data_for_excel, columns=log_rows_for_excel[0])
output_df = pd.concat([output_df, summary_df], ignore_index=True)
output_df.to_excel(OUTPUT_FILE, index=False)
